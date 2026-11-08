"""Owner-scoped, rebuildable library search; workbook indexes remain authoritative.

Generation files are immutable. A writer publishes CURRENT with os.replace;
readers pin a filename until their query ends. No workbook databases are attached.
The host's workspace lock serializes source replacement with catalog operations.
"""

from __future__ import annotations

import base64
import hashlib
import json
import math
import re
import sqlite3
import threading
import time
import unicodedata
import uuid
from collections import OrderedDict
from collections.abc import Iterable
from contextlib import contextmanager, nullcontext
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
from urllib.parse import quote

from linexcel.catalog_build import WorkbookBuilder, schema
from linexcel.lazy_store import WorkbookIndex, graph_neighborhood
from linexcel.limits import validate_limits

TRACE_VERSION = "linexcel.search.v1"
CATALOG_VERSION = 4
KINDS = frozenset({"cell", "sheet", "name", "table", "column"})


def _json(value) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), allow_nan=False)


def normalize(text: str) -> str:
    """Stable Unicode case/diacritic folding shared by FTS and LIKE fallback."""
    return "".join(
        c
        for c in unicodedata.normalize("NFD", text.casefold())
        if unicodedata.category(c) != "Mn"
    )


def terms_for(text: str) -> list[str]:
    """Treat underscores as separators, as the unicode61 tokenizer does."""
    return re.findall(r"[^\W_]+", normalize(text))


def fts5_available() -> bool:
    with sqlite3.connect(":memory:") as db:
        try:
            db.execute("CREATE VIRTUAL TABLE probe USING fts5(text)")
            return True
        except sqlite3.OperationalError:
            return False


@dataclass(frozen=True)
class CatalogLimits:
    entries_per_workbook: int = 100_000
    numeric_cells_per_workbook: int = 100_000
    text_occurrences: int = 8
    header_scan_cells: int = 50_000
    header_candidate_rows: int = 512
    header_columns: int = 256
    header_bands: int = 64
    header_row_labels: int = 4
    header_regions: int = 256
    catalog_bytes: int = 256 * 1024 * 1024
    build_seconds: int = 120
    connections: int = 4
    hits: int = 500
    response_bytes: int = 1024 * 1024
    scan_entries: int = 50_000
    query_ms: int = 2000
    query_chars: int = 512
    text_chars: int = 4096
    lineage_nodes: int = 200
    lineage_depth: int = 20
    lineage_edges: int = 1000
    lineage_workbooks: int = 32
    lineage_ms: int = 2000
    lineage_links: int = 100_000

    def __post_init__(self):
        validate_limits(**asdict(self))
        if self.connections < 1 or self.response_bytes < 4096:
            raise ValueError("connections must be >=1 and response_bytes >=4096")


def revision(project_id: str, sha256: str) -> dict:
    return {"id": f"{project_id}:{sha256}", "projectId": project_id, "sha256": sha256}


def cell_handle(rev: dict, node_id: str) -> str:
    data = _json([rev["projectId"], rev["sha256"], node_id]).encode()
    return "lx1." + base64.urlsafe_b64encode(data).decode().rstrip("=")


def parse_handle(handle: str) -> tuple[str, str, str]:
    try:
        if not isinstance(handle, str) or not handle.startswith("lx1."):
            raise ValueError
        if len(handle) > 4096:
            raise ValueError
        raw = handle[4:]
        parts = json.loads(
            base64.b64decode(raw + "=" * (-len(raw) % 4), altchars=b"-_", validate=True)
        )
        if (
            not isinstance(parts, list)
            or len(parts) != 3
            or not all(isinstance(p, str) for p in parts)
            or not re.fullmatch(r"[a-f0-9]{32}", parts[0])
            or not re.fullmatch(r"[a-f0-9]{64}", parts[1])
            or len(parts[2]) > 2048
        ):
            raise ValueError
        return tuple(parts)
    except (ValueError, TypeError, UnicodeError, RecursionError):
        raise ValueError("Invalid cell handle") from None


class IndexPool:
    """Operation-local LRU: never shares SQLite objects across request threads."""

    def __init__(self, budget: int):
        self.budget = budget
        self.indexes: OrderedDict[Path, WorkbookIndex] = OrderedDict()
        self.peak = 0

    def get(self, path: Path) -> WorkbookIndex:
        if path in self.indexes:
            self.indexes.move_to_end(path)
            return self.indexes[path]
        while len(self.indexes) >= self.budget:
            _, index = self.indexes.popitem(last=False)
            index.close()
        self.indexes[path] = index = WorkbookIndex(path)
        self.peak = max(self.peak, len(self.indexes))
        return index

    def close(self):
        for index in self.indexes.values():
            index.close()
        self.indexes.clear()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()


@dataclass(frozen=True)
class Source:
    project_id: str
    filename: str
    path: Path
    imported_at: float = 0


class Catalog:
    def __init__(self, directory: Path, limits: CatalogLimits | None = None):
        self.directory = directory
        self.directory.mkdir(parents=True, exist_ok=True)
        self.limits = limits or CatalogLimits()
        self.fts5 = fts5_available()  # Check at service startup, not first query.
        self.writer = threading.RLock()
        # Workbook connections share one budget even across concurrent readers.
        self.source_reads = threading.RLock()
        self.readers = threading.RLock()
        self.pins: dict[str, int] = {}
        self.protected: set[str] = set()
        self.build_state: dict = {"status": "idle"}
        self.failed_manifest: list[dict] = []
        self.active_manifest: list[dict] = []

    def _current(self) -> dict | None:
        try:
            current = json.loads((self.directory / "CURRENT").read_text("utf-8"))
            if not re.fullmatch(r"[a-f0-9]{32}", current["generation"]):
                raise ValueError("Invalid catalog generation")
            return current
        except (OSError, ValueError, KeyError, TypeError):
            return None

    @contextmanager
    def pin(self):
        db = None
        with self.readers:
            current = self._current()
            if current is None:
                raise ValueError("Catalog has not been built")
            key = current["generation"]
            self.pins[key] = self.pins.get(key, 0) + 1
        try:
            db = sqlite3.connect(
                "file:"
                + quote(str((self.directory / f"{key}.sqlite").resolve()), safe="/:")
                + "?mode=ro",
                uri=True,
            )
            yield db, current
        finally:
            if db is not None:
                db.close()
            with self.readers:
                self.pins[key] -= 1
                if not self.pins[key]:
                    del self.pins[key]
                self._collect()

    def _published(self, current):
        """Check both artifacts before trusting a pointer's matching signature."""
        if not current:
            return False
        key = current["generation"]
        try:
            manifest = json.loads((self.directory / f"{key}.json").read_text("utf-8"))
            if (
                manifest["signature"] != current["signature"]
                or manifest["generation"] != key
                or manifest["version"] != CATALOG_VERSION
            ):
                return False
            with self.pin() as (db, _):
                db.execute("SELECT id FROM entries LIMIT 1").fetchone()
            return True
        except (OSError, ValueError, KeyError, sqlite3.Error):
            return False

    def _sweep_staging(self):
        # Only the single writer calls this, before creating its own staging
        # file. Reader-side collection must never unlink an active build.
        for path in self.directory.glob("*.building*"):
            if re.fullmatch(r"[a-f0-9]{32}\.building(?:-journal)?", path.name):
                path.unlink(missing_ok=True)
        (self.directory / "CURRENT.tmp").unlink(missing_ok=True)

    def _collect(self):
        current = self._current()
        keep = current["generation"] if current else None
        for pattern in ("*.sqlite", "*.json"):
            for path in self.directory.glob(pattern):
                if (
                    path.stem != keep
                    and not self.pins.get(path.stem)
                    and path.stem not in self.protected
                ):
                    path.unlink(missing_ok=True)

    def _manifest(self, sources: list[Source], pool: IndexPool) -> list[dict]:
        result = []
        for source in sorted(sources, key=lambda s: s.project_id):
            if not re.fullmatch(r"[a-f0-9]{32}", source.project_id):
                raise ValueError("Invalid project id")
            try:
                index = pool.get(source.path)
                meta = index.meta
                sha = meta.get("sha256", "")
                valid = isinstance(sha, str) and bool(
                    re.fullmatch(r"[a-f0-9]{64}", sha)
                )
                result.append(
                    {
                        "revision": revision(source.project_id, sha) if valid else None,
                        "projectId": source.project_id,
                        "filename": source.filename,
                        "importedAt": source.imported_at,
                        "nodes": meta["nodeCount"]
                        if type(meta.get("nodeCount")) is int
                        else len(index),
                        "indexBytes": source.path.stat().st_size,
                        "indexSchema": meta.get("schemaVersion"),
                        "sourceComplete": meta.get("status") == "complete",
                        "rebuildRequired": not valid,
                    }
                )
            except (OSError, sqlite3.Error, ValueError):
                result.append(
                    {
                        "revision": None,
                        "projectId": source.project_id,
                        "filename": source.filename,
                        "nodes": 0,
                        "rebuildRequired": True,
                        "sourceComplete": False,
                    }
                )
        return result

    def ensure(self, sources: Iterable[Source], *, force: bool = False) -> dict:
        sources = list(sources)
        with (
            self.writer,
            self.source_reads if self.limits.connections == 1 else nullcontext(),
            IndexPool(1) as pool,
        ):
            started = time.monotonic()
            manifest = self._manifest(sources, pool)
            self.active_manifest = manifest
            signature = self._signature(manifest)
            current = self._current()
            if (
                not force
                and current
                and current["signature"] == signature
                and self._published(current)
            ):
                self.build_state = {"status": "ready", "phase": "complete"}
                return current
            self.build_state = {"status": "building", "startedAt": time.time()}
            try:
                result = self._build(sources, manifest, signature, pool, started)
                self.build_state = {
                    "status": "ready",
                    "durationMs": result["buildDurationMs"],
                }
                return result
            except Exception as error:
                cap = (
                    "catalog_bytes"
                    if "catalog_bytes" in str(error)
                    or isinstance(error, sqlite3.Error)
                    and "full" in str(error).lower()
                    else "build_failed"
                )
                self.build_state = {
                    "status": "failed",
                    "error": type(error).__name__,
                    "capsHit": [cap],
                }
                self.failed_manifest = manifest
                for item in manifest:
                    item.setdefault(
                        "coverage",
                        {
                            "nodesTotal": item["nodes"],
                            "nodesScanned": 0,
                            "nodesUnscanned": item["nodes"],
                            "entries": 0,
                            "numericObserved": 0,
                            "numericLabelled": 0,
                            "textTruncated": 0,
                            "capsHit": [],
                        },
                    )
                    item["coverage"]["capsHit"] = sorted(
                        {
                            *item["coverage"]["capsHit"],
                            cap,
                        }
                    )
                raise

    def _signature(self, manifest):
        from linexcel.catalog_links import pin_signature

        return hashlib.sha256(
            _json(
                [
                    CATALOG_VERSION,
                    self.fts5,
                    asdict(self.limits),
                    manifest,
                    pin_signature(self),
                ]
            ).encode()
        ).hexdigest()

    def needs_refresh(self, sources):
        with IndexPool(1) as pool:
            current = self._current()
            return not (
                current
                and current["signature"]
                == self._signature(self._manifest(list(sources), pool))
                and self._published(current)
            )

    def failed_status(self, *, bounded=True) -> dict:
        coverage = self._coverage(self.failed_manifest)
        coverage.update(complete=False, capsHit=self.build_state["capsHit"])
        current = self._current()
        result = {
            "generation": current["generation"] if current else None,
            "revisions": list(self.failed_manifest),
            "coverage": coverage,
            "buildState": dict(self.build_state),
            "untrustedContent": True,
        }
        return self._bound(result) if bounded else result

    def _bound(self, result: dict) -> dict:
        """Bound every public search/status envelope, including build failures."""
        budget = self.limits.response_bytes
        if len(_json(result).encode()) <= budget:
            return result
        coverage = result["coverage"]
        coverage.update(
            complete=False, capsHit=sorted({*coverage["capsHit"], "response_bytes"})
        )
        trace: dict[str, Any] = result.get("trace", {})
        if trace:
            trace.update(complete=False, capsHit=coverage["capsHit"])
        for key, target in (
            ("results", result),
            ("sheets", coverage),
            ("revisions", trace),
            ("revisions", result),
        ):
            if not target or key not in target:
                continue
            entries = target[key]
            size = len(_json(result).encode())
            while size > budget and entries:
                removed = entries.pop()
                size -= len(_json(removed).encode()) + 1
                if key == "results":
                    trace["counts"]["returned"] = len(entries)
                    result["pagination"]["hasMore"] = True
                elif key == "sheets":
                    coverage["sheetsOmitted"] = coverage.get("sheetsOmitted", 0) + 1
                else:
                    counts = trace["counts"] if target is trace else coverage
                    counts["revisionsOmitted"] = counts.get("revisionsOmitted", 0) + 1
            # Account for count fields added while trimming and their digit widths.
            while len(_json(result).encode()) > budget and entries:
                entries.pop()
                if key == "results":
                    trace["counts"]["returned"] = len(entries)
                elif key == "sheets":
                    coverage["sheetsOmitted"] += 1
                else:
                    counts["revisionsOmitted"] += 1
        if len(_json(result).encode()) > budget:
            raise ValueError("response_bytes: metadata exceeds budget; narrow scope")
        return result

    def _build(self, sources, manifest, signature, pool, start):
        generation = uuid.uuid4().hex
        scratch = self.directory / f"{generation}.building"
        caps: set[str] = set()
        with self.readers:
            self._sweep_staging()
            self._collect()
            used = sum(
                p.stat().st_size for p in self.directory.iterdir() if p.is_file()
            )
        # Include manifest and atomic pointer overhead, as well as all retained
        # generations, in the owner's catalog budget. SQLite enforces the rest.
        manifest_reserve = 4096 + 2 * len(_json(manifest).encode())
        available = self.limits.catalog_bytes - used - manifest_reserve
        if available < 40960:
            raise ValueError(
                "catalog_bytes: insufficient space for a new immutable generation"
            )
        db = sqlite3.connect(scratch, isolation_level=None)
        db.execute("PRAGMA journal_mode=MEMORY")
        db.execute("PRAGMA synchronous=OFF")
        db.execute("PRAGMA max_page_count=" + str(available // 4096))
        source_map = {s.project_id: s for s in sources}
        try:
            schema(db, self.fts5)
            from linexcel.catalog_links import build_links

            link_summary = build_links(self, db, sources, manifest, pool, start)
            completed_nodes = 0
            for item in manifest:
                if item["rebuildRequired"]:
                    item["coverage"] = {
                        "nodesTotal": item["nodes"],
                        "nodesScanned": 0,
                        "nodesUnscanned": item["nodes"],
                        "entries": 0,
                        "numericObserved": 0,
                        "numericLabelled": 0,
                        "textTruncated": 0,
                        "capsHit": ["index_rebuild_required"],
                    }
                    caps.add("index_rebuild_required")
                    continue
                builder = WorkbookBuilder(
                    db,
                    pool.get(source_map[item["projectId"]].path),
                    item,
                    self.limits,
                    self.fts5,
                    lambda: time.monotonic() - start >= self.limits.build_seconds,
                    lambda scanned: self.build_state.update(
                        nodesScanned=completed_nodes + scanned
                    ),
                )
                self.build_state.update(
                    phase="headers", nodesTotal=sum(r["nodes"] for r in manifest)
                )
                builder.structural()
                self.build_state.update(
                    phase="cells",
                    projectId=item["projectId"],
                    sheetsTotal=len(builder.sheets),
                )
                builder.body()
                completed_nodes += item["coverage"]["nodesScanned"]
                caps.update(item["coverage"]["capsHit"])
            db.close()
            result: dict[str, Any] = {
                "version": CATALOG_VERSION,
                "generation": generation,
                "signature": signature,
                "backend": "fts5" if self.fts5 else "like",
                "revisions": manifest,
                "links": link_summary,
                "capsHit": sorted(caps),
                "limits": asdict(self.limits),
                "bytes": scratch.stat().st_size,
                "buildDurationMs": round((time.monotonic() - start) * 1000, 3),
            }
            # Public manifest contains no filesystem paths or source cell text.
            with self.readers:
                self.protected.add(generation)
            manifest_path = self.directory / f"{generation}.json"
            manifest_path.write_text(_json(result), "utf-8")
            if (
                sum(p.stat().st_size for p in self.directory.iterdir() if p.is_file())
                + 256
                > self.limits.catalog_bytes
            ):
                raise ValueError(
                    "catalog_bytes: generation metadata exceeds remaining budget"
                )
            scratch.replace(self.directory / f"{generation}.sqlite")
            temporary = self.directory / "CURRENT.tmp"
            temporary.write_text(
                _json({"generation": generation, "signature": signature}), "utf-8"
            )
            with self.readers:
                temporary.replace(self.directory / "CURRENT")
                self._collect()
            return result
        finally:
            db.close()
            scratch.unlink(missing_ok=True)
            with self.readers:
                self.protected.discard(generation)
                self._collect()

    def status(self, *, bounded=True) -> dict:
        with self.pin() as (_, current):
            result = json.loads(
                (self.directory / f"{current['generation']}.json").read_text("utf-8")
            )
            result: dict[str, Any] = {
                **result,
                "coverage": self._coverage(result["revisions"]),
                "buildState": dict(self.build_state),
                "totalCatalogBytes": sum(
                    p.stat().st_size for p in self.directory.iterdir() if p.is_file()
                ),
                "untrustedContent": True,
            }
            return self._bound(result) if bounded else result

    @staticmethod
    def _coverage(revisions: list[dict]) -> dict:
        caps = sorted({c for r in revisions for c in r["coverage"]["capsHit"]})
        keys = (
            "nodesTotal",
            "nodesScanned",
            "entries",
            "numericObserved",
            "numericLabelled",
            "textTruncated",
            "nodesUnscanned",
        )
        return {
            **{key: sum(r["coverage"][key] for r in revisions) for key in keys},
            "revisionsTotal": len(revisions),
            "capsHit": caps,
            "complete": not caps and all(r["sourceComplete"] for r in revisions),
            "numericPolicy": (
                "Labelled finite numeric cells only; labels are heuristics"
            ),
            "labelMethods": (
                "cell_labels: eight cells left/above, tables and exact names; "
                "catalog-only bounded header_band inference"
            ),
            "sheets": [
                {"revision": r["revision"], **sheet}
                for r in revisions
                for sheet in r["coverage"].get("sheets", [])
            ],
        }

    def validate_search(self, query, *, kinds=(), offset=0, limit=50):
        if not isinstance(query, str) or len(query) > self.limits.query_chars:
            raise ValueError("Query exceeds query_chars")
        kinds = sorted(set(kinds))
        if set(kinds) - KINDS:
            raise ValueError("Unknown search kind")
        if (
            isinstance(offset, bool)
            or isinstance(limit, bool)
            or not isinstance(offset, int)
            or not isinstance(limit, int)
            or offset < 0
            or offset > self.limits.hits
            or limit < 1
            or limit > self.limits.hits
        ):
            raise ValueError("Pagination exceeds hits budget")

    def search(
        self,
        sources: Iterable[Source],
        query: str,
        *,
        kinds=(),
        project=None,
        revision_id=None,
        offset=0,
        limit=50,
        generation=None,
        reconcile=True,
    ) -> dict:
        sources = list(sources)
        self.validate_search(query, kinds=kinds, offset=offset, limit=limit)
        kinds = sorted(set(kinds))
        try:
            if reconcile:
                self.ensure(sources)
        except (OSError, ValueError, sqlite3.Error):
            if self.build_state["status"] != "failed":
                raise
            status = self.failed_status(bounded=False)
            revisions = [
                r["revision"]
                for r in status["revisions"]
                if (project is None or r["projectId"] == project)
                and (
                    revision_id is None
                    or (r["revision"] and r["revision"]["id"] == revision_id)
                )
            ]
            if (project or revision_id) and not revisions:
                raise KeyError("Revision not found") from None
            return self._bound(
                {
                    "status": "failed",
                    "generation": status["generation"],
                    "results": [],
                    "coverage": status["coverage"],
                    "buildState": status["buildState"],
                    "pagination": {
                        "offset": offset,
                        "limit": limit,
                        "matched": 0,
                        "hasMore": False,
                    },
                    "trace": {
                        "schemaVersion": TRACE_VERSION,
                        "tool": "search_catalog",
                        "terms": terms_for(query),
                        "normalization": (
                            "NFD, casefold, remove combining marks; word tokens; AND"
                        ),
                        "scope": {
                            "project": project,
                            "revision": revision_id,
                            "kinds": kinds,
                        },
                        "revisions": revisions,
                        "generation": status["generation"],
                        "backend": "fts5" if self.fts5 else "like",
                        "counts": {
                            "eligibleEntries": 0,
                            "matched": 0,
                            "returned": 0,
                            "sqliteVmStepsApprox": 0,
                            "indexConnectionsPeak": 0,
                        },
                        "capsHit": status["coverage"]["capsHit"],
                        "durationMs": 0,
                        "complete": False,
                    },
                    "untrustedContent": True,
                }
            )
        start = time.monotonic()
        folded = normalize(query)
        terms = terms_for(folded)
        try:
            numeric = float(query.strip())
            if not math.isfinite(numeric):
                numeric = None
        except ValueError:
            numeric = None
        with (
            self.source_reads,
            self.pin() as (db, current),
            IndexPool(max(1, self.limits.connections - 1)) as pool,
        ):
            if generation and generation != current["generation"]:
                raise ValueError("Catalog generation changed; restart pagination")
            manifest = json.loads(
                (self.directory / f"{current['generation']}.json").read_text("utf-8")
            )
            authorized = {
                r["revision"]["id"]
                for r in self._manifest(sources, pool)
                if r["revision"]
            }
            stale = sum(
                bool(r["revision"]) and r["revision"]["id"] not in authorized
                for r in manifest["revisions"]
            )
            revisions = [
                r
                for r in manifest["revisions"]
                if not r["revision"] or r["revision"]["id"] in authorized
                if (project is None or r["projectId"] == project)
                and (
                    revision_id is None
                    or (r["revision"] and r["revision"]["id"] == revision_id)
                )
            ]
            if (project or revision_id) and not revisions:
                raise KeyError("Revision not found")
            coverage = self._coverage(revisions)
            caps = set(coverage["capsHit"])
            if stale:
                caps.add("stale_generation")
                coverage["staleRevisionsOmitted"] = stale
            source_map = {s.project_id: s for s in sources}
            rev_map = {r["projectId"]: r for r in revisions}
            clauses, args = [], []
            if revisions:
                clauses.append(
                    "e.project IN (" + ",".join("?" for _ in revisions) + ")"
                )
                args.extend(r["projectId"] for r in revisions)
            else:
                clauses.append("0")
            if kinds:
                clauses.append("e.kind IN (" + ",".join("?" for _ in kinds) + ")")
                args.extend(kinds)
            backend = manifest["backend"]
            if numeric is not None:
                clauses.append("n.number=?")
                args.append(numeric)
            elif backend == "fts5" and terms:
                clauses.append(
                    "e.id IN (SELECT rowid FROM search WHERE search MATCH ?)"
                )
                args.append(" AND ".join('"' + t + '"' for t in terms))
            elif terms:
                for term in terms:
                    clauses.append("e.text LIKE ? ESCAPE '\\'")
                    args.append(
                        "%" + term.replace("\\", "\\\\").replace("_", "\\_") + "%"
                    )
                # LIKE is a bounded coarse filter; retain FTS word boundaries.
                db.create_function(
                    "catalog_token_match",
                    1,
                    lambda text: int(set(terms) <= set(terms_for(text))),
                    deterministic=True,
                )
                clauses.append("catalog_token_match(e.text)")
            elif query.strip():
                clauses.append("0")
            scanned = 0
            deadline = start + self.limits.query_ms / 1000

            def progress():
                nonlocal scanned
                scanned += 100
                return int(time.monotonic() >= deadline)

            db.set_progress_handler(progress, 100)
            # LIKE's input is bounded BEFORE matching, with an explicit count.
            base = "entries"
            eligible = sum(r["coverage"]["entries"] for r in revisions)
            if backend == "like" and numeric is None and terms:
                base = "(SELECT * FROM entries ORDER BY id LIMIT ?)"
                args.insert(0, self.limits.scan_entries)
                total_entries = sum(
                    r["coverage"]["entries"] for r in manifest["revisions"]
                )
                if total_entries > self.limits.scan_entries:
                    caps.add("scan_entries")
            rows = []
            try:
                select = (
                    "e.id,e.project,e.kind,n.node,e.sheet,n.address,n.labels,e.name,"
                    "1,'[]',1"
                    if numeric is not None
                    else "e.id,e.project,e.kind,e.node,e.sheet,e.address,"
                    "e.labels,e.name,"
                    "e.occurrence_count,e.occurrences,e.count_complete"
                )
                join = (
                    " JOIN numeric_cells n ON n.entry=e.id"
                    if numeric is not None
                    else ""
                )
                cursor = db.execute(
                    "SELECT "
                    + select
                    + " FROM "
                    + base
                    + " e"
                    + join
                    + " WHERE "
                    + " AND ".join(clauses)
                    + " ORDER BY e.project,e.kind,e.id"
                    + (",n.id" if join else "")
                    + " LIMIT ?",
                    (*args, self.limits.hits + 1),
                )
                for row in cursor:
                    rows.append(row)
            except sqlite3.OperationalError as error:
                if "interrupt" not in str(error):
                    raise
                caps.add("query_ms")
            finally:
                db.set_progress_handler(None, 0)
            if len(rows) > self.limits.hits:
                caps.add("hits")
            matched = min(len(rows), self.limits.hits)
            hits = []
            evidence_bytes = 0
            for (
                entry_id,
                pid,
                kind,
                nid,
                sheet,
                address,
                labels,
                name,
                occurrence_count,
                occurrences,
                count_complete,
            ) in rows[offset : min(offset + limit, self.limits.hits)]:
                item = rev_map[pid]
                node = pool.get(source_map[pid].path)[nid] if nid else None
                hit = {
                    "kind": kind,
                    "occurrenceCount": occurrence_count,
                    "occurrenceCountComplete": bool(count_complete),
                    "occurrenceScope": (
                        "indexed cells; unassessed/unscanned cells are excluded"
                    ),
                    "occurrences": [
                        {
                            "address": ref["address"],
                            "handle": cell_handle(item["revision"], ref["node"]),
                        }
                        for ref in json.loads(occurrences)
                        if ref.get("node")
                    ],
                    "occurrencesOmitted": max(
                        0, occurrence_count - len(json.loads(occurrences))
                    ),
                    "revision": item["revision"],
                    "filename": item["filename"],
                    "sheet": sheet,
                    "address": address,
                    "labels": json.loads(labels),
                    "handle": cell_handle(item["revision"], nid)
                    if node and kind == "cell"
                    else None,
                    "objectId": nid or f"catalog:{current['generation']}:{entry_id}",
                    "name": name,
                    "value": node.get("value") if node else None,
                    "valueProvenance": node.get("provenance", "unknown")
                    if node
                    else None,
                    "cacheStatus": node.get("cache_status") if node else None,
                    "formula": node.get("formula") if node else None,
                    "untrustedContent": True,
                }
                if numeric is not None and node:
                    hit["occurrences"] = [{"address": address, "handle": hit["handle"]}]
                    hit["occurrencesOmitted"] = 0
                if kind != "cell":
                    hit["occurrencesOmitted"] = 0
                elif numeric is None:
                    hit["occurrenceGroup"] = {
                        "generation": current["generation"],
                        "id": entry_id,
                    }
                if node and node.get("name"):
                    hit["name"] = node["name"]
                size = len(_json(hit).encode())
                if evidence_bytes + size > self.limits.response_bytes:
                    caps.add("response_bytes")
                    break
                hits.append(hit)
                evidence_bytes += size
            result: dict[str, Any] = {
                "generation": current["generation"],
                "results": hits,
                "pagination": {
                    "offset": offset,
                    "limit": limit,
                    "matched": matched,
                    "hasMore": offset + len(hits) < matched,
                },
                "coverage": coverage,
                "trace": {
                    "schemaVersion": TRACE_VERSION,
                    "tool": "search_catalog",
                    "terms": terms,
                    "normalization": (
                        "NFD, casefold, remove combining marks; word tokens; AND"
                    ),
                    "scope": {
                        "project": project,
                        "revision": revision_id,
                        "kinds": kinds,
                    },
                    "revisions": [r["revision"] for r in revisions],
                    "generation": current["generation"],
                    "backend": backend,
                    "counts": {
                        "eligibleEntries": eligible,
                        "matched": matched,
                        "returned": len(hits),
                        "sqliteVmStepsApprox": scanned,
                        "indexConnectionsPeak": pool.peak,
                    },
                    "capsHit": sorted(caps),
                    "durationMs": round((time.monotonic() - start) * 1000, 3),
                    "complete": coverage["complete"] and not caps,
                },
                "untrustedContent": True,
            }
            result["coverage"] = {
                **coverage,
                "capsHit": sorted(caps),
                "complete": coverage["complete"] and not caps,
            }
            result["trace"].update(
                capsHit=sorted(caps), complete=result["coverage"]["complete"]
            )
            return self._bound(result)

    def occurrences(self, sources, *, generation, group, after=0, limit=50):
        from linexcel.catalog_occurrences import occurrences

        return occurrences(
            self, sources, generation=generation, group=group, after=after, limit=limit
        )

    def lineage(
        self,
        sources: Iterable[Source],
        handle: str,
        *,
        depth=1,
        limit=200,
        direction=None,
        edges=None,
        workbooks=None,
        milliseconds=None,
    ) -> dict:
        if direction is not None:
            from linexcel.catalog_lineage import directed_lineage

            return directed_lineage(
                self,
                list(sources),
                handle,
                direction=direction,
                depth=depth,
                limit=limit,
                edges=edges,
                workbooks=workbooks,
                milliseconds=milliseconds,
            )
        pid, sha, nid = parse_handle(handle)
        if not 0 <= depth <= min(
            20, self.limits.lineage_depth
        ) or not 1 <= limit <= min(200, self.limits.lineage_nodes):
            raise ValueError("Lineage budget exceeded")
        # Authorize using the current owner's source list before opening a path.
        source = next((s for s in sources if s.project_id == pid), None)
        if source is None:
            raise KeyError("Cell not found")
        with self.source_reads:
            with WorkbookIndex(source.path) as index:
                if index.meta.get("sha256") != sha or nid not in index:
                    raise KeyError("Cell revision not found")
            graph = graph_neighborhood(source.path, nid, limit=limit, depth=depth)
        rev = revision(pid, sha)
        for node in graph["nodes"]:
            node["revision"] = rev
            node["handle"] = cell_handle(rev, node["id"])
        result = {
            "revision": rev,
            "filename": source.filename,
            "graph": graph,
            "coverage": {
                "complete": False,
                "scope": "same-workbook neighborhood",
                "nodesReturned": len(graph["nodes"]),
                "depth": depth,
                "capsHit": ["lineage_nodes"]
                if graph["pagination"]["truncated"]
                else [],
                "boundaries": ["requested_depth", "external_workbooks_not_traversed"],
            },
            "untrustedContent": True,
        }
        while len(_json(result).encode()) > self.limits.response_bytes:
            caps = result["coverage"]["capsHit"]
            if "response_bytes" not in caps:
                caps.append("response_bytes")
            if graph["formulaPatterns"]:
                graph["formulaPatterns"].pop()
            elif len(graph["nodes"]) > 1:
                graph["nodes"].pop()
                ids = {node["id"] for node in graph["nodes"]}
                graph["edges"] = [
                    e
                    for e in graph["edges"]
                    if e["source"] in ids and e["target"] in ids
                ]
                graph["pagination"].update(returned=len(ids), truncated=True)
                result["coverage"]["nodesReturned"] = len(ids)
            else:
                raise ValueError(
                    "response_bytes: cell evidence exceeds response budget"
                )
        return result
