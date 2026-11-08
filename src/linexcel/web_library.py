"""Library adapter: authorization and source mutation use the task-store lock."""

from __future__ import annotations

import json
import sqlite3
import threading

from linexcel.catalog import (
    TRACE_VERSION,
    Catalog,
    IndexPool,
    Source,
    fts5_available,
    revision,
    terms_for,
)


class Library:
    def __init__(self, store):
        self.store = store
        self.catalogs: dict[str, Catalog] = {}
        self.fts5 = fts5_available()
        self.pending = {}
        self.dirty = set()
        self.mutations = {}

    def catalog(self, owner: str) -> Catalog:
        if owner not in self.catalogs:
            self.catalogs[owner] = Catalog(
                self.store.root / owner / "catalog", self.store.config.catalog_limits
            )
        return self.catalogs[owner]

    def sources(self, owner: str) -> list[Source]:
        from linexcel.web import _read

        sources = []
        for manifest in (self.store.root / owner / "projects").glob("*/project.json"):
            if (owner, manifest.parent.name) in self.store.unreadable_projects:
                continue
            project = _read(manifest)
            path = self.store.project_path(owner, project["id"])
            sources.append(
                Source(
                    project["id"],
                    project["filename"],
                    path / "index.sqlite",
                    project["createdAt"],
                )
            )
        return sources

    def mutation(self, owner):
        """Acquire before store.lock, never wait for a build under that lock."""
        with self.store.lock:
            return self.mutations.setdefault(owner, threading.RLock())

    def refresh(self, owner: str):
        """Coalesce rebuilds on the existing bounded task executor."""
        with self.store.lock:
            if self.store.closed:
                return
            self.dirty.add(owner)
            if owner in self.pending:
                return
            catalog = self.catalog(owner)
            catalog.build_state = {"status": "queued", "phase": "catalog"}
            self.pending[owner] = self.store.executor.submit(self._reconcile, owner)

    def _reconcile(self, owner):
        try:
            while True:
                with self.mutation(owner):
                    with self.store.lock:
                        self.dirty.discard(owner)
                        sources = self.sources(owner)
                        catalog = self.catalog(owner)
                    try:
                        catalog.ensure(sources)
                    except (OSError, ValueError, sqlite3.Error):
                        pass  # Catalog records failed state; imports stay successful.
                with self.store.lock:
                    if owner not in self.dirty or self.store.closed:
                        return
        finally:
            with self.store.lock:
                self.pending.pop(owner, None)
                if owner in self.dirty and not self.store.closed:
                    self.refresh(owner)

    def _pending_status(self, catalog, sources, project=None, revision_id=None):
        from linexcel.catalog_build import coverage

        projects = {s.project_id for s in sources}
        revisions = [
            {**item}
            for item in catalog.active_manifest
            if item["projectId"] in projects
            and (not project or item["projectId"] == project)
            and (
                not revision_id
                or item["revision"]
                and item["revision"]["id"] == revision_id
            )
        ]
        for item in revisions:
            item["coverage"] = coverage(item["nodes"])
            item["coverage"]["capsHit"] = ["catalog_refresh_pending"]
        counts = catalog._coverage(revisions)
        counts.update(complete=False, capsHit=["catalog_refresh_pending"])
        return catalog._bound(
            {
                "generation": None,
                "revisions": revisions,
                "coverage": counts,
                "buildState": dict(catalog.build_state),
                "stale": True,
                "untrustedContent": True,
            }
        )

    def query(
        self, owner: str, operation: str, params: dict, *, synchronous=False
    ) -> dict:
        from linexcel.web import APIError

        allowed = {
            "search": {
                "q",
                "query",
                "kinds",
                "project",
                "revision",
                "offset",
                "limit",
                "generation",
            },
            "status": set(),
            "lineage": {"handle", "depth", "limit"},
        }
        if operation not in allowed or set(params) - allowed[operation]:
            raise APIError(400, "Invalid library parameter")

        def value(key, default=""):
            values = params.get(key, [default])
            if (
                len(values) != 1
                or not isinstance(values[0], str)
                or len(values[0]) > 4096
            ):
                raise ValueError("Invalid library parameter")
            return values[0]

        try:
            with self.store.lock:
                catalog = self.catalog(owner)
                sources = self.sources(owner)
                if operation == "lineage":
                    if owner in self.pending and catalog.limits.connections == 1:
                        raise APIError(
                            409, "Catalog building; retry lineage after completion"
                        )
                    return catalog.lineage(
                        sources,
                        value("handle"),
                        depth=int(value("depth", "1")),
                        limit=int(value("limit", "200")),
                    )
                if operation == "search":
                    if "q" in params and "query" in params:
                        raise ValueError("Supply q or query, not both")
                    if value("project") and value("project") not in {
                        s.project_id for s in sources
                    }:
                        raise KeyError("Revision not found")
                    if value("revision"):
                        if owner in self.pending and catalog.limits.connections == 1:
                            manifest = catalog.active_manifest
                        else:
                            with IndexPool(1) as pool:
                                manifest = catalog._manifest(sources, pool)
                        valid = {r["revision"]["id"] for r in manifest if r["revision"]}
                        if value("revision") not in valid:
                            raise KeyError("Revision not found")
                        if value("project") and not value("revision").startswith(
                            value("project") + ":"
                        ):
                            raise KeyError("Revision not found")
                    catalog.validate_search(
                        value("q", value("query")),
                        kinds=value("kinds").split(",") if value("kinds") else (),
                        offset=int(value("offset", "0")),
                        limit=int(value("limit", "50")),
                    )
                if synchronous:
                    try:
                        catalog.ensure(sources)
                    except (OSError, ValueError, sqlite3.Error):
                        if catalog.build_state["status"] != "failed":
                            raise
                        return catalog.failed_status()
                elif (
                    owner not in self.pending
                    and catalog.needs_refresh(sources)
                    and catalog.build_state["status"] != "failed"
                ):
                    self.refresh(owner)
                pending = owner in self.pending
                missing_scope = False
                if (
                    pending
                    and operation == "search"
                    and (value("project") or value("revision"))
                    and catalog._published(catalog._current())
                ):
                    with catalog.pin() as (_, current):
                        snapshot = json.loads(
                            (
                                catalog.directory / f"{current['generation']}.json"
                            ).read_text("utf-8")
                        )
                    missing_scope = not any(
                        (not value("project") or r["projectId"] == value("project"))
                        and (
                            not value("revision")
                            or r["revision"]
                            and r["revision"]["id"] == value("revision")
                        )
                        for r in snapshot["revisions"]
                    )
                if (
                    not catalog._published(catalog._current())
                    or missing_scope
                    or pending
                    and catalog.limits.connections == 1
                ):
                    if catalog.build_state["status"] == "failed":
                        result = catalog.failed_status()
                    else:
                        result = self._pending_status(
                            catalog,
                            sources,
                            value("project") if operation == "search" else None,
                            value("revision") if operation == "search" else None,
                        )
                    if operation == "search":
                        result["results"] = []
                        result["status"] = "building" if pending else "failed"
                        result["pagination"] = {
                            "offset": int(value("offset", "0")),
                            "limit": int(value("limit", "50")),
                            "matched": 0,
                            "hasMore": False,
                        }
                        result["trace"] = {
                            "schemaVersion": TRACE_VERSION,
                            "tool": "search_catalog",
                            "terms": terms_for(value("q", value("query"))),
                            "normalization": (
                                "NFD, casefold, remove combining marks; "
                                "word tokens; AND"
                            ),
                            "scope": {
                                "project": value("project") or None,
                                "revision": value("revision") or None,
                                "kinds": value("kinds").split(",")
                                if value("kinds")
                                else [],
                            },
                            "revisions": [
                                r["revision"] for r in result.get("revisions", [])
                            ],
                            "generation": result["generation"],
                            "backend": "fts5" if catalog.fts5 else "like",
                            "counts": {
                                "eligibleEntries": 0,
                                "matched": 0,
                                "returned": 0,
                                "sqliteVmStepsApprox": 0,
                                "indexConnectionsPeak": 0,
                            },
                            "durationMs": 0,
                            "complete": False,
                            "capsHit": result["coverage"]["capsHit"],
                        }
                    return catalog._bound(result)
                if operation == "status":
                    result = catalog.status(bounded=False)
                    # Removed/replaced revisions must never survive in stale
                    # status metadata, even while the replacement is building.
                    with IndexPool(1) as pool:
                        authorized = {
                            r["revision"]["id"]
                            for r in catalog._manifest(sources, pool)
                            if r["revision"]
                        }
                    result["revisions"] = [
                        r
                        for r in result["revisions"]
                        if r["revision"]
                        and r["revision"]["id"] in authorized
                        or not r["revision"]
                        and r["projectId"] in {s.project_id for s in sources}
                    ]
                    result["coverage"] = catalog._coverage(result["revisions"])
                    return self._state(catalog, result, pending)
                if "q" in params and "query" in params:
                    raise ValueError("Supply q or query, not both")
                result = catalog.search(
                    sources,
                    value("q", value("query")),
                    kinds=value("kinds").split(",") if value("kinds") else (),
                    project=value("project") or None,
                    revision_id=value("revision") or None,
                    offset=int(value("offset", "0")),
                    limit=int(value("limit", "50")),
                    generation=value("generation") or None,
                    reconcile=False,
                )
                return self._state(catalog, result, pending)
        except KeyError:
            raise APIError(404, "Library revision or cell not found") from None
        except (ValueError, OSError, sqlite3.Error) as error:
            state = self.catalog(owner).build_state
            if state["status"] == "failed":
                raise APIError(
                    503,
                    "Library build failed; inspect status and budgets",
                ) from None
            raise APIError(400, str(error)) from None

    @staticmethod
    def _state(catalog, result, pending):
        state = dict(catalog.build_state)
        result.update(buildState=state, stale=pending or state["status"] == "failed")
        if result["stale"]:
            coverage = result["coverage"]
            coverage.update(
                complete=False,
                capsHit=sorted(
                    {
                        *coverage["capsHit"],
                        "catalog_refresh_pending"
                        if pending
                        else "catalog_build_failed",
                    }
                ),
            )
            if "trace" in result:
                result["trace"].update(complete=False, capsHit=coverage["capsHit"])
        return catalog._bound(result)


def project_revision(project: dict, path) -> dict:
    """Read identity from the immutable index; old projects need no migration."""
    from linexcel.lazy_store import WorkbookIndex, file_sha256

    index = path / "index.sqlite"
    sha = None
    if index.is_file():
        try:
            with WorkbookIndex(index) as store:
                sha = store.meta.get("sha256")
        except (OSError, sqlite3.Error, ValueError):
            pass  # Catalog status exposes the separate rebuild requirement.
    if not sha and (path / "workbook.xlsx").is_file():
        sha = file_sha256(path / "workbook.xlsx")
    return {
        **project,
        "workbookRevision": revision(project["id"], sha) if sha else None,
    }
