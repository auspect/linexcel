"""Bounded source scans for generation-local text groups, with resumable pages."""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import time
from typing import Any, cast

from linexcel.header_bands import Band, HeaderBands, catalog_labels
from linexcel.lazy_store import WorkbookIndex, _unpack


def source_rows(index, sheet, after, limits, deadline, caps):
    """Interrupt SQLite work as well as Python processing at the deadline."""
    cursor = None
    while True:
        index.db.set_progress_handler(lambda: int(time.monotonic() >= deadline), 100)
        try:
            if cursor is None:
                cursor = index.db.execute(
                    "SELECT seq,id,row,col,data FROM nodes NOT INDEXED "
                    "WHERE sheet IS ? AND seq>? "
                    "AND kind IN ('input','cell','array_result') ORDER BY seq LIMIT ?",
                    (sheet, after, limits.scan_entries + 1),
                )
            row = cursor.fetchone()
        except sqlite3.OperationalError as error:
            if "interrupt" not in str(error):
                raise
            caps.append("query_ms")
            return
        finally:
            index.db.set_progress_handler(None, 0)
        if row is None:
            return
        yield row


def occurrences(catalog, sources, *, generation, group, after=0, limit=50):
    from linexcel.catalog import _json, brief, cell_handle, normalize, revision

    limits = catalog.limits
    if (
        not isinstance(generation, str)
        or not re.fullmatch(r"[a-f0-9]{32}", generation)
        or any(type(n) is not int for n in (group, after, limit))
        or not 1 <= group <= 2**63 - 1
        or not 0 <= after <= 2**63 - 1
        or not 1 <= limit <= limits.hits
    ):
        raise ValueError("Invalid occurrence page")
    start = time.monotonic()
    with catalog.source_reads, catalog.pin() as (db, current):
        if current["generation"] != generation or not catalog._published(current):
            raise ValueError("Catalog generation changed; restart occurrence lookup")
        entry = db.execute(
            "SELECT project,sheet,group_key FROM entries WHERE id=? AND kind='cell'",
            (group,),
        ).fetchone()
        if not entry or entry[2] is None:
            raise KeyError("Text group not found")
        pid, sheet, key = entry
        source = next((s for s in sources if s.project_id == pid), None)
        if source is None:
            raise KeyError("Text group not found")
        manifest = json.loads(
            (catalog.directory / f"{generation}.json").read_text("utf-8")
        )
        item = next(r for r in manifest["revisions"] if r["projectId"] == pid)
        with WorkbookIndex(source.path) as index:
            sha = index.meta.get("sha256")
            if not item["revision"] or item["revision"]["sha256"] != sha:
                raise KeyError("Text group revision not found")
            rev = revision(pid, sha)
            record = db.execute(
                "SELECT data FROM header_bands WHERE project=? AND sheet IS ?",
                (pid, sheet),
            ).fetchone()
            bands = HeaderBands.__new__(HeaderBands)
            bands.sheet, bands.rows = sheet, {}
            bands.bands = (
                [
                    Band(
                        b["header_row"],
                        b["first_row"],
                        b["last_row"],
                        {int(c): t for c, t in b["headers"].items()},
                        b["row_columns"],
                    )
                    for b in json.loads(record[0])
                ]
                if record
                else []
            )
            result: dict[str, Any] = {
                "generation": generation,
                "group": group,
                "revision": rev,
                "sheet": sheet,
                "results": [],
                "pagination": {"after": after, "limit": limit, "nextAfter": None},
                "coverage": {
                    "scope": (
                        "authoritative matching cells in this group's sheet; "
                        "includes unassessed cells"
                    ),
                    "complete": False,
                    "cellsScanned": 0,
                    "capsHit": [],
                },
                "untrustedContent": True,
            }
            coverage = cast(dict[str, Any], result["coverage"])
            last = after
            exhausted = True
            for seq, nid, row, col, blob in source_rows(
                index,
                sheet,
                after,
                limits,
                start + limits.query_ms / 1000,
                coverage["capsHit"],
            ):
                if coverage["cellsScanned"] >= limits.scan_entries:
                    coverage["capsHit"].append("scan_entries")
                    exhausted = False
                    break
                if time.monotonic() - start >= limits.query_ms / 1000:
                    coverage["capsHit"].append("query_ms")
                    exhausted = False
                    break
                coverage["cellsScanned"] += 1
                node = _unpack(blob)
                value = node.get("value")
                text = value if isinstance(value, str) and value.strip() else None
                if text is None and row and col:
                    # Populate only this row's literal context, never the sheet.
                    if row not in bands.rows:
                        columns = sorted(
                            {c for b in bands.bands for c in b.row_columns}
                        )
                        bands.rows = {row: {}}
                        if columns:
                            bands.rows[row] = {
                                c: _unpack(data)
                                for c, data in index.db.execute(
                                    "SELECT col,data FROM nodes "
                                    "WHERE sheet IS ? AND row=? "
                                    "AND col IN ("
                                    + ",".join("?" for _ in columns)
                                    + ")",
                                    (sheet, row, *columns),
                                )
                            }
                    labels = catalog_labels(
                        index, sheet, row, col, bands.labels(row, col)
                    )
                    text = " ".join(label["text"] for label in labels)
                if text and hashlib.sha256(normalize(text).encode()).digest() == key:
                    if len(result["results"]) >= limit:
                        coverage["capsHit"].append("page_limit")
                        exhausted = False
                        break
                    hit = {
                        "address": node.get("cell"),
                        "handle": cell_handle(rev, nid),
                        "value": brief(value),
                    }
                    result["results"].append(hit)
                    # Reserve the continuation envelope before committing a hit.
                    result["pagination"]["nextAfter"] = seq
                    if len(_json(result).encode()) + 128 > limits.response_bytes:
                        result["results"].pop()
                        coverage["capsHit"].append("response_bytes")
                        exhausted = False
                        break
                last = seq
            if "query_ms" in coverage["capsHit"]:
                exhausted = False
            coverage["complete"] = exhausted
            result["pagination"]["nextAfter"] = None if exhausted else last
            if not exhausted and last == after:
                raise ValueError(
                    "Occurrence budget cannot advance; increase query/response budget"
                )
            if len(_json(result).encode()) > limits.response_bytes:
                raise ValueError("response_bytes: occurrence envelope exceeds budget")
            return result
