"""Rebuildable link table and durable, owner-local revision pins."""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import time
from collections import Counter

from linexcel.link_facts import facts_available

STATUSES = {"resolved_by_name", "pinned", "ambiguous", "unresolved"}


def read_pins(catalog):
    path = catalog.pin_directory / "link-pins.sqlite"
    if not path.exists():
        return {}
    with sqlite3.connect(path) as db:
        return {
            key: {"projectId": pid, "sha256": sha}
            for key, pid, sha in db.execute("SELECT link,project,sha FROM pins")
        }


def pin_signature(catalog):
    return sorted(
        (key, value["projectId"], value["sha256"])
        for key, value in read_pins(catalog).items()
    )


def build_links(catalog, db, sources, manifest, pool, started):
    db.executescript("""
        CREATE TABLE links(id TEXT PRIMARY KEY,source_project TEXT NOT NULL,
            source_sha TEXT NOT NULL,source_node TEXT NOT NULL,
            target_project TEXT,target_sha TEXT,target_sheet TEXT,
            r1 INTEGER,c1 INTEGER,r2 INTEGER,c2 INTEGER,status TEXT NOT NULL,
            data TEXT NOT NULL) WITHOUT ROWID;
        CREATE TABLE link_candidates(link TEXT NOT NULL,project TEXT NOT NULL,
            sha TEXT NOT NULL,PRIMARY KEY(link,project,sha)) WITHOUT ROWID;
    """)
    from linexcel.catalog import _json

    pins = read_pins(catalog)
    revisions = {r["projectId"]: r for r in manifest if r["revision"]}
    names = {}
    for item in revisions.values():
        names.setdefault(item["filename"].lower(), []).append(item["revision"])
    counts = Counter({s: 0 for s in sorted(STATUSES)})
    caps = set()
    missing = []
    total = 0
    for source in sorted(sources, key=lambda s: s.project_id):
        if source.project_id not in revisions:
            missing.append(source.project_id)
            continue
        item = revisions[source.project_id]
        index = pool.get(source.path)
        if not facts_available(index):
            missing.append(source.project_id)
            continue
        for node, ordinal, data in index.db.execute(
            "SELECT node,ordinal,data FROM external_facts ORDER BY node,ordinal"
        ):
            if total >= catalog.limits.lineage_links:
                caps.add("link_budget")
                break
            if time.monotonic() - started >= catalog.limits.build_seconds:
                caps.add("link_build_time")
                break
            fact = json.loads(data)
            key = hashlib.sha256(
                _json([item["revision"]["id"], node, ordinal]).encode()
            ).hexdigest()
            candidates = names.get(fact["filename"].lower(), [])
            target = None
            pinned = pins.get(key)
            stale = False
            if pinned:
                status = "pinned"
                candidate = revisions.get(pinned["projectId"], {}).get("revision")
                if candidate and candidate["sha256"] == pinned["sha256"]:
                    target = candidate
                else:
                    stale = True
            elif len(candidates) == 1:
                status, target = "resolved_by_name", candidates[0]
            else:
                status = "ambiguous" if candidates else "unresolved"
            fact.update(
                id=key,
                sourceRevision=item["revision"],
                sourceNode=node,
                status=status,
                targetRevision=target,
                candidates=candidates,
                pin=pinned,
                stalePin=stale,
                versionCaveat="May not be the revision the link was saved against."
                if status == "resolved_by_name"
                else None,
            )
            bounds = fact["bounds"] or [None] * 4
            db.execute(
                "INSERT INTO links VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    key,
                    source.project_id,
                    item["revision"]["sha256"],
                    node,
                    target["projectId"] if target else None,
                    target["sha256"] if target else None,
                    (fact["sheet"] or "").lower(),
                    *bounds,
                    status,
                    _json(fact),
                ),
            )
            reverse_candidates = {r["id"]: r for r in candidates}
            if target:
                reverse_candidates[target["id"]] = target
            for candidate in reverse_candidates.values():
                db.execute(
                    "INSERT INTO link_candidates VALUES (?,?,?)",
                    (key, candidate["projectId"], candidate["sha256"]),
                )
            counts[status] += 1
            total += 1
        if caps:
            break
    if missing:
        caps.add("external_facts_rebuild_required")
    if total:
        for statement in (
            "CREATE INDEX links_source ON links(source_project,source_sha,source_node)",
            "CREATE INDEX links_target ON links("
            "target_project,target_sha,target_sheet,r1,r2)",
            "CREATE INDEX link_candidate_target ON link_candidates(project,sha)",
        ):
            try:
                db.execute(statement)
            except sqlite3.OperationalError as error:
                if "full" not in str(error).lower():
                    raise
                caps.add("catalog_bytes")
                break
    return {
        "counts": dict(counts),
        "total": total,
        "complete": not caps,
        "capsHit": sorted(caps),
        "rebuildRequired": missing,
    }


def set_pin(catalog, sources, link, project=None, sha=None, *, remove=False):
    """Caller supplies only owner-authorized sources; no path comes from input."""
    if not isinstance(link, str) or not re.fullmatch(r"[a-f0-9]{64}", link):
        raise ValueError("Invalid link id")
    sources = list(sources)
    with catalog.writer, catalog.source_reads:
        catalog.ensure(sources)
        with catalog.pin() as (db, _):
            row = db.execute("SELECT data FROM links WHERE id=?", (link,)).fetchone()
            if row is None:
                raise KeyError("Link not found")
        if not remove:
            if not isinstance(project, str) or not re.fullmatch(
                r"[a-f0-9]{32}", project
            ):
                raise ValueError("Invalid target project")
            if not isinstance(sha, str) or not re.fullmatch(r"[a-f0-9]{64}", sha):
                raise ValueError("Invalid target hash")
            from linexcel.catalog import IndexPool

            with IndexPool(1) as pool:
                authorized = [
                    r["revision"]
                    for r in catalog._manifest(sources, pool)
                    if r["revision"]
                ]
            if not any(
                r["projectId"] == project and r["sha256"] == sha for r in authorized
            ):
                raise KeyError("Target revision not found")
        # This database is deliberately outside the disposable catalog directory.
        with sqlite3.connect(catalog.pin_directory / "link-pins.sqlite") as db:
            db.execute(
                "CREATE TABLE IF NOT EXISTS pins(link TEXT PRIMARY KEY,"
                "project TEXT NOT NULL,sha TEXT NOT NULL)"
            )
            if remove:
                db.execute("DELETE FROM pins WHERE link=?", (link,))
            else:
                db.execute(
                    "INSERT OR REPLACE INTO pins VALUES (?,?,?)", (link, project, sha)
                )
        catalog.ensure(sources)
        return {
            "linkId": link,
            "pin": None if remove else {"projectId": project, "sha256": sha},
            "generation": catalog._current()["generation"],
        }


def list_links(catalog, sources, *, status="", project="", offset=0, limit=50):
    from linexcel.catalog import IndexPool, _json

    if status and status not in STATUSES:
        raise ValueError("Invalid link status")
    catalog.validate_search("", offset=offset, limit=limit)
    sources = list(sources)
    if project and project not in {s.project_id for s in sources}:
        raise KeyError("Revision not found")
    with catalog.source_reads, catalog.pin() as (db, current), IndexPool(1) as pool:
        authorized = {
            r["revision"]["id"]
            for r in catalog._manifest(sources, pool)
            if r["revision"]
        }
        manifest = json.loads(
            (catalog.directory / f"{current['generation']}.json").read_text("utf-8")
        )
        results = []
        matched = 0
        stale = False
        # Bounded by lineage_links at build time; do not expose deleted revisions
        # from a pinned old generation during asynchronous reconciliation.
        for (data,) in db.execute("SELECT data FROM links ORDER BY id"):
            fact = json.loads(data)
            if fact["sourceRevision"]["id"] not in authorized:
                stale = True
                continue
            if (
                fact["targetRevision"]
                and fact["targetRevision"]["id"] not in authorized
            ):
                fact.update(targetRevision=None, stalePin=bool(fact["pin"]))
                fact["status"] = "pinned" if fact["pin"] else "unresolved"
                stale = True
            fact["candidates"] = [
                r for r in fact["candidates"] if r["id"] in authorized
            ]
            if (
                status
                and fact["status"] != status
                or project
                and fact["sourceRevision"]["projectId"] != project
            ):
                continue
            if offset <= matched < offset + limit:
                results.append(fact)
            matched += 1
        info = manifest["links"]
        caps = set(info["capsHit"])
        if stale:
            caps.add("stale_generation")
        result = {
            "generation": current["generation"],
            "links": results,
            "pagination": {
                "offset": offset,
                "limit": limit,
                "matched": matched,
                "hasMore": offset + len(results) < matched,
            },
            "coverage": {"complete": not caps, "capsHit": sorted(caps)},
            "untrustedContent": True,
        }
        while len(_json(result).encode()) > catalog.limits.response_bytes and results:
            results.pop()
            result["coverage"].update(
                complete=False, capsHit=sorted({*caps, "response_bytes"})
            )
            result["pagination"]["hasMore"] = True
        return result
