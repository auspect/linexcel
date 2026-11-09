"""Bounded directed lineage over immutable, owner-authorized revisions."""

from __future__ import annotations

import json
import sqlite3
import time
from collections import deque

from linexcel.catalog import IndexPool, _json, cell_handle, parse_handle, revision
from linexcel.limits import validate_limits
from linexcel.link_facts import FACTS_VERSION
from linexcel.refs import Rect, a1, parse_ref
from linexcel.values import _values_differ

TRACE_VERSION = "linexcel.lineage.v1"
SCOPE = "loaded owner revisions only; never-loaded workbooks are not visible"
BUDGETS = {
    "depth",
    "node_budget",
    "edge_budget",
    "workbook_budget",
    "time_budget",
    "response_bytes",
    "link_budget",
    "link_build_time",
    "index_rebuild_required",
    "source_index_incomplete",
    "catalog_bytes",
}


def components(nodes, edges):
    """Iterative Kosaraju SCCs; neither a workbook cycle nor depth uses recursion."""
    forward = {n: [] for n in nodes}
    reverse = {n: [] for n in nodes}
    for edge in edges:
        forward[edge["source"]].append(edge["target"])
        reverse[edge["target"]].append(edge["source"])
    visited, order = set(), []
    for root in forward:
        if root in visited:
            continue
        stack = [(root, False)]
        while stack:
            node, finish = stack.pop()
            if finish:
                order.append(node)
            elif node not in visited:
                visited.add(node)
                stack.append((node, True))
                stack.extend((n, False) for n in forward[node] if n not in visited)
    visited, cycles = set(), []
    for root in reversed(order):
        if root in visited:
            continue
        stack, group = [root], []
        while stack:
            node = stack.pop()
            if node not in visited:
                visited.add(node)
                group.append(node)
                stack.extend(reverse[node])
        if len(group) > 1 or root in forward[root]:
            cycles.append(sorted(group))
    return cycles


class Traversal:
    def __init__(
        self,
        catalog,
        sources,
        db,
        pool,
        manifest,
        *,
        direction,
        depth,
        limit,
        edges,
        workbooks,
        deadline,
    ):
        self.catalog, self.db, self.pool = catalog, db, pool
        self.sources = {s.project_id: s for s in sources}
        self.revisions = {r["projectId"]: r for r in manifest if r["revision"]}
        self.direction, self.depth = direction, depth
        self.limit, self.edge_limit, self.workbook_limit = limit, edges, workbooks
        self.deadline = deadline
        self.nodes, self.raw, self.edges = {}, {}, {}
        self.boundaries, self.boundary_keys = [], set()
        self.workbooks, self.queue = set(), deque()
        self.caps = set()
        self.active = False

    def boundary(self, node, reason, **detail):
        key = (node, reason, _json(detail))
        if key not in self.boundary_keys:
            self.boundary_keys.add(key)
            self.boundaries.append({"node": node, "reason": reason, **detail})
        if reason in BUDGETS:
            self.caps.add(reason)

    def index(self, pid):
        index = self.pool.get(self.sources[pid].path)
        if index.meta.get("sha256") != self.revisions[pid]["revision"]["sha256"]:
            raise KeyError("Revision changed during lineage")
        if self.active:
            index.db.set_progress_handler(
                lambda: int(time.monotonic() >= self.deadline), 100
            )
        return index

    def add(self, pid, nid, level, *, synthetic=None, parent=None):
        rev = self.revisions[pid]["revision"]
        key = cell_handle(rev, nid)
        if key in self.nodes:
            return key
        if len(self.nodes) >= self.limit:
            self.boundary(parent, "node_budget")
            return None
        if pid not in self.workbooks and len(self.workbooks) >= self.workbook_limit:
            self.boundary(parent, "workbook_budget")
            return None
        index = self.index(pid)
        raw = synthetic or index.get(nid) or index.implicit_array_member(nid)
        if raw is None:
            self.boundary(parent, "source_unavailable", revision=rev, address=nid)
            return None
        self.workbooks.add(pid)
        self.raw[key] = raw
        self.nodes[key] = {
            "id": key,
            "handle": key,
            "sourceNode": nid,
            "revision": rev,
            "filename": self.sources[pid].filename,
            "sheet": raw.get("sheet"),
            "address": raw.get("cell") or raw.get("reference"),
            "kind": raw["kind"],
            "labels": raw.get("cellLabels", []),
            "value": raw.get("value"),
            "valueProvenance": raw.get("provenance", "unknown"),
            "cacheStatus": raw.get("cache_status"),
            "formula": raw.get("formula"),
            "bounds": raw.get("bounds"),
            "implicitBlanks": bool(raw.get("implicitBlanks")),
        }
        if raw.get("cell") and not raw.get("external"):
            self.nodes[key]["labels"] = index.cell_labels(
                raw["sheet"], raw["row"], raw["column"]
            )
        self.queue.append((key, pid, nid, level))
        return key

    def edge(self, source, target, kind, **details):
        key = (source, target, kind, details.get("linkId"))
        if key in self.edges:
            return
        if len(self.edges) >= self.edge_limit:
            self.boundary(
                target if self.direction == "upstream" else source, "edge_budget"
            )
            return
        self.edges[key] = {"source": source, "target": target, "kind": kind, **details}

    def synthetic_range(self, pid, sheet, bounds):
        rect = Rect(sheet, *bounds)
        nid = rect.to_a1()
        return nid, {
            "id": nid,
            "kind": "range",
            "sheet": sheet,
            "reference": a1(rect.r1, rect.c1) + ":" + a1(rect.r2, rect.c2),
            "bounds": bounds,
            "provenance": "source_range",
            "implicitBlanks": True,
        }

    def external(self, key, fact, level):
        if not fact["supported"]:
            self.boundary(
                key,
                "unsupported_reference",
                reference=fact["reference"],
                referenceType=fact["referenceType"],
                linkId=fact["id"],
            )
            return
        target = fact["targetRevision"]
        if not target:
            reason = (
                "ambiguous_external"
                if fact["status"] == "ambiguous"
                else "unresolved_external"
            )
            self.boundary(
                key,
                reason,
                linkId=fact["id"],
                relationshipPath=fact["relationshipPath"],
                stalePin=fact["stalePin"],
            )
            if fact["cached"]:
                self.boundary(
                    key,
                    "cache_only",
                    linkId=fact["id"],
                    cached=fact["cached"],
                    label="cached value, source unavailable",
                )
            return
        if level >= self.depth:
            self.boundary(key, "depth")
            return
        if len(self.edges) >= self.edge_limit:
            self.boundary(key, "edge_budget")
            return
        pid = target["projectId"]
        if pid not in self.workbooks:
            if len(self.workbooks) >= self.workbook_limit:
                self.boundary(key, "workbook_budget")
                return
            if len(self.nodes) >= self.limit:
                self.boundary(key, "node_budget")
                return
        index = self.index(pid)
        sheets = {s.lower(): s for s in index.meta.get("sheets", [])}
        sheet = sheets.get((fact["sheet"] or "").lower())
        if sheet is None:
            self.boundary(
                key, "source_unavailable", linkId=fact["id"], sheet=fact["sheet"]
            )
            if fact["cached"]:
                self.boundary(
                    key,
                    "cache_only",
                    linkId=fact["id"],
                    cached=fact["cached"],
                    label="cached value, source unavailable",
                )
            return
        r1, c1, r2, c2 = fact["bounds"]
        diagnostic = []
        # Diagnostic is bounded by the node budget, never used for identity.
        for cached in fact["cached"][: self.limit]:
            row = index.db.execute(
                "SELECT id FROM nodes WHERE sheet=? AND row=? AND col=?",
                (sheet, cached["row"], cached["column"]),
            ).fetchone()
            current = index[row[0]].get("value") if row else None
            old = cached["value"]
            agreement = (
                None
                if current is None
                else (
                    not _values_differ(current, old, None)
                    if isinstance(current, (int, float, bool))
                    and isinstance(old, (int, float, bool))
                    else type(current) is type(old) and current == old
                )
            )
            diagnostic.append(
                {
                    "address": a1(cached["row"], cached["column"]),
                    "cached": old,
                    "current": current,
                    "agreement": agreement,
                }
            )
        if (r1, c1) == (r2, c2):
            nid = Rect(sheet, r1, c1, r2, c2).to_a1()
            raw = None
            if nid not in index:
                # A verified missing stored cell is blank, unless its sheet is partial.
                if any(
                    s["name"] == sheet and s["status"] != "complete"
                    for s in index.meta.get("sheetDetails", [])
                ):
                    self.boundary(key, "source_unavailable", linkId=fact["id"])
                    return
                raw = index.implicit_array_member(nid) or {
                    "id": nid,
                    "kind": "input",
                    "sheet": sheet,
                    "cell": a1(r1, c1),
                    "row": r1,
                    "column": c1,
                    "value": None,
                    "provenance": "saved_blank",
                }
        else:
            nid, raw = self.synthetic_range(pid, sheet, fact["bounds"])
        child = self.add(pid, nid, level + 1, synthetic=raw, parent=key)
        if child:
            self.edge(
                child,
                key,
                "external_link",
                linkId=fact["id"],
                resolutionStatus=fact["status"],
                relationshipPath=fact["relationshipPath"],
                versionCaveat=fact["versionCaveat"],
                cacheDiagnostic=diagnostic,
                cacheDiagnosticsOmitted=max(0, len(fact["cached"]) - len(diagnostic)),
            )

    def expand(self, key, pid, nid, level):
        if time.monotonic() >= self.deadline:
            self.boundary(key, "time_budget")
            return
        raw = self.raw[key]
        index = self.index(pid)
        kind = raw["kind"]
        if any(
            s["name"] == raw.get("sheet") and s["status"] != "complete"
            for s in index.meta.get("sheetDetails", [])
        ):
            self.boundary(key, "source_index_incomplete")
        if kind in {"macro", "vba", "vba_function", "udf"}:
            self.boundary(key, "macro")
            return
        if kind in {"power_query", "query", "connection"}:
            self.boundary(key, "power_query")
            return
        for destination in index.meta.get("powerQueryDestinations", []):
            rect = parse_ref(
                destination.get("ref") or "", default_sheet=destination["sheet"]
            )
            if (
                rect
                and raw.get("row") is not None
                and rect.sheet == raw.get("sheet")
                and rect.r1 <= raw["row"] <= rect.r2
                and rect.c1 <= raw["column"] <= rect.c2
            ):
                self.boundary(
                    key,
                    "power_query",
                    query=destination["query"],
                    table=destination["table"],
                )
                if self.direction == "upstream":
                    return
        if raw.get("external"):
            self.boundary(key, "cache_only", label="cached value, source unavailable")
            return
        if kind == "unresolved":
            self.boundary(key, "unsupported_reference", reference=raw.get("reference"))
            return
        if raw.get("capabilityDiagnostics") or any(
            "dynamic" in str(d).lower() or "tokenization failed" in str(d).lower()
            for d in raw.get("diagnostics", [])
        ):
            self.boundary(
                key,
                "unsupported_reference",
                referenceType="index_capability_or_unparsed",
            )
        if self.direction == "upstream":
            facts = [
                json.loads(row[0])
                for row in self.db.execute(
                    "SELECT data FROM links WHERE source_project=? AND source_sha=? "
                    "AND source_node=? ORDER BY id LIMIT ?",
                    (
                        pid,
                        self.revisions[pid]["revision"]["sha256"],
                        nid,
                        self.edge_limit + 1,
                    ),
                )
            ]
            if len(facts) > self.edge_limit:
                self.boundary(key, "edge_budget")
            placeholders = {f["placeholder"] for f in facts}
            for fact in facts:
                self.external(key, fact, level)
            index = self.index(pid)
            neighbors = [
                (row[0], "formula")
                for row in index.db.execute(
                    "SELECT source FROM edges WHERE target=? ORDER BY source LIMIT ?",
                    (nid, self.edge_limit + 1),
                )
                if row[0] not in placeholders
            ]
            if kind == "array_result" and raw.get("arrayAnchor"):
                if raw["arrayAnchor"] not in {n for n, _ in neighbors}:
                    neighbors.append((raw["arrayAnchor"], "formula"))
            if kind == "range" and raw.get("bounds"):
                neighbors += [
                    (row[0], "range")
                    for row in index.db.execute(
                        "SELECT id FROM nodes WHERE sheet=? AND row BETWEEN ? AND ? "
                        "AND col BETWEEN ? AND ? "
                        "AND kind IN ('cell','input','array_result') "
                        "ORDER BY row,col LIMIT ?",
                        (
                            raw["sheet"],
                            raw["bounds"][0],
                            raw["bounds"][2],
                            raw["bounds"][1],
                            raw["bounds"][3],
                            self.limit + 1,
                        ),
                    )
                ]
                known = {n for n, _ in neighbors}
                neighbors += [
                    (anchor, "formula")
                    for anchor in index.array_anchors(
                        raw["sheet"], raw["bounds"], limit=self.edge_limit + 1
                    )
                    if anchor not in known
                ]
            if kind == "input" and not neighbors and not facts:
                self.boundary(key, "input_constant")
        else:
            neighbors = [
                (row[0], "formula")
                for row in index.db.execute(
                    "SELECT target FROM edges WHERE source=? ORDER BY target LIMIT ?",
                    (nid, self.edge_limit + 1),
                )
            ]
            if bounds := raw.get("arrayBounds"):
                extent = Rect(raw["sheet"], *bounds)
                if extent.ncells <= self.limit:
                    for r in range(extent.r1, extent.r2 + 1):
                        for c in range(extent.c1, extent.c2 + 1):
                            member = Rect(raw["sheet"], r, c, r, c).to_a1()
                            if member not in index:
                                candidate = index.implicit_array_member(member)
                                if candidate and candidate.get("arrayAnchor") == nid:
                                    neighbors.append((member, "formula"))
                                elif candidate:
                                    self.boundary(
                                        key, "unsupported_reference", reference=member
                                    )
                else:
                    stored = index.db.execute(
                        "SELECT count(*) FROM nodes WHERE sheet=? "
                        "AND row BETWEEN ? AND ? AND col BETWEEN ? AND ? "
                        "AND kind IN ('cell','input','array_result')",
                        (extent.sheet, extent.r1, extent.r2, extent.c1, extent.c2),
                    ).fetchone()[0]
                    if stored < extent.ncells:
                        # Never enumerate an enormous implicit output rectangle
                        # or certify its unvisited downstream members as complete.
                        self.boundary(
                            key, "depth" if level >= self.depth else "node_budget"
                        )
            if raw.get("row") is not None and index.meta.get(
                "externalFactsVersion"
            ) in {1, FACTS_VERSION}:
                neighbors += [
                    (row[0], "range")
                    for row in index.db.execute(
                        "SELECT node FROM lineage_ranges WHERE sheet=? "
                        "AND ? BETWEEN r1 AND r2 AND ? BETWEEN c1 AND c2 "
                        "ORDER BY node LIMIT ?",
                        (
                            raw["sheet"].lower(),
                            raw["row"],
                            raw["column"],
                            self.edge_limit + 1,
                        ),
                    )
                ]
            # Membership propagates the selected cell's influence only to this
            # range's exact local consumers. Its other members have not been
            # reached. Cell lookups already include every overlapping external
            # operand and uncertainty; only a range root selects its whole span.
            if kind != "range" or level == 0:
                self.reverse_external(key, pid, raw, level)
        for neighbor, edge_kind in neighbors:
            if time.monotonic() >= self.deadline:
                self.boundary(key, "time_budget")
                break
            if level >= self.depth:
                self.boundary(key, "depth")
                break
            if len(self.edges) >= self.edge_limit:
                self.boundary(key, "edge_budget")
                break
            candidate = self.index(pid).get(neighbor)
            if candidate and candidate.get("external"):
                # Legacy indexes and capped link generations may retain copied
                # external cache nodes. They are never source-revision edges.
                self.boundary(
                    key, "unresolved_external", reference=candidate.get("reference")
                )
                if candidate.get("value") is not None:
                    self.boundary(
                        key,
                        "cache_only",
                        cached=candidate["value"],
                        label="cached value, source unavailable",
                    )
                continue
            child = self.add(pid, neighbor, level + 1, parent=key)
            if child:
                self.edge(
                    child, key, edge_kind
                ) if self.direction == "upstream" else self.edge(key, child, edge_kind)

    def reverse_external(self, key, pid, raw, level):
        if not raw.get("sheet"):
            return
        rect = parse_ref(
            raw.get("cell") or raw.get("reference") or "", default_sheet=raw["sheet"]
        )
        if rect is None:
            return
        self.possible_unresolved_readers(key, rect)
        uncertain = self.db.execute(
            "SELECT l.data FROM links l JOIN link_candidates c ON c.link=l.id "
            "WHERE c.project=? AND c.sha=? "
            "AND (l.target_project IS NULL OR l.r1 IS NULL) "
            "ORDER BY l.id LIMIT ?",
            (pid, self.revisions[pid]["revision"]["sha256"], self.edge_limit + 1),
        ).fetchall()
        if len(uncertain) > self.edge_limit:
            self.boundary(key, "edge_budget")
        for (data,) in uncertain:
            fact = json.loads(data)
            bounds = fact["bounds"]
            if bounds and (
                fact["sheet"].lower() != raw["sheet"].lower()
                or not Rect(raw["sheet"], *bounds).intersects(rect)
            ):
                continue
            if not fact["supported"]:
                self.boundary(
                    key,
                    "unsupported_reference",
                    linkId=fact["id"],
                    referenceType=fact["referenceType"],
                )
            elif fact["status"] == "ambiguous":
                self.boundary(
                    key,
                    "ambiguous_external",
                    linkId=fact["id"],
                    role="potential_reader",
                )
        facts = [
            json.loads(row[0])
            for row in self.db.execute(
                "SELECT data FROM links WHERE target_project=? AND target_sha=? "
                "AND target_sheet=? AND r1<=? AND r2>=? AND c1<=? AND c2>=? "
                "ORDER BY id LIMIT ?",
                (
                    pid,
                    self.revisions[pid]["revision"]["sha256"],
                    raw["sheet"].lower(),
                    rect.r2,
                    rect.r1,
                    rect.c2,
                    rect.c1,
                    self.edge_limit + 1,
                ),
            )
        ]
        for fact in facts:
            if level >= self.depth:
                self.boundary(key, "depth")
                break
            if len(self.edges) >= self.edge_limit:
                self.boundary(key, "edge_budget")
                break
            child = self.add(
                fact["sourceRevision"]["projectId"],
                fact["sourceNode"],
                level + 1,
                parent=key,
            )
            if child:
                self.edge(
                    key,
                    child,
                    "external_link",
                    linkId=fact["id"],
                    resolutionStatus=fact["status"],
                    relationshipPath=fact["relationshipPath"],
                    versionCaveat=fact["versionCaveat"],
                )

    def possible_unresolved_readers(self, key, rect):
        # Filename mismatches cannot exclude a renamed/moved target. Only known
        # sheet/rectangle disjointness can exclude an identity-unresolved link.
        # Unsupported forms have no verified rectangle and remain possible.
        predicate = (
            "target_project IS NULL AND (r1 IS NULL OR (target_sheet=? "
            "AND r1<=? AND r2>=? AND c1<=? AND c2>=?))"
        )
        params = (rect.sheet.lower(), rect.r2, rect.r1, rect.c2, rect.c1)
        count = self.db.execute(
            "SELECT count(*) FROM links WHERE " + predicate, params
        ).fetchone()[0]
        if not count:
            return
        samples = []
        try:
            sample_limit = self.catalog.limits.lineage_reader_samples
            if sample_limit:
                samples = [
                    row[0]
                    for row in self.db.execute(
                        "SELECT id FROM links WHERE "
                        + predicate
                        + " ORDER BY id LIMIT ?",
                        (*params, sample_limit),
                    )
                ]
        finally:
            # Preserve a completed count even if the sample scan times out.
            self.boundary(
                key,
                "possible_unresolved_reader",
                count=count,
                linkIds=samples,
                sampleOmitted=count - len(samples),
            )


def directed_lineage(
    catalog,
    sources,
    handle,
    *,
    direction,
    depth=1,
    limit=200,
    edges=None,
    workbooks=None,
    milliseconds=None,
):
    started = time.monotonic()
    if direction not in {"upstream", "downstream"}:
        raise ValueError("direction must be upstream or downstream")
    budgets = {
        "depth": depth,
        "limit": limit,
        "edges": catalog.limits.lineage_edges if edges is None else edges,
        "workbooks": catalog.limits.lineage_workbooks
        if workbooks is None
        else workbooks,
        "milliseconds": catalog.limits.lineage_ms
        if milliseconds is None
        else milliseconds,
    }
    validate_limits(**budgets)
    maxima = {
        "depth": catalog.limits.lineage_depth,
        "limit": catalog.limits.lineage_nodes,
        "edges": catalog.limits.lineage_edges,
        "workbooks": catalog.limits.lineage_workbooks,
        "milliseconds": catalog.limits.lineage_ms,
    }
    if (
        not limit
        or not budgets["workbooks"]
        or any(v > maxima[k] for k, v in budgets.items())
    ):
        raise ValueError("Lineage budget exceeded")
    pid, sha, nid = parse_handle(handle)
    if pid not in {s.project_id for s in sources}:
        raise KeyError("Cell not found")
    catalog.ensure(sources)
    # Build/reconciliation is outside traversal time, separately disclosed.
    ready = time.monotonic()
    with (
        catalog.source_reads,
        catalog.pin() as (db, current),
        IndexPool(max(1, catalog.limits.connections - 1)) as pool,
    ):
        manifest = catalog._manifest(sources, pool)
        revisions = {r["projectId"]: r for r in manifest if r["revision"]}
        if pid not in revisions or revisions[pid]["revision"]["sha256"] != sha:
            raise KeyError("Cell revision not found")
        traversal = Traversal(
            catalog,
            sources,
            db,
            pool,
            manifest,
            direction=direction,
            depth=depth,
            limit=limit,
            edges=budgets["edges"],
            workbooks=budgets["workbooks"],
            deadline=ready + budgets["milliseconds"] / 1000,
        )
        index = traversal.index(pid)
        root = None
        if nid not in index:
            rect = parse_ref(nid)
            if (
                not rect
                or rect.sheet not in index.meta.get("sheets", [])
                or rect.to_a1() != nid
            ):
                raise KeyError("Cell not found")
            if rect.ncells > 1:
                _, root = traversal.synthetic_range(
                    pid, rect.sheet, [rect.r1, rect.c1, rect.r2, rect.c2]
                )
            elif all(
                s["status"] == "complete"
                for s in index.meta.get("sheetDetails", [])
                if s["name"] == rect.sheet
            ):
                root = index.implicit_array_member(nid) or {
                    "id": nid,
                    "kind": "input",
                    "sheet": rect.sheet,
                    "cell": a1(rect.r1, rect.c1),
                    "row": rect.r1,
                    "column": rect.c1,
                    "value": None,
                    "provenance": "saved_blank",
                }
            else:
                raise KeyError("Cell source unavailable")
        traversal.add(pid, nid, 0, synthetic=root)
        traversal.active = True
        snapshot = json.loads(
            (catalog.directory / f"{current['generation']}.json").read_text("utf-8")
        )
        for cap in snapshot["links"]["capsHit"]:
            traversal.boundary(
                handle,
                "index_rebuild_required"
                if cap == "external_facts_rebuild_required"
                else cap,
            )

        # SQLite progress callbacks bound scans even when no rows match.
        def progress():
            return int(time.monotonic() >= traversal.deadline)

        db.set_progress_handler(progress, 100)
        try:
            while traversal.queue:
                key, source_pid, node, level = traversal.queue.popleft()
                index = traversal.index(source_pid)
                index.db.set_progress_handler(progress, 100)
                try:
                    traversal.expand(key, source_pid, node, level)
                finally:
                    # The original connection may have been evicted while crossing.
                    for opened in pool.indexes.values():
                        opened.db.set_progress_handler(None, 0)
                if "time_budget" in traversal.caps:
                    for pending, *_ in traversal.queue:
                        traversal.boundary(pending, "time_budget")
                    break
        except sqlite3.OperationalError as error:
            if "interrupt" not in str(error).lower():
                raise
            traversal.boundary(handle, "time_budget")
        finally:
            db.set_progress_handler(None, 0)
        cycles = components(traversal.nodes, list(traversal.edges.values()))
        for group in cycles:
            for node in group:
                traversal.boundary(node, "cycle", component=group)
        result = {
            "revision": revision(pid, sha),
            "filename": traversal.sources[pid].filename,
            "direction": direction,
            "nodes": list(traversal.nodes.values()),
            "edges": list(traversal.edges.values()),
            "boundaries": traversal.boundaries,
            "stronglyConnectedComponents": cycles,
            "revisions": [],
            "coverage": {},
            "trace": {},
            "untrustedContent": True,
        }

        def summarize():
            caps = sorted(traversal.caps)
            exhaustive = not caps
            reasons = sorted({b["reason"] for b in result["boundaries"]})
            complete = exhaustive and not (set(reasons) - {"input_constant"})
            result["revisions"] = sorted(
                [traversal.revisions[p]["revision"] for p in traversal.workbooks],
                key=lambda r: r["id"],
            )
            result["coverage"] = {
                "complete": complete,
                "exhaustive": exhaustive,
                "scope": SCOPE,
                "nodesReturned": len(result["nodes"]),
                "edgesReturned": len(result["edges"]),
                "depth": depth,
                "capsHit": caps,
                "boundaries": reasons,
            }
            result["trace"] = {
                "schemaVersion": TRACE_VERSION,
                "tool": "directed_lineage",
                "direction": direction,
                "generation": current["generation"],
                "root": handle,
                "scope": SCOPE,
                "revisions": result["revisions"],
                "limits": {
                    **budgets,
                    "unresolvedReaderSamples": catalog.limits.lineage_reader_samples,
                },
                "counts": {
                    "nodes": len(result["nodes"]),
                    "edges": len(result["edges"]),
                    "workbooks": len(result["revisions"]),
                    "indexConnectionsPeak": pool.peak,
                },
                "complete": complete,
                "exhaustive": exhaustive,
                "capsHit": caps,
                "preparationMs": round((ready - started) * 1000, 3),
                "durationMs": round((time.monotonic() - ready) * 1000, 3),
            }

        summarize()
        while len(_json(result).encode()) > catalog.limits.response_bytes:
            traversal.boundary(handle, "response_bytes")
            if len(result["nodes"]) > 1:
                result["nodes"].pop()
                ids = {n["id"] for n in result["nodes"]}
                result["edges"] = [
                    e
                    for e in result["edges"]
                    if e["source"] in ids and e["target"] in ids
                ]
                result["boundaries"] = [
                    b for b in traversal.boundaries if b["node"] in ids
                ]
                result["stronglyConnectedComponents"] = components(ids, result["edges"])
                for boundary in result["boundaries"]:
                    if boundary["reason"] == "cycle":
                        original = boundary["component"]
                        boundary["component"] = [n for n in original if n in ids]
                        boundary["componentNodesOmitted"] = (
                            boundary.get("componentNodesOmitted", 0)
                            + len(original)
                            - len(boundary["component"])
                        )
            elif result["edges"]:
                result["edges"].pop()
            else:
                raise ValueError(
                    "response_bytes: cell evidence exceeds response budget"
                )
            summarize()
        return result
