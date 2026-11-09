"""Engine-free, disk-backed hierarchy for bounded whole-workbook exploration.

Built only in a resource-isolated worker. Every view partitions the indexed
nodes into at most 32 children plus the outside group. Edge weights and internal
edge counts preserve the complete indexed topology, including cycles.
"""

from __future__ import annotations

import json
import math
import os
import re
from pathlib import Path

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

from linexcel.lazy_store import WorkbookIndex, _array_edge_filter, _edge_count

VERSION = 5
FANOUT = 32
OUTSIDE = -1


def generation(path: Path) -> str:
    stat = path.stat()
    return f"{stat.st_mtime_ns}:{stat.st_ctime_ns}:{stat.st_size}"


def _attach_path(path: Path) -> str:
    value = str(path.resolve())
    if os.name == "nt" and not value.startswith("\\\\?\\"):
        value = (
            "\\\\?\\UNC\\" + value[2:]
            if value.startswith("\\\\")
            else "\\\\?\\" + value
        )
    return value


def build_graph(index: Path, output: Path) -> dict:
    """Write a derived SQLite file; the coordinator publishes it atomically."""
    stamp = generation(index)
    with WorkbookIndex(index) as source, WorkbookIndex(output, writable=True) as target:
        db = target.db
        db.executescript("""
            PRAGMA journal_mode=OFF;
            PRAGMA temp_store=FILE;
            PRAGMA cache_size=-16384;
            CREATE TABLE nodes(id INTEGER PRIMARY KEY,node_id TEXT UNIQUE,
                               parent INTEGER,kind TEXT);
            CREATE INDEX nodes_parent ON nodes(parent);
            CREATE TABLE groups(id INTEGER PRIMARY KEY,parent INTEGER,
                                label TEXT,sheet TEXT,node_count INTEGER,
                                level TEXT,signature TEXT);
            CREATE INDEX groups_parent ON groups(parent);
            CREATE INDEX groups_sheet ON groups(sheet,level);
            CREATE TABLE members(view INTEGER,node INTEGER,child INTEGER,
                                 PRIMARY KEY(view,node)) WITHOUT ROWID;
            CREATE TABLE edges(view INTEGER,source INTEGER,target INTEGER,
                               weight INTEGER,PRIMARY KEY(view,source,target))
                               WITHOUT ROWID;
            CREATE TABLE metadata(data TEXT);
        """)
        next_group = -2

        def bounds(children):
            cells = [
                coordinate_to_tuple(cell)
                for child in children
                for cell in child[3:5]
                if re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", cell)
            ]
            if not cells:
                return (children[0][3], children[-1][4]) if children else ("", "")
            return (
                get_column_letter(min(c[1] for c in cells))
                + str(min(c[0] for c in cells)),
                get_column_letter(max(c[1] for c in cells))
                + str(max(c[0] for c in cells)),
            )

        def group(children, label, sheet=None, level="partition", signature=None):
            nonlocal next_group
            gid = next_group
            next_group -= 1
            count = sum(c[1] for c in children)
            db.execute(
                "INSERT INTO groups VALUES (?,NULL,?,?,?,?,?)",
                (gid, label[:256], sheet, count, level, signature),
            )
            db.executemany(
                "UPDATE groups SET parent=? WHERE id=?", [(gid, c[0]) for c in children]
            )
            first, last = bounds(children) if children else (label, label)
            return (
                gid,
                count,
                label[:256],
                first,
                last,
            )

        def pack(children, sheet):
            while len(children) > FANOUT:
                packed = []
                for i in range(0, len(children), FANOUT):
                    chunk = children[i : i + FANOUT]
                    if len(chunk) == 1:
                        packed.append(chunk[0])
                    else:
                        first, last = bounds(chunk)
                        packed.append(group(chunk, first + " … " + last, sheet))
                children = packed
            return children

        sheets, leaves, batch, patterns = [], [], [], []
        current_sheet = None
        current_pattern = None

        # The grouping counts are computed by SQLite, never by loading a full
        # workbook graph into the coordinator or browser. Keep the derived
        # table on disk even for hundreds of thousands of unique patterns.
        db.execute("ATTACH DATABASE ? AS original", (_attach_path(index),))
        db.executescript("""
            CREATE TABLE pattern_counts AS
            SELECT coalesce(sheet,'') sheet,coalesce(pattern_id,'') pattern_id,
                   count(*) size
            FROM original.nodes GROUP BY coalesce(sheet,''),coalesce(pattern_id,'');
            CREATE UNIQUE INDEX pattern_counts_key
            ON pattern_counts(sheet,pattern_id);
        """)

        def flush_leaf():
            if not batch:
                return
            first, last = batch[0][1], batch[-1][1]
            if current_sheet:
                first, last = first.rsplit("!", 1)[-1], last.rsplit("!", 1)[-1]
            label = first if len(batch) == 1 else first + " … " + last
            gid, *_ = group([], label, current_sheet)
            db.execute("UPDATE groups SET node_count=? WHERE id=?", (len(batch), gid))
            db.executemany(
                "INSERT INTO nodes VALUES (?,?,?,?)",
                [(seq, nid, gid, kind) for seq, nid, kind in batch],
            )
            leaves.append((gid, len(batch), label, first, last))
            batch.clear()

        def flush_sheet():
            flush_pattern()
            if not patterns:
                return
            # Pattern hashes are not spreadsheet positions. Sort their actual
            # bounding coordinates and propagate bounds over every child, not
            # merely the first/last hash bucket's (possibly unrelated) cells.
            patterns.sort(
                key=lambda c: (
                    coordinate_to_tuple(c[3])[::-1]
                    if re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", c[3])
                    else (0, 0)
                )
            )
            children = pack(patterns, current_sheet)
            item = group(children, current_sheet or "∅", current_sheet, "sheet")
            title = current_sheet or "∅"
            sheets.append((item[0], item[1], title, title, title))
            patterns.clear()

        def flush_pattern():
            flush_leaf()
            if not leaves:
                return
            signature = None
            if current_pattern:
                row = source.db.execute(
                    "SELECT data FROM patterns WHERE id=?", (current_pattern,)
                ).fetchone()
                if row:
                    signature = str(json.loads(row[0]).get("signature", ""))[:256]
            title = signature or ("◇ " + (current_sheet or "∅"))
            level = "pattern" if current_pattern else "sources"
            if len(leaves) == 1:
                # A bounded leaf already owns its cells. Reuse it as the
                # semantic group instead of introducing a one-child partition.
                item = leaves[0]
                db.execute(
                    "UPDATE groups SET label=?,level=?,signature=? WHERE id=?",
                    (title[:256], level, signature, item[0]),
                )
            else:
                item = group(
                    pack(leaves, current_sheet), title, current_sheet, level, signature
                )
            expected = db.execute(
                "SELECT size FROM pattern_counts WHERE sheet=? AND pattern_id=?",
                (current_sheet or "", current_pattern or ""),
            ).fetchone()[0]
            if item[1] != expected:
                raise ValueError("Formula group membership count differs from index")
            patterns.append(item)
            leaves.clear()

        for seq, nid, kind, sheet, pattern in source.db.execute(
            "SELECT seq,id,kind,coalesce(sheet,''),coalesce(pattern_id,'') "
            "FROM nodes ORDER BY sheet,coalesce(pattern_id,''),seq"
        ):
            if sheet != current_sheet:
                flush_sheet()
                current_sheet = sheet
                current_pattern = pattern
            elif pattern != current_pattern:
                flush_pattern()
                current_pattern = pattern
            batch.append((seq, nid, kind))
            if len(batch) == FANOUT:
                flush_leaf()
        flush_sheet()
        total = sum(s[1] for s in sheets)
        db.execute(
            "INSERT INTO groups VALUES (0,NULL,'Classeur',NULL,?,'workbook',NULL)",
            (total,),
        )
        db.executemany(
            "UPDATE groups SET parent=0 WHERE id=?",
            [(c[0],) for c in pack(sheets, None)],
        )
        # Integer membership keys avoid duplicating long Excel references at
        # every level. Recursive traversal and aggregation stay in SQLite.
        db.executescript("""
            INSERT INTO members
            WITH RECURSIVE m(view,node,child) AS (
                SELECT parent,id,id FROM nodes
                UNION ALL
                SELECT g.parent,m.node,g.id FROM m JOIN groups g ON g.id=m.view
                WHERE g.parent IS NOT NULL
            ) SELECT * FROM m;
            CREATE INDEX members_node ON members(node,view,child);
            CREATE TABLE raw_edges(source INTEGER,target INTEGER);
        """)
        db.execute(f"""
            INSERT INTO raw_edges SELECT a.id,b.id FROM original.edges e
            JOIN nodes a ON a.node_id=e.source JOIN nodes b ON b.node_id=e.target
            JOIN original.nodes t ON t.id=e.target
            WHERE {_array_edge_filter(source, schema="original")}
        """)
        edge_count = db.execute("SELECT count(*) FROM raw_edges").fetchone()[0]
        if edge_count != _edge_count(source):
            raise ValueError("Indexed edge endpoints are missing")
        db.executescript("""
            INSERT INTO edges
            SELECT a.view,a.child,coalesce(b.child,-1),count(*) FROM raw_edges e
            JOIN members a ON a.node=e.source
            LEFT JOIN members b ON b.node=e.target AND b.view=a.view
            GROUP BY a.view,a.child,coalesce(b.child,-1);
            INSERT INTO edges
            SELECT b.view,-1,b.child,count(*) FROM raw_edges e
            JOIN members b ON b.node=e.target
            LEFT JOIN members a ON a.node=e.source AND a.view=b.view
            WHERE a.node IS NULL GROUP BY b.view,b.child;
            DROP TABLE raw_edges;
            DROP TABLE members;
            DROP TABLE pattern_counts;
        """)
        meta = {
            "version": VERSION,
            "generation": stamp,
            "nodeCount": total,
            "edgeCount": edge_count,
            "maxViewNodes": FANOUT + 1,
            "hierarchy": "workbook/sheet/r1c1-pattern/cell",
            "indexStatus": source.meta.get("status", "complete"),
        }
        db.execute("INSERT INTO metadata VALUES (?)", (json.dumps(meta),))
        db.commit()
        db.execute("VACUUM")
    if generation(index) != stamp:
        raise ValueError("Index changed during graph construction")
    return meta


def read_view(
    path: Path, stamp: str, *, view=0, node_id=None, sheet=None, index_path=None
) -> dict:
    """Indexed, bounded reads only; never aggregate the workbook in HTTP."""
    with WorkbookIndex(path) as store:
        db = store.db
        meta = json.loads(db.execute("SELECT data FROM metadata").fetchone()[0])
        if meta["generation"] != stamp or meta["version"] != VERSION:
            raise ValueError("Graph generation is stale")
        if node_id:
            row = db.execute(
                "SELECT parent FROM nodes WHERE node_id=?", (node_id,)
            ).fetchone()
            if row is None:
                raise KeyError(node_id)
            view = row[0]
        elif sheet is not None:
            row = db.execute(
                "SELECT id FROM groups WHERE sheet=? AND level='sheet'", (sheet,)
            ).fetchone()
            if row is None:
                raise KeyError(sheet)
            view = row[0]
        group_row = db.execute(
            "SELECT parent,label,node_count FROM groups WHERE id=?", (view,)
        ).fetchone()
        if group_row is None:
            raise KeyError(view)
        parent, label, count = group_row
        nodes = [
            {
                "id": str(gid),
                "viewId": gid,
                "kind": "group",
                "label": name,
                "sheet": sheet,
                "nodeCount": size,
                "internalEdges": 0,
                "level": level,
                "signature": signature,
            }
            for gid, name, sheet, size, level, signature in db.execute(
                "SELECT id,label,sheet,node_count,level,signature FROM groups "
                "WHERE parent=? ORDER BY id DESC",
                (view,),
            )
        ]
        nodes += [
            {
                "id": str(nid),
                "nodeId": cell,
                "kind": kind,
                "label": cell,
                "nodeCount": 1,
                "internalEdges": 0,
            }
            for nid, cell, kind in db.execute(
                "SELECT id,node_id,kind FROM nodes WHERE parent=? ORDER BY id", (view,)
            )
        ]
        if index_path is not None:
            with WorkbookIndex(index_path) as original:
                for node in nodes:
                    if "nodeId" not in node:
                        continue
                    source = original[node["nodeId"]]
                    formula = source.get("sourceFormula") or source.get("formula") or ""
                    value = source.get("cachedValue")
                    if value is None:
                        value = source.get("value")
                    node.update(
                        formula=str(formula)[:512],
                        value=value[:256] if isinstance(value, str) else value,
                        valueSource=source.get("valueSource", "unknown"),
                        contentTruncated=len(str(formula)) > 512
                        or (isinstance(value, str) and len(value) > 256),
                    )
        if parent is not None:
            nodes.append(
                {
                    "id": str(OUTSIDE),
                    "kind": "outside",
                    "viewId": parent,
                    "label": "↗",
                    "nodeCount": meta["nodeCount"] - count,
                    "internalEdges": 0,
                }
            )
        by_id = {n["id"]: n for n in nodes}
        edges, incident = [], 0
        for a, b, weight in db.execute(
            "SELECT source,target,weight FROM edges WHERE view=?", (view,)
        ):
            incident += weight
            if a == b:
                by_id[str(a)]["internalEdges"] += weight
            else:
                edges.append({"source": str(a), "target": str(b), "weight": weight})
        if str(OUTSIDE) in by_id:
            by_id[str(OUTSIDE)]["internalEdges"] = meta["edgeCount"] - incident
        # Deterministic bounded layout computed on the server. Exact topology,
        # not a claim that position is a dependency rank (cycles are retained).
        columns = max(1, math.ceil(math.sqrt(len(nodes))))
        for i, node in enumerate(nodes):
            node["position"] = {"x": (i % columns) * 270, "y": (i // columns) * 125}
        breadcrumbs = []
        ancestor = view
        while ancestor is not None:
            p, title, level, size = db.execute(
                "SELECT parent,label,level,node_count FROM groups WHERE id=?",
                (ancestor,),
            ).fetchone()
            breadcrumbs.append(
                {"viewId": ancestor, "label": title, "level": level, "nodeCount": size}
            )
            ancestor = p
        return {
            "viewId": view,
            "parent": parent,
            "label": label,
            "focusedNodeId": node_id,
            "scopeNodeCount": count,
            "nodes": nodes,
            "edges": edges,
            "breadcrumbs": breadcrumbs[::-1],
            "meta": meta,
        }
