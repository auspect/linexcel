"""Engine-free, disk-backed hierarchy for bounded whole-workbook exploration.

Built only in a resource-isolated worker. Every view partitions the indexed
nodes into at most 32 children plus the outside group. Edge weights and internal
edge counts preserve the complete indexed topology, including cycles.
"""

from __future__ import annotations

import json
import math
import os
from pathlib import Path

from linexcel.lazy_store import WorkbookIndex

VERSION = 1
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
                                label TEXT,sheet TEXT,node_count INTEGER);
            CREATE INDEX groups_parent ON groups(parent);
            CREATE TABLE members(view INTEGER,node INTEGER,child INTEGER,
                                 PRIMARY KEY(view,node)) WITHOUT ROWID;
            CREATE TABLE edges(view INTEGER,source INTEGER,target INTEGER,
                               weight INTEGER,PRIMARY KEY(view,source,target))
                               WITHOUT ROWID;
            CREATE TABLE metadata(data TEXT);
        """)
        next_group = -2

        def group(children, label, sheet=None):
            nonlocal next_group
            gid = next_group
            next_group -= 1
            count = sum(c[1] for c in children)
            db.execute(
                "INSERT INTO groups VALUES (?,NULL,?,?,?)",
                (gid, label[:256], sheet, count),
            )
            db.executemany(
                "UPDATE groups SET parent=? WHERE id=?", [(gid, c[0]) for c in children]
            )
            return (
                gid,
                count,
                label[:256],
                children[0][3] if children else label,
                children[-1][4] if children else label,
            )

        def pack(children, sheet):
            while len(children) > FANOUT:
                children = [
                    group(
                        children[i : i + FANOUT],
                        children[i][3]
                        + " … "
                        + children[min(i + FANOUT, len(children)) - 1][4],
                        sheet,
                    )
                    for i in range(0, len(children), FANOUT)
                ]
            return children

        sheets, leaves, batch = [], [], []
        current_sheet = None

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
            flush_leaf()
            if not leaves:
                return
            children = pack(leaves, current_sheet)
            if len(children) == 1:
                item = children[0]
                db.execute(
                    "UPDATE groups SET label=? WHERE id=?",
                    (current_sheet or "∅", item[0]),
                )
            else:
                item = group(children, current_sheet or "∅", current_sheet)
            title = current_sheet or "∅"
            sheets.append((item[0], item[1], title, title, title))
            leaves.clear()

        for seq, nid, kind, sheet in source.db.execute(
            "SELECT seq,id,kind,coalesce(sheet,'') FROM nodes ORDER BY sheet,seq"
        ):
            if sheet != current_sheet:
                flush_sheet()
                current_sheet = sheet
            batch.append((seq, nid, kind))
            if len(batch) == FANOUT:
                flush_leaf()
        flush_sheet()
        total = sum(s[1] for s in sheets)
        db.execute("INSERT INTO groups VALUES (0,NULL,'Classeur',NULL,?)", (total,))
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
        db.execute("ATTACH DATABASE ? AS original", (_attach_path(index),))
        db.execute("""
            INSERT INTO raw_edges SELECT a.id,b.id FROM original.edges e
            JOIN nodes a ON a.node_id=e.source JOIN nodes b ON b.node_id=e.target
        """)
        edge_count = db.execute("SELECT count(*) FROM raw_edges").fetchone()[0]
        if edge_count != source.db.execute("SELECT count(*) FROM edges").fetchone()[0]:
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
        """)
        meta = {
            "version": VERSION,
            "generation": stamp,
            "nodeCount": total,
            "edgeCount": edge_count,
            "maxViewNodes": FANOUT + 1,
            "indexStatus": source.meta.get("status", "complete"),
        }
        db.execute("INSERT INTO metadata VALUES (?)", (json.dumps(meta),))
        db.commit()
        db.execute("VACUUM")
    if generation(index) != stamp:
        raise ValueError("Index changed during graph construction")
    return meta


def read_view(path: Path, stamp: str, *, view=0, node_id=None) -> dict:
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
            }
            for gid, name, sheet, size in db.execute(
                "SELECT id,label,sheet,node_count FROM groups "
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
            p, title = db.execute(
                "SELECT parent,label FROM groups WHERE id=?", (ancestor,)
            ).fetchone()
            breadcrumbs.append({"viewId": ancestor, "label": title})
            ancestor = p
        return {
            "viewId": view,
            "parent": parent,
            "label": label,
            "scopeNodeCount": count,
            "nodes": nodes,
            "edges": edges,
            "breadcrumbs": breadcrumbs[::-1],
            "meta": meta,
        }
