"""Measure the data node prototype on one workbook; prints and writes JSON.

usage:
  bench.py build <xlsx> <store.sqlite> <out.json>
      stream the workbook with serve's reader into the value store, detect
      regions, describe them; peak memory in a fresh process
  bench.py windows <store.sqlite> <sheet> <out.json>
  bench.py link <store.sqlite> <index.sqlite> <out.json>
  bench.py verify <store.sqlite> <index.sqlite> <out.json> [samples]
      compare a reproducible sample of region cells with serve's index nodes
  bench.py evaluate <xlsx> <store.sqlite> <index.sqlite> <cell> <out.json>
      serve's targeted calculation, then the same with region batch loading
"""

import json
import random
import sqlite3
import sys
import time
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import data_node  # noqa: E402


def peak():
    try:
        sys.path.insert(0, str(HERE / "probes"))
        from winmem import peak_mb

        return peak_mb()
    except Exception as exc:  # report, do not hide
        return {"peak": f"unavailable: {exc}"}


def write(out, result):
    Path(out).write_text(
        json.dumps(result, indent=1, ensure_ascii=False, default=str), "utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, default=str)[:4000])


def build(xlsx, store, out):
    result = {
        "workbook": Path(xlsx).name,
        "build": data_node.build_store(Path(xlsx), Path(store)),
    }
    started = time.perf_counter()
    regions = data_node.detect_regions(Path(store))
    result["detectAndDescribeSeconds"] = round(time.perf_counter() - started, 2)
    db = sqlite3.connect(store)
    result["sheets"] = [
        {"name": n, "width": w, "literalCells": lc, "formulaCells": fc}
        for n, w, lc, fc in db.execute(
            "SELECT name, width, literal_cells, formula_cells FROM sheets"
        )
    ]
    result["exceptions"] = dict(
        db.execute("SELECT kind, count(*) FROM exceptions GROUP BY kind").fetchall()
    )
    result["wideCells"] = db.execute("SELECT count(*) FROM wide").fetchone()[0]
    try:
        result["tableMB"] = {
            n: round(s / 1e6, 1)
            for n, s in db.execute(
                "SELECT name, sum(pgsize) FROM dbstat GROUP BY name "
                "ORDER BY 2 DESC LIMIT 8"
            )
        }
    except sqlite3.Error as exc:
        result["tableMB"] = str(exc)
    db.close()
    literal = sum(s["literalCells"] for s in result["sheets"])
    result["literalCells"] = literal
    result["storeBytesPerLiteralCell"] = round(
        Path(store).stat().st_size / max(literal, 1), 1
    )
    result["usPerLiteralCell"] = round(
        result["build"]["seconds"] * 1e6 / max(literal, 1), 1
    )
    result["regions"] = [
        {
            k: r[k]
            for k in (
                "id",
                "kind",
                "headerRow",
                "rowCount",
                "columnCount",
                "cells",
                "exceptions",
                "calculable",
                "fingerprint",
                "describeSeconds",
            )
            if k in r
        }
        | {
            "schema": [(c["name"], c["type"], c["formats"]) for c in r["schema"]][:60],
            "excludedFormulaColumns": r.get("excludedFormulaColumns"),
            "table": r.get("table"),
            "preambleRows": r.get("preambleRows"),
            "previewFirstRow": r["preview"]["rows"][:1],
        }
        for r in regions
    ]
    result["regionCells"] = sum(r["cells"] for r in regions)
    result["regionShareOfLiterals"] = round(result["regionCells"] / max(literal, 1), 4)
    result.update(peak())
    write(out, result)


def windows(store, sheet, out):
    db = sqlite3.connect(store)
    grid, width = db.execute(
        "SELECT grid, width FROM sheets WHERE name=?", (sheet,)
    ).fetchone()
    last = db.execute(f'SELECT max(row) FROM "{grid}"').fetchone()[0]
    db.close()
    result = {"sheet": sheet, "lastRow": last, "windows": []}
    for row in (1, last // 2, max(1, last - 59)):
        result["windows"].append(
            {"row": row, **data_node.window(Path(store), sheet, row, 1, 60, 20)}
        )
    write(out, result)


def link(store, index, out):
    started = time.perf_counter()
    links = data_node.link(Path(store), Path(index))
    write(
        out,
        {
            "seconds": round(time.perf_counter() - started, 2),
            "links": links[:40],
            "linkCount": len(links) - 1,
        },
    )


def verify(store, index, out, samples=2000):
    """Reproducible sample (seed 20261114) of region cells vs serve's index."""
    from linexcel.lazy import _id
    from linexcel.lazy_store import WorkbookIndex

    rng = random.Random(20261114)
    db = sqlite3.connect(store)
    regions = db.execute("SELECT id, sheet, r1, c1, r2, c2 FROM regions").fetchall()
    grids = dict(db.execute("SELECT name, grid FROM sheets").fetchall())
    checked, mismatches = 0, []
    from openpyxl.utils import get_column_letter

    with WorkbookIndex(Path(index)) as nodes:
        for rid, sheet, r1, c1, r2, c2 in regions:
            for _ in range(max(1, int(samples) // max(1, len(regions)))):
                row, col = rng.randint(r1, r2), rng.randint(c1, c2)
                stored = db.execute(
                    f'SELECT c{col} FROM "{grids[sheet]}" WHERE row=?', (row,)
                ).fetchone()
                stored = stored[0] if stored else None
                exception = db.execute(
                    "SELECT kind, value_json FROM exceptions "
                    "WHERE sheet=? AND row=? AND col=? AND kind!='format'",
                    (sheet, row, col),
                ).fetchone()
                node_id = _id(sheet, f"{get_column_letter(col)}{row}")
                node = nodes.get(node_id)
                if node is not None and node.get("kind") != "input":
                    node = None  # formula cells are not region values
                expected = (
                    None
                    if node is None
                    else node.get("cached_comparison_value", node.get("value"))
                )
                if exception:
                    kind = exception[0]
                    stored = (
                        kind,
                        json.loads(exception[1]) if kind in {"bool", "error"} else None,
                    )
                    expected = (
                        None
                        if node is None
                        else (
                            "unreadable"
                            if node.get("diagnostics")
                            else {"b": "bool", "e": "error", "d": "iso_date"}.get(
                                node.get("sourceDataType"), "json"
                            ),
                            node.get("value")
                            if node.get("sourceDataType") in {"b", "e"}
                            and not node.get("diagnostics")
                            else None,
                        )
                    )
                checked += 1
                if stored != expected or (type(stored) is not type(expected)):
                    mismatches.append(
                        {"cell": node_id, "store": stored, "index": expected}
                    )
    db.close()
    write(
        out,
        {
            "checked": checked,
            "mismatches": len(mismatches),
            "sample": mismatches[:10],
            "seed": 20261114,
        },
    )


def evaluate(xlsx, store, index, cell, out):
    import formualizer as fz

    from linexcel import lazy
    from linexcel.lazy_store import WorkbookIndex

    result = {"cell": cell}
    started = time.perf_counter()
    baseline = lazy.evaluate_node(Path(xlsx), cell, index_path=Path(index))
    result["serve"] = {
        "seconds": round(time.perf_counter() - started, 2),
        "status": baseline["status"],
        "value": baseline.get("value"),
        "cached": baseline.get("cachedValue"),
        "comparison": baseline.get("comparison"),
        "inputs": baseline.get("coverage", {}).get("inputCells"),
        "diagnostics": baseline.get("diagnostics", [])[:3],
    }
    regions = {}
    for rid, sheet, r1, c1, r2, c2 in sqlite3.connect(store).execute(
        "SELECT id, sheet, r1, c1, r2, c2 FROM regions"
    ):
        regions.setdefault(sheet, []).append((rid, (r1, c1, r2, c2)))
    created = []
    real_workbook = fz.Workbook

    def recording_workbook(*args, **kwargs):
        engine = real_workbook(*args, **kwargs)
        created.append(engine)
        return engine

    class RegionAwareIndex(WorkbookIndex):
        """Selects only inputs outside data regions; region sub-ranges are
        loaded in row chunks with set_values_batch instead of per cell."""

        pieces: list = []
        loaded = False
        timing: dict = {}

        def select_range_inputs(self, sheet, bounds):
            r1, c1, r2, c2 = bounds
            outside = []
            for rid, (a, b, c, d) in regions.get(sheet, []):
                sub = (max(r1, a), max(c1, b), min(r2, c), min(c2, d))
                if sub[0] <= sub[2] and sub[1] <= sub[3]:
                    self.pieces.append((sheet, sub, rid))
                    outside.append(sub)
            condition = "".join(
                " AND NOT (row BETWEEN ? AND ? AND col BETWEEN ? AND ?)"
                for _ in outside
            )
            args = [x for sub in outside for x in (sub[0], sub[2], sub[1], sub[3])]
            # Same columns as serve's selection, keyed by position since F26.
            self.db.execute(
                "INSERT OR IGNORE INTO selected_inputs SELECT sheet,row,col,id "
                "FROM nodes WHERE sheet=? AND row BETWEEN ? AND ? "
                "AND col BETWEEN ? AND ? AND kind='input'" + condition,
                (sheet, r1, r2, c1, c2, *args),
            )

        def selected_input_count(self):
            region_cells = sum(
                (s[2] - s[0] + 1) * (s[3] - s[1] + 1) for _, s, _ in set(self.pieces)
            )
            return super().selected_input_count() + region_cells

        def selected_inputs(self):
            if created and not self.loaded:
                self.loaded = True
                started = time.perf_counter()
                for sheet, (r1, c1, r2, c2), _rid in sorted(set(self.pieces)):
                    rows, unreadable = data_node.region_values(
                        Path(store), sheet, (r1, c1, r2, c2)
                    )
                    if unreadable:
                        raise ValueError(
                            f"Unreadable region cells in {sheet}: {unreadable[:3]}"
                        )
                    for start in range(0, len(rows), data_node.CHUNK_ROWS):
                        created[-1].set_values_batch(
                            sheet,
                            r1 + start,
                            c1,
                            rows[start : start + data_node.CHUNK_ROWS],
                        )
                self.timing["regionLoadSeconds"] = round(
                    time.perf_counter() - started, 2
                )
            yield from super().selected_inputs()

    fz.Workbook = recording_workbook
    try:
        started = time.perf_counter()
        with RegionAwareIndex(Path(index)) as nodes:
            proto = lazy._evaluate_index(nodes, cell, 20000)
            pieces = sorted({(s, sub) for s, sub, _ in nodes.pieces})
        result["regionLoading"] = {
            "seconds": round(time.perf_counter() - started, 2),
            "status": proto["status"],
            "value": proto.get("value"),
            "cached": proto.get("cachedValue"),
            "comparison": proto.get("comparison"),
            "inputs": proto.get("coverage", {}).get("inputCells"),
            "subRanges": [[s, list(sub)] for s, sub in pieces][:10],
            **RegionAwareIndex.timing,
            "diagnostics": proto.get("diagnostics", [])[:3],
        }
    finally:
        fz.Workbook = real_workbook
    result["sameValue"] = result["serve"]["value"] == result["regionLoading"]["value"]
    write(out, result)


if __name__ == "__main__":
    command, *args = sys.argv[1:]
    {
        "build": build,
        "windows": windows,
        "link": link,
        "verify": verify,
        "evaluate": evaluate,
    }[command](*args)
