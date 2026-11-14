"""F06 on the real corpus: serve's build_index, timed in a fresh process.

usage: measure_import.py <xlsx> <index.sqlite> <result.json>
"""

import json
import sqlite3
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from winmem import peak_mb  # noqa: E402

started = time.perf_counter()
from linexcel.lazy_store import build_index  # noqa: E402

workbook, index, out = map(Path, sys.argv[1:4])
index.unlink(missing_ok=True)
result = {"workbook": workbook.name, "bytes": workbook.stat().st_size}
try:
    build_index(workbook, index)
    result["status"] = "ok"
except BaseException as exc:  # record, never hide
    result["status"] = "failed"
    result["error"] = f"{type(exc).__name__}: {exc}"
result["seconds"] = round(time.perf_counter() - started, 2)
result.update(peak_mb())
if index.exists():
    result["indexMB"] = round(index.stat().st_size / 1e6, 1)
    db = sqlite3.connect(index)
    meta = json.loads(
        db.execute("select value from metadata where key='meta'").fetchone()[0]
    )
    result.update(
        nodeCounts=meta.get("nodeCounts"),
        edgeCount=meta.get("edgeCount"),
        indexStatus=meta.get("status"),
        diagnostics=len(meta.get("diagnostics", [])),
        diagnosticSample=meta.get("diagnostics", [])[:5],
        sheetStatus={s["name"]: s["status"] for s in meta.get("sheetDetails", [])},
        hiddenRowsStored=sum(
            len(s.get("hiddenRows", [])) for s in meta.get("sheetDetails", [])
        ),
        metaJsonBytes=len(json.dumps(meta)),
    )
    try:
        result["tableMB"] = {
            name: round(size / 1e6, 1)
            for name, size in db.execute(
                "select name, sum(pgsize) from dbstat "
                "group by name order by 2 desc limit 12"
            )
        }
    except sqlite3.Error as exc:
        result["tableMB"] = f"dbstat unavailable: {exc}"
    cells = sum(
        (meta.get("nodeCounts") or {}).get(k, 0)
        for k in ("input", "cell", "array_result")
    )
    if cells:
        result["storedCells"] = cells
        result["usPerCell"] = round(result["seconds"] * 1e6 / cells, 1)
        result["indexBytesPerCell"] = round(index.stat().st_size / cells, 1)
        avg = db.execute(
            "select avg(length(data)) from "
            "(select data from nodes where kind='input' limit 200000)"
        ).fetchone()[0]
        result["avgInputBlobBytes"] = round(avg or 0, 1)
    db.close()
out.write_text(json.dumps(result, indent=1, ensure_ascii=False), "utf-8")
print(json.dumps(result, ensure_ascii=False))
