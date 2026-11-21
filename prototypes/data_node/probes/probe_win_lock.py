"""Cost of per-statement SQLite locking on serve's read paths (Windows).

Times evaluate_node and graph_neighborhood on a built index twice: as shipped
(read-only connection in autocommit: every SELECT takes and releases the file
lock), then with one read transaction held for the connection's lifetime.
Values must be identical; only time may differ.
usage: probe_win_lock.py <xlsx> <index> <out.json> <cell> [<cell> ...]
"""

import json
import sys
import time
from pathlib import Path

from linexcel import lazy
from linexcel.lazy_store import WorkbookIndex, graph_neighborhood

xlsx, index, out, *cells = sys.argv[1:]
original = WorkbookIndex.__init__


def with_read_transaction(self, path, *, writable=False):
    original(self, path, writable=writable)
    if not writable:
        self.db.execute("BEGIN")


result = {"workbook": Path(xlsx).name, "cells": {}}
for cell in cells:
    record = {}
    for mode in ("autocommit", "readTransaction"):
        WorkbookIndex.__init__ = (
            original if mode == "autocommit" else with_read_transaction
        )
        started = time.perf_counter()
        evaluation = lazy.evaluate_node(Path(xlsx), cell, index_path=Path(index))
        evaluate_seconds = time.perf_counter() - started
        started = time.perf_counter()
        around = graph_neighborhood(Path(index), cell, depth=2)
        record[mode] = {
            "evaluateSeconds": round(evaluate_seconds, 2),
            "status": evaluation["status"],
            "value": evaluation.get("value"),
            "closureNodes": evaluation.get("dependency_count"),
            "neighborhoodSeconds": round(time.perf_counter() - started, 3),
            "neighborhoodNodes": len(around["nodes"]),
        }
    WorkbookIndex.__init__ = original
    record["sameValue"] = (
        record["autocommit"]["value"] == record["readTransaction"]["value"]
    )
    result["cells"][cell] = record
    print(cell, json.dumps(record), flush=True)
Path(out).write_text(json.dumps(result, indent=1, default=str), "utf-8")
