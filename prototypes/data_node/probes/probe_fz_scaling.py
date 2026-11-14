"""Scaling of engine value loading in formualizer 0.14.0.

Per-cell set_value vs one set_values_batch, numbers vs text, 10k..250k rows
x 2 columns, then SUM / SUMIFS time. Real fpna ledger columns (Entity, Amount).
usage: probe_fz_scaling.py <fpna.xlsx> <out.json>
"""

import json
import sys
import time
from pathlib import Path

import formualizer as fz
import python_calamine as pc

rows = (
    pc.CalamineWorkbook.from_path(sys.argv[1])
    .get_sheet_by_name("Ledger_01")
    .to_python()
)
header = rows[0]
ent, amt = header.index("Entity"), header.index("Amount")
results = []


def engine():
    cfg = fz.EvaluationConfig()
    cfg.enable_parallel = False
    wb = fz.Workbook(config=fz.WorkbookConfig(eval_config=cfg))
    wb.add_sheet("D")
    wb.add_sheet("C")
    return wb


for n in (10_000, 50_000, 100_000, 250_000):
    body = rows[1 : n + 1]
    for kind, data in (
        ("numbers", [[r[amt], r[amt] * 2] for r in body]),
        ("text", [[r[ent], r[ent] + "x"] for r in body]),
        ("mixed", [[r[ent], r[amt]] for r in body]),
    ):
        record = {"rows": len(data), "kind": kind}
        wb = engine()
        t = time.perf_counter()
        for i, row in enumerate(data, 1):
            wb.set_value("D", i, 1, row[0])
            wb.set_value("D", i, 2, row[1])
        record["perCellSeconds"] = round(time.perf_counter() - t, 3)
        wb = engine()
        t = time.perf_counter()
        wb.set_values_batch("D", 1, 1, data)
        record["batchSeconds"] = round(time.perf_counter() - t, 3)
        if kind == "mixed":
            wb.set_formula(
                "C",
                1,
                1,
                f'=SUMIFS(D!B1:B{len(data)},D!A1:A{len(data)},"{data[0][0]}")',
            )
            t = time.perf_counter()
            record["sumifs"] = wb.evaluate_cell("C", 1, 1)
            record["sumifsSeconds"] = round(time.perf_counter() - t, 3)
        record["batchUsPerCell"] = round(
            record["batchSeconds"] * 1e6 / (2 * len(data)), 2
        )
        record["perCellUsPerCell"] = round(
            record["perCellSeconds"] * 1e6 / (2 * len(data)), 2
        )
        print(record, flush=True)
        results.append(record)
Path(sys.argv[2]).write_text(json.dumps(results, indent=1), "utf-8")
