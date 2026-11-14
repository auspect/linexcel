"""Is set_values_batch superlinear in rows per call, or in sheet height?

Load the same 250k x 2 real ledger columns as: one call; chunks of 10k rows;
chunks of 1k rows; and one column at a time. Also a native load_workbook of a
copy is not possible here (no writer), so compare with load_workbook of the
fpna file itself (all sheets) as the engine's own ingest path.
usage: probe_fz_chunks.py <fpna.xlsx> <out.json>
"""

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import formualizer as fz  # noqa: E402
import python_calamine as pc  # noqa: E402
from winmem import peak_mb  # noqa: E402

rows = (
    pc.CalamineWorkbook.from_path(sys.argv[1])
    .get_sheet_by_name("Ledger_01")
    .to_python()
)
header = rows[0]
ent, amt = header.index("Entity"), header.index("Amount")
data = [[r[ent], r[amt]] for r in rows[1:]]
n = len(data)
out = {"rows": n}


def engine():
    cfg = fz.EvaluationConfig()
    cfg.enable_parallel = False
    wb = fz.Workbook(config=fz.WorkbookConfig(eval_config=cfg))
    wb.add_sheet("D")
    wb.add_sheet("C")
    return wb


def check(wb):
    wb.set_formula("C", 1, 1, f'=SUMIFS(D!B1:B{n},D!A1:A{n},"{data[0][0]}")')
    return wb.evaluate_cell("C", 1, 1)


for label, chunk in (("oneCall", n), ("chunks10k", 10_000), ("chunks1k", 1_000)):
    wb = engine()
    t = time.perf_counter()
    for start in range(0, n, chunk):
        wb.set_values_batch("D", start + 1, 1, data[start : start + chunk])
    seconds = round(time.perf_counter() - t, 3)
    out[label] = {"seconds": seconds, "sumifs": check(wb)}
    print(label, out[label], flush=True)

# rows inserted bottom-up: does order matter?
wb = engine()
t = time.perf_counter()
for start in reversed(range(0, n, 10_000)):
    wb.set_values_batch("D", start + 1, 1, data[start : start + 10_000])
out["chunks10kReverse"] = {
    "seconds": round(time.perf_counter() - t, 3),
    "sumifs": check(wb),
}
print("reverse", out["chunks10kReverse"], flush=True)

# engine's own reader on the real file (all 9 sheets, 2.7M cells + 175k formulas)
t = time.perf_counter()
wb = fz.load_workbook(sys.argv[1])
out["nativeLoadWorkbookSeconds"] = round(time.perf_counter() - t, 2)
t = time.perf_counter()
wb.set_formula(
    "Summary",
    20,
    20,
    f'=SUMIFS(Ledger_01!J2:J{n + 1},Ledger_01!B2:B{n + 1},"{data[0][0]}")',
)
out["nativeSumifs"] = wb.evaluate_cell("Summary", 20, 20)
out["nativeSumifsSeconds"] = round(time.perf_counter() - t, 3)
out.update(peak_mb())
print(out, flush=True)
Path(sys.argv[2]).write_text(json.dumps(out, indent=1), "utf-8")
