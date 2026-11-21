"""formualizer 0.14.0: can a values-only region be loaded as one unit?

Beliefs probed (lot 2):
  B1 no region/provider API: a data region must still become engine cells;
  B2 set_values_batch is much faster than per-cell set_value on real data;
  B3 set_values_batch keeps types (text, number, bool, empty, error) and
     SUMIFS/structured refs over it give the same result as per-cell loading;
  B4 add_table works over batch-loaded values;
  B5 range_page pages a range without evaluating it.

Data: the real ledger table of fpna-20mb (tblLedger_01, 249,675 rows x 10).
usage: probe_fz_batch.py <fpna.xlsx> <out.json>
"""

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import formualizer as fz  # noqa: E402
import python_calamine as pc  # noqa: E402
from winmem import peak_mb  # noqa: E402

path, out = Path(sys.argv[1]), Path(sys.argv[2])
result = {"formualizer": fz.__version__ if hasattr(fz, "__version__") else None}
import importlib.metadata as md  # noqa: E402

result["formualizer"] = md.version("formualizer")
result["publicNamesWithRegionOrLazy"] = sorted(
    n
    for n in dir(fz)
    if any(
        k in n.lower()
        for k in ("region", "lazy", "provider", "columnar", "arrow", "frame")
    )
)
wbmembers = dir(fz.Workbook)
result["workbookMembersWithBatchRegionTable"] = sorted(
    n
    for n in wbmembers
    if any(
        k in n.lower()
        for k in ("batch", "region", "table", "page", "bulk", "load", "values")
    )
)

t = time.perf_counter()
rows = (
    pc.CalamineWorkbook.from_path(str(path)).get_sheet_by_name("Ledger_01").to_python()
)
result["calamineReadSeconds"] = round(time.perf_counter() - t, 2)
header, body = rows[0], rows[1:]
result["rows"], result["columns"] = len(body), len(header)
result["header"] = header
result["pythonTypesRow1"] = [type(v).__name__ for v in body[0]]
# Engine values: dates as serials would need linexcel's conversion; keep the
# probe on what the engine receives today (numbers, text, bools).
import datetime as dt  # noqa: E402

EPOCH = dt.datetime(1899, 12, 30)


def engine_value(v):
    if isinstance(v, dt.datetime):
        return (v - EPOCH).total_seconds() / 86400
    if isinstance(v, dt.date):
        return (dt.datetime(v.year, v.month, v.day) - EPOCH).days
    if v == "":
        return None
    return v


body = [[engine_value(v) for v in r] for r in body]
data = [list(header)] + body
amount_col = header.index("Amount") + 1 if "Amount" in header else len(header)
entity_col = header.index("Entity") + 1 if "Entity" in header else 2
from openpyxl.utils import get_column_letter as letter  # noqa: E402

A, E = letter(amount_col), letter(entity_col)
first_entity = body[0][entity_col - 1]
last = len(data)
formula_sum = f"=SUM(Ledger_01!{A}2:{A}{last})"
formula_sumifs = (
    f'=SUMIFS(Ledger_01!{A}2:{A}{last},Ledger_01!{E}2:{E}{last},"{first_entity}")'
)
formula_struct = f'=SUMIFS(tblLedger_01[Amount],tblLedger_01[Entity],"{first_entity}")'
expected_sum = sum(
    r[amount_col - 1] for r in body if isinstance(r[amount_col - 1], (int, float))
)
expected_sumifs = sum(
    r[amount_col - 1]
    for r in body
    if r[entity_col - 1] == first_entity and isinstance(r[amount_col - 1], (int, float))
)
result["expected"] = {"sum": expected_sum, "sumifs": expected_sumifs}


def engine():
    cfg = fz.EvaluationConfig()
    cfg.enable_parallel = False
    wb = fz.Workbook(config=fz.WorkbookConfig(eval_config=cfg))
    wb.add_sheet("Ledger_01")
    wb.add_sheet("Calc")
    return wb


def evaluate(wb):
    wb.set_formula("Calc", 1, 1, formula_sum)
    wb.set_formula("Calc", 2, 1, formula_sumifs)
    t = time.perf_counter()
    values = [wb.evaluate_cell("Calc", 1, 1), wb.evaluate_cell("Calc", 2, 1)]
    return values, round(time.perf_counter() - t, 3)


# per-cell, as lazy.py does today
wb = engine()
t = time.perf_counter()
for r, row in enumerate(data, 1):
    for c, v in enumerate(row, 1):
        if v is not None:
            wb.set_value("Ledger_01", r, c, v)
load = time.perf_counter() - t
values, seconds = evaluate(wb)
result["perCell"] = {
    "loadSeconds": round(load, 2),
    "evalSeconds": seconds,
    "values": values,
}

# one batch
wb = engine()
t = time.perf_counter()
wb.set_values_batch("Ledger_01", 1, 1, data)
load = time.perf_counter() - t
values, seconds = evaluate(wb)
result["batch"] = {
    "loadSeconds": round(load, 2),
    "evalSeconds": seconds,
    "values": values,
}

# B4 table over batch-loaded values + structured reference
try:
    wb.add_table(
        "tblLedger_01",
        "Ledger_01",
        (1, 1, last, len(header)),
        header,
        header_row=True,
        totals_row=False,
    )
    wb.set_formula("Calc", 3, 1, formula_struct)
    result["tableStructuredRef"] = wb.evaluate_cell("Calc", 3, 1)
except Exception as exc:  # report, do not hide
    result["tableStructuredRef"] = f"{type(exc).__name__}: {exc}"

# B5 range_page without evaluation
try:
    page = wb.range_page(f"Ledger_01!A1:{letter(len(header))}{last}", offset=0, limit=3)
    d = page.to_dict()
    result["rangePage"] = {
        "total": d.get("total"),
        "next": d.get("next_offset"),
        "first": str(d.get("items", [])[:2])[:400],
    }
except Exception as exc:
    result["rangePage"] = f"{type(exc).__name__}: {exc}"

# B3 types: text / bool / empty / error / number round trip
wb2 = engine()
wb2.set_values_batch(
    "Ledger_01", 1, 1, [["t", True, None, 1.5, {"type": "Error", "kind": "Div"}]]
)
try:
    got = wb2.get_values(fz.RangeAddress("Ledger_01", 1, 1, 1, 5))
    result["typesRoundTrip"] = repr(got)
except Exception as exc:
    result["typesRoundTrip"] = f"{type(exc).__name__}: {exc}"
for c, f in enumerate(
    [
        "=ISTEXT(Ledger_01!A1)",
        "=ISLOGICAL(Ledger_01!B1)",
        "=ISBLANK(Ledger_01!C1)",
        "=ISERROR(Ledger_01!E1)",
        "=ERROR.TYPE(Ledger_01!E1)",
    ],
    1,
):
    wb2.set_formula("Calc", 1, c, f)
result["typeChecks"] = [repr(wb2.evaluate_cell("Calc", 1, c)) for c in range(1, 6)]
result.update(peak_mb())
out.write_text(json.dumps(result, indent=1, default=str), "utf-8")
print(json.dumps(result, indent=1, default=str))
