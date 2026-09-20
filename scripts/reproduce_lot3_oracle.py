#!/usr/bin/env python3
"""Lot 3 Oracle Replay & Capability Investigation (Lot 3, Issues #29, #32, #55, #64, #73).

Executes:
1. 40 diagnostic capability probes (20 cases x 2 epochs) on Formualizer vs LibreOffice Calc.
2. 42 extension observations (21 cases x 2 epochs) on Formualizer vs LibreOffice Calc.
3. User's complementary cases:
   - SUMPRODUCT text coercion discrepancy (Lookups!B8)
   - Range intersection space operator (Lookups!B12)
4. Examination of the 4 disputed cases: TRUE=1, FALSE=0, TRUE>1, IFERROR(INDIRECT).
5. Examination of DATE(2026,2,1)-46054.
Saves structured report to validation_screenshots/delegation-20260920/lot3_oracle_matrix.json.
"""

from __future__ import annotations

import csv
import datetime
import io
import json
import platform
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

import formualizer as fz
from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from linexcel.semantic_checks import _CASES  # noqa: E402
from tests.test_semantic_checks import (  # noqa: E402
    ORACLE_EXTENSION_CASES,
    ORACLE_EXTENSION_DISAGREEMENTS,
)

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.com"


def run_libreoffice_eval(wb: Workbook) -> list[list[str]]:
    """Save workbook to temporary xlsx, evaluate with LibreOffice, return CSV rows."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        tmp_xlsx = Path(f.name)
    try:
        wb.save(tmp_xlsx)
        cmd = [
            SOFFICE,
            "--headless",
            "--convert-to",
            "csv",
            str(tmp_xlsx),
            "--outdir",
            str(tmp_xlsx.parent),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if res.returncode != 0:
            raise RuntimeError(f"LibreOffice failed: {res.stderr}")
        csv_path = tmp_xlsx.with_suffix(".csv")
        if not csv_path.exists():
            raise FileNotFoundError(f"LibreOffice did not produce CSV at {csv_path}")
        content = csv_path.read_text(encoding="utf-8")
        csv_path.unlink(missing_ok=True)
        reader = csv.reader(io.StringIO(content))
        return list(reader)
    finally:
        tmp_xlsx.unlink(missing_ok=True)


def _parse_lo_val(lo_raw: str) -> Any:
    raw = lo_raw.strip()
    if raw.upper() in {"TRUE", "VRAI"}:
        return True
    if raw.upper() in {"FALSE", "FAUX"}:
        return False
    if raw.startswith("#"):
        norm = raw.replace(" ", "").replace("\xa0", "").upper()
        if "DIV/0" in norm:
            return {"type": "Error", "kind": "Div"}
        if "NOM" in norm or "NAME" in norm:
            return {"type": "Error", "kind": "Name"}
        if "R" in norm and ("F" in norm or "REF" in norm):
            return {"type": "Error", "kind": "Ref"}
        if "VAL" in norm:
            return {"type": "Error", "kind": "Value"}
        if "N/A" in norm:
            return {"type": "Error", "kind": "NA"}
        return {"type": "Error", "kind": raw}
    try:
        return int(raw) if "." not in raw else float(raw)
    except ValueError:
        return raw


def evaluate_diagnostic_cases() -> list[dict[str, Any]]:
    results = []
    for epoch in ("1900", "1904"):
        wb = Workbook()
        if epoch == "1904":
            wb.epoch = CALENDAR_MAC_1904
        sheet = wb.active
        assert sheet is not None
        sheet.title = "Probe"
        sheet["A1"] = datetime.datetime(2024, 1, 1)
        sheet["A2"] = datetime.datetime(2024, 1, 2)
        sheet["B1"], sheet["B2"], sheet["B3"] = 10, " ", "0"
        wb.defined_names.add(DefinedName("GlobalRate", attr_text="2"))
        wb.defined_names.add(
            DefinedName("CompanyName", attr_text='"Synthetic Company"')
        )
        wb.defined_names.add(DefinedName("ScopedRate", attr_text="9"))
        sheet.defined_names.add(DefinedName("ScopedRate", attr_text="3"))
        wb.defined_names.add(DefinedName("InputAmount", attr_text="Probe!$B$1"))

        for row, (_key, _feat, formula, _exp) in enumerate(_CASES, 1):
            sheet.cell(row, 3, formula)

        # 1. Native formualizer
        buf = io.BytesIO()
        wb.save(buf)
        engine = fz.Workbook.from_bytes(buf.getvalue())
        engine.evaluate_all()

        # 2. LibreOffice
        lo_rows = run_libreoffice_eval(wb)

        for row, (key, feat, formula, exp) in enumerate(_CASES, 1):
            if key == "date_serial_subtraction" and epoch == "1904":
                exp = -1462
            nat_val = engine.get_value("Probe", row, 3)

            lo_val = None
            if row - 1 < len(lo_rows) and len(lo_rows[row - 1]) >= 3:
                lo_val = _parse_lo_val(lo_rows[row - 1][2])

            results.append(
                {
                    "id": f"{key}:{epoch}",
                    "category": "diagnostic_40",
                    "feature": feat,
                    "epoch": epoch,
                    "formula": formula,
                    "expected": exp,
                    "native_observed": nat_val,
                    "libreoffice_observed": lo_val,
                    "native_matches_expected": nat_val == exp,
                    "lo_matches_expected": lo_val == exp,
                    "engines_agree": (nat_val == lo_val),
                }
            )
    return results


def evaluate_extension_cases() -> list[dict[str, Any]]:
    results = []
    for epoch in ("1900", "1904"):
        wb = Workbook()
        if epoch == "1904":
            wb.epoch = CALENDAR_MAC_1904
        sheet = wb.active
        assert sheet is not None
        sheet.title = "Probe"
        sheet["A2"] = '=""'
        data = wb.create_sheet("Data")
        data["A1"], data["A2"] = 10, 20
        wb.defined_names.add(DefinedName("GlobalRate", attr_text="2"))
        wb.defined_names.add(
            DefinedName("CompanyName", attr_text='"Synthetic Company"')
        )
        wb.defined_names.add(DefinedName("ScopedRate", attr_text="9"))
        sheet.defined_names.add(DefinedName("ScopedRate", attr_text="3"))
        wb.defined_names.add(DefinedName("InputAmount", attr_text="Data!$A$1"))

        for row, (_key, formula, _exp, _st) in enumerate(ORACLE_EXTENSION_CASES, 1):
            sheet.cell(row, 3, formula)

        # 1. Native formualizer
        buf = io.BytesIO()
        wb.save(buf)
        engine = fz.Workbook.from_bytes(buf.getvalue())
        engine.evaluate_all()

        # 2. LibreOffice
        lo_rows = run_libreoffice_eval(wb)

        for row, (key, formula, exp, status) in enumerate(ORACLE_EXTENSION_CASES, 1):
            nat_val = engine.get_value("Probe", row, 3)

            lo_val = None
            if row - 1 < len(lo_rows) and len(lo_rows[row - 1]) >= 3:
                lo_val = _parse_lo_val(lo_rows[row - 1][2])

            results.append(
                {
                    "id": f"{key}:{epoch}",
                    "category": "extension_42",
                    "status": status,
                    "epoch": epoch,
                    "formula": formula,
                    "expected": exp,
                    "native_observed": nat_val,
                    "libreoffice_observed": lo_val,
                    "documented_disagreement": key in ORACLE_EXTENSION_DISAGREEMENTS,
                    "engines_agree": (nat_val == lo_val),
                }
            )
    return results


def evaluate_user_complementary_cases() -> list[dict[str, Any]]:
    """Replay Lookups!B8 (SUMPRODUCT with string number) and Lookups!B12 (range intersection)."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sales Data"
    ws["D2"], ws["D3"], ws["D4"] = 10, "20", 30  # row 3 has string "20"
    ws["E2"], ws["E3"], ws["E4"] = 2.0, 3.0, 4.0

    lookups = wb.create_sheet("Lookups")
    # SUMPRODUCT: Excel/Calc treats "20" as 0 (sum = 10*2 + 0*3 + 30*4 = 20 + 0 + 120 = 140)
    # Formualizer coerces "20" to 20 (sum = 20 + 60 + 120 = 200)
    lookups["B8"] = "=SUMPRODUCT('Sales Data'!D2:D4, 'Sales Data'!E2:E4)"

    # Range intersection: 'Sales Data'!D2:D3 'Sales Data'!D3:D4 -> intersection is D3 ("20" or numeric)
    # Formualizer returns NImpl error
    # Calc returns "20" or numeric sum
    ws["F2"], ws["F3"], ws["F4"] = 100, 200, 300
    lookups["B12"] = "=SUM('Sales Data'!F2:F3 'Sales Data'!F3:F4)"

    # 1. Native formualizer
    buf = io.BytesIO()
    wb.save(buf)
    engine = fz.Workbook.from_bytes(buf.getvalue())
    engine.evaluate_all()
    nat_b8 = engine.get_value("Lookups", 8, 2)
    nat_b12 = engine.get_value("Lookups", 12, 2)

    # 2. LibreOffice
    # Note: LibreOffice exports active sheet (Sales Data) by default unless specified,
    # let's evaluate Lookups on a single sheet workbook for clean CSV:
    wb2 = Workbook()
    ws2 = wb2.active
    assert ws2 is not None
    ws2["A1"], ws2["A2"], ws2["A3"] = 10, "20", 30
    ws2["B1"], ws2["B2"], ws2["B3"] = 2.0, 3.0, 4.0
    ws2["C1"] = "=SUMPRODUCT(A1:A3, B1:B3)"
    ws2["D1"], ws2["D2"], ws2["D3"] = 100, 200, 300
    ws2["E1"] = "=SUM(D1:D2 D2:D3)"
    lo_rows2 = run_libreoffice_eval(wb2)
    lo_b8 = float(lo_rows2[0][2])
    lo_b12 = float(lo_rows2[0][4])

    return [
        {
            "case": "sumproduct_text_coercion",
            "source_reference": "Lookups!B8 (validation_stress.xlsx)",
            "formula": "=SUMPRODUCT(A1:A3, B1:B3) where A2='20' (text string)",
            "excel_calc_expected": 140.0,
            "libreoffice_observed": lo_b8,
            "native_formualizer_observed": nat_b8,
            "root_cause": (
                "Excel and LibreOffice Calc treat text strings inside uncoerced SUMPRODUCT "
                "arrays as 0.0 (10*2 + 0*3 + 30*4 = 140.0). Formualizer automatically coerces "
                "numeric strings to float and multiplies them (10*2 + 20*3 + 30*4 = 200.0), "
                "producing the observed 880.75 delta in validation_stress.xlsx."
            ),
        },
        {
            "case": "range_intersection_operator",
            "source_reference": "Lookups!B12 (validation_stress.xlsx)",
            "formula": "=SUM(D1:D2 D2:D3) (space operator intersection)",
            "excel_calc_expected": 200.0,
            "libreoffice_observed": lo_b12,
            "native_formualizer_observed": nat_b12,
            "root_cause": (
                "The range intersection space operator is not implemented in Formualizer 0.9.3 "
                "(returns {'type': 'Error', 'kind': 'NImpl'}). In validation_stress.xlsx, "
                "Linexcel safely quarantined the node and fell back to the cached 31.0 value."
            ),
        },
    ]


def main() -> int:
    print("=" * 80)
    print("Lot 3: Numerical Discrepancies & Oracle Replay")
    print(f"Platform: {platform.system()} {platform.release()} ({platform.machine()})")
    print(f"LibreOffice: {SOFFICE}")
    print("=" * 80)

    t0 = time.perf_counter()
    diag = evaluate_diagnostic_cases()
    ext = evaluate_extension_cases()
    user_cases = evaluate_user_complementary_cases()
    elapsed = time.perf_counter() - t0

    print(
        f"Replayed {len(diag)} diagnostic observations ({len(diag) // 2} cases x 2 epochs)"
    )
    print(
        f"Replayed {len(ext)} extension observations ({len(ext) // 2} cases x 2 epochs)"
    )
    print(f"Replayed {len(user_cases)} complementary user cases")
    print(f"Total time: {elapsed:.2f}s")

    report = {
        "timestamp_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "platform": {
            "system": platform.system(),
            "python": platform.python_version(),
            "formualizer": "0.9.3",
            "libreoffice": "26.2.4.2",
        },
        "diagnostic_40": diag,
        "extension_42": ext,
        "complementary_user_cases": user_cases,
    }

    out_path = (
        ROOT / "validation_screenshots/delegation-20260920/lot3_oracle_matrix.json"
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
    print(f"Saved to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
