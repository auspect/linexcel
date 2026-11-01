"""Dynamic reference closure derives from source expressions, never cached values."""

import pytest
from openpyxl import Workbook

from linexcel.lazy import build_structure, evaluate_node


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ('=INDIRECT("A2")', 7),
        ('=SUM(INDIRECT("A1:A3"))', 21),
        ('=INDIRECT("A"&B1)', 7),
        ('=INDIRECT("A"&(B2-1))', 7),
        ('=INDIRECT("R2C1",FALSE)', 7),
        ('=INDIRECT("RC[-2]",FALSE)', 5),
        ("=SUM(OFFSET(A1,1,0,2,1))", 16),
        ("=SUM(OFFSET(A1:A2,1,0))", 16),
        ("=SUM(OFFSET(A1,0,0,B2,1))", 21),
        ('=SUM(OFFSET(INDIRECT("A1"),1,0,2,1))', 16),
    ],
)
def test_source_address_expressions_expand_exact_closure(tmp_path, formula, expected):
    book = Workbook()
    book.active["A1"], book.active["A2"], book.active["A3"] = 5, 7, 9
    book.active["B1"], book.active["B2"] = 2, "=B1+1"
    book.active["C1"] = formula
    book.active["D10"] = "=UNSUPPORTED_UNRELATED()"
    path = tmp_path / "dynamic.xlsx"
    book.save(path)
    graph = build_structure(path)
    target = next(n for n in graph["nodes"] if n["id"] == "Sheet!C1")
    assert target["sourceFormula"] == formula
    assert target["dynamicReferences"]
    result = evaluate_node(path, "Sheet!C1")
    assert result["status"] == "completed", result["diagnostics"]
    assert result["value"] == expected
    assert result["volatile"]
    assert "Sheet!D10" not in [s["nodeId"] for s in result["steps"]]


@pytest.mark.parametrize(
    "formula",
    [
        '=INDIRECT("A"&RAND())',
        "=OFFSET(A1,-1,0)",
        "=OFFSET(A1,0,0,0,1)",
        "=OFFSET(A1,0.5,0)",
        '=INDIRECT("R[-9]C1",FALSE)',
        '=INDIRECT("[missing.xlsx]Sheet!A1")',
    ],
)
def test_unestablished_dynamic_addresses_are_diagnostic(tmp_path, formula):
    book = Workbook()
    book.active["A1"] = 5
    book.active["C1"] = formula
    path = tmp_path / "dynamic-guard.xlsx"
    book.save(path)
    result = evaluate_node(path, "Sheet!C1")
    assert result["status"] == "unsupported"
    assert result["value"] is None
    assert any(
        d["code"] == "dynamic_reference_unresolved"
        for d in result["capabilityDiagnostics"]
    )
