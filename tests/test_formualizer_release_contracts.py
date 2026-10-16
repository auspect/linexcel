"""Versioned native corrections must survive linexcel's import adaptations.

Older supported engines are skipped rather than making their known incorrect
results an acceptance oracle. These cases do not certify row visibility or
locale-dependent date parsing.
"""

import io
from importlib.metadata import version

import formualizer as fz
import pytest
from openpyxl import Workbook
from packaging.version import Version

from linexcel.engine import _open_workbook

FORMUALIZER_VERSION = Version(version("formualizer"))


@pytest.fixture(params=["native", "linexcel"])
def mode(request):
    return request.param


def _evaluate_workbook(cells, mode):
    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    for address, value in cells.items():
        sheet[address] = value
    stream = io.BytesIO()
    book.save(stream)
    data = stream.getvalue()
    engine = (
        fz.Workbook.from_bytes(data)
        if mode == "native"
        else _open_workbook(data, parallel=False)
    )
    engine.evaluate_all()
    return engine


@pytest.mark.skipif(
    FORMUALIZER_VERSION < Version("0.11.0"),
    reason="Left-associative powers and AND/OR error propagation require 0.11.0.",
)
@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=2^3^2", 64),
        ("=AND(FALSE,1/0)", {"type": "Error", "kind": "Div"}),
        ("=OR(TRUE,1/0)", {"type": "Error", "kind": "Div"}),
    ],
)
def test_power_associativity_and_logical_errors(mode, formula, expected):
    engine = _evaluate_workbook({"D1": formula}, mode)
    assert engine.get_value("S", 1, 4) == expected


@pytest.mark.skipif(
    FORMUALIZER_VERSION < Version("0.11.1"),
    reason="Literal percent, underscore and backslash criteria require 0.11.1.",
)
@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ('=COUNTIF(A1:A7,"a%b")', 1),
        ('=COUNTIF(A1:A7,"a_b")', 1),
        ('=COUNTIF(A1:A7,"a\\b")', 1),
        ('=SUMIF(A1:A7,"a%b",B1:B7)', 1),
        ('=SUMIF(A1:A7,"a_b",B1:B7)', 3),
        ('=SUMIF(A1:A7,"a\\b",B1:B7)', 4),
        ('=AVERAGEIF(A1:A7,"a%b",B1:B7)', 1),
        ('=AVERAGEIF(A1:A7,"a_b",B1:B7)', 3),
        ('=AVERAGEIF(A1:A7,"a\\b",B1:B7)', 4),
        # Excel's own wildcard and its escape retain their meaning.
        ('=COUNTIF(A1:A7,"a*b")', 7),
        ('=COUNTIF(A1:A7,"a~*b")', 1),
    ],
)
def test_criteria_keep_sql_like_characters_literal(mode, formula, expected):
    labels = ["a%b", "axb", "a_b", "a\\b", "a*b", "a?b", "a~b"]
    cells = {f"A{row}": label for row, label in enumerate(labels, 1)}
    cells.update({f"B{row}": row for row in range(1, 8)})
    cells["D1"] = formula

    engine = _evaluate_workbook(cells, mode)

    # The three literal criteria each match exactly one input, with weights
    # 1, 3 and 4. The unescaped Excel wildcard matches all seven inputs.
    assert engine.get_value("S", 1, 4) == expected


@pytest.mark.skipif(
    FORMUALIZER_VERSION < Version("0.11.1"),
    reason="AGGREGATE options 4-7 require 0.11.1.",
)
@pytest.mark.parametrize(
    ("option", "expected"),
    [
        (4, {"type": "Error", "kind": "Div"}),
        (5, {"type": "Error", "kind": "Div"}),
        (6, 4),
        (7, 4),
    ],
)
def test_aggregate_options_propagate_or_ignore_a_real_error(mode, option, expected):
    engine = _evaluate_workbook(
        {"B1": 1, "B2": "=1/0", "B3": 3, "D1": f"=AGGREGATE(9,{option},B1:B3)"},
        mode,
    )

    assert engine.get_value("S", 2, 2) == {"type": "Error", "kind": "Div"}
    assert engine.get_value("S", 1, 4) == expected
