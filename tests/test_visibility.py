"""Reference semantics tested through real native dependency evaluation."""

import math

import formualizer as fz
import pytest

from linexcel.refs import parse_ref
from linexcel.visibility import (
    VisibilityError,
    register_visibility_functions,
    rewrite_visibility_formula,
)


def calculate(formula, values, *, hidden=(), source_formulas=None, filters=False):
    engine = fz.Workbook()
    engine.add_sheet("Sheet")
    source_formulas = source_formulas or {}
    for row, value in enumerate(values, 1):
        if value is not None:
            engine.set_value("Sheet", row, 1, value)

    def resolve(reference, sheet):
        rect = parse_ref(reference, sheet)
        assert rect is not None
        return {
            "expression": rect.to_a1(),
            "bounds": [rect.r1, rect.c1, rect.r2, rect.c2],
            "cells": [
                {
                    "row": row,
                    "column": 1,
                    "formula": source_formulas.get(row),
                    "context": {"rowHidden": row in hidden},
                }
                for row in range(rect.r1, rect.r2 + 1)
                if rect.c1 <= 1 <= rect.c2
            ],
            "activeFilters": filters,
            "hiddenRows": hidden,
            "visibilityKnown": True,
        }

    plan = rewrite_visibility_formula(formula, "Sheet", resolve)
    register_visibility_functions(engine, plan)
    engine.set_formula("Sheet", 1, 4, plan.formula)
    engine.set_formula("Sheet", 1, 5, "=D1*2")
    value = engine.evaluate_cell("Sheet", 1, 4)
    if isinstance(value, (float, int)):
        assert engine.evaluate_cell("Sheet", 1, 5) == value * 2
    return value, plan


@pytest.mark.parametrize(
    "number,expected",
    [
        (1, 20),
        (2, 3),
        (3, 5),
        (4, 30),
        (5, 10),
        (6, 6000),
        (7, 10),
        (8, math.sqrt(200 / 3)),
        (9, 60),
        (10, 100),
        (11, 200 / 3),
    ],
)
def test_all_subtotal_operations_keep_reference_types(number, expected):
    result, _ = calculate(f"=SUBTOTAL({number},A1:A6)", [10, 20, 30, True, "40", None])
    assert result == pytest.approx(expected)


@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=SUBTOTAL(9,A1:A4)", 60),
        ("=SUBTOTAL(109,A1:A4)", 40),
        ("=AGGREGATE(9,0,A1:A4)", 60),
        ("=AGGREGATE(9,1,A1:A4)", 40),
        ("=AGGREGATE(9,4,A1:A4)", 120),
        ("=AGGREGATE(9,5,A1:A4)", 100),
    ],
)
def test_hidden_rows_and_nested_totals(formula, expected):
    result, plan = calculate(
        formula, [10, 20, 30, 60], hidden={2}, source_formulas={4: "=SUBTOTAL(9,A1:A3)"}
    )
    assert result == expected
    assert plan.evidence[0]["ranges"][0]["bounds"] == [1, 1, 4, 1]


def test_errors_propagate_unless_option_excludes_them():
    error = {"type": "Error", "kind": "Div"}
    values = [10, error, 20]
    assert calculate("=SUBTOTAL(9,A1:A3)", values)[0] == error
    assert calculate("=AGGREGATE(9,6,A1:A3)", values)[0] == 30
    assert calculate("=SUBTOTAL(2,A1:A3)", values)[0] == 2
    assert calculate("=SUBTOTAL(3,A1:A3)", values)[0] == 3
    assert calculate("=SUBTOTAL(109,A1)", [error], hidden={1})[0] == 0
    assert calculate("=AGGREGATE(9,6,A1)", [error])[0] == 0


@pytest.mark.parametrize(
    "formula", ["=AGGREGATE(9,6,A1:A2)", "=SUBTOTAL(2,A1:A2)", "=SUBTOTAL(3,A1:A2)"]
)
def test_uncomputed_engine_value_is_never_ignored(formula):
    error = {"type": "Error", "kind": "NImpl"}
    assert calculate(formula, [error, 5])[0] == error


@pytest.mark.parametrize(
    "number,expected",
    [
        (1, {"type": "Error", "kind": "Div"}),
        (2, 0),
        (3, 0),
        (4, 0),
        (5, 0),
        (6, 0),
        (7, {"type": "Error", "kind": "Div"}),
        (9, 0),
    ],
)
def test_empty_reference_result_is_typed(number, expected):
    assert calculate(f"=SUBTOTAL({number},A1:A3)", [None, None, None])[0] == expected


@pytest.mark.parametrize(
    "number,k,expected",
    [
        (12, None, 25),
        (13, None, {"type": "Error", "kind": "Na"}),
        (14, 2, 30),
        (15, 2, 20),
        (16, 0.25, 17.5),
        (17, 1, 17.5),
        (18, 0.25, 12.5),
        (19, 1, 12.5),
    ],
)
def test_order_statistics(number, k, expected):
    suffix = "" if k is None else f",{k}"
    assert (
        calculate(f"=AGGREGATE({number},6,A1:A4{suffix})", [10, 20, 30, 40])[0]
        == expected
    )


def test_active_filter_mode_must_be_unambiguous():
    assert (
        calculate("=SUBTOTAL(109,A1:A3)", [10, 20, 30], hidden={2}, filters=True)[0]
        == 40
    )
    with pytest.raises(VisibilityError, match="manually hidden"):
        calculate("=SUBTOTAL(9,A1:A3)", [10, 20, 30], filters=True)


def test_calls_inside_outer_arithmetic_and_single_cells():
    result, _ = calculate(
        "=SUM(SUBTOTAL(109,A1),AGGREGATE(9,5,A2:A3))+1", [10, 20, 30], hidden={2}
    )
    assert result == 41


@pytest.mark.parametrize(
    "formula",
    [
        "=SUBTOTAL(A1,A2:A3)",
        "=AGGREGATE(9,1+0,A1:A3)",
        "=AGGREGATE(14,3,A1:A3*2,1)",
        "=SUBTOTAL(9,A1:A20001)",
    ],
)
def test_unsupported_options_arrays_and_oversized_regions_are_explicit(formula):
    with pytest.raises(VisibilityError):
        calculate(formula, [10, 20, 30])
