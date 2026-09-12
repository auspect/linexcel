"""Compatibility function names keep native semantics and original lineage."""

import io
import math

import formualizer as fz
import pytest
from openpyxl import Workbook

from linexcel.analyzer import analyze_workbook
from linexcel.engine import _open_workbook


def workbook_bytes(formulas):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Model"
    for address, formula in formulas.items():
        sheet[address] = formula
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


@pytest.mark.parametrize(
    "argument",
    [
        "0",
        "-1",
        "2",
        "TRUE",
        '"0"',
        '"bad"',
        "1/0",
        "A20",
        "{0,1}",
        "{0;1}",
        '{0,"bad"}',
        "{0,1/0}",
        "A20:A21",
    ],
)
def test_normsdist_delegates_scalar_range_coercion_and_errors(argument):
    engine = _open_workbook(
        workbook_bytes(
            {
                "B1": f"=NORMSDIST({argument})",
                "C1": f"=NORM.S.DIST({argument},TRUE)",
            }
        ),
        True,
    )
    engine.evaluate_all()
    assert engine.get_value("Model", 1, 2) == engine.get_value("Model", 1, 3)


@pytest.mark.parametrize(
    "arguments",
    [
        "0,0,1,TRUE",
        "0,0,1,FALSE",
        "2,1,3,1",
        '"0",0,1,TRUE',
        '"bad",0,1,TRUE',
        "0,0,0,TRUE",
        "0,0,-1,TRUE",
        "1/0,0,1,TRUE",
        '0,0,1,"TRUE"',
        '0,0,1,"bad"',
        "A20,A21,1,FALSE",
        "{0,1},{0,1},{1,0},{TRUE,FALSE}",
        "0,0,1,{FALSE,TRUE}",
    ],
)
def test_normdist_delegates_all_four_native_arguments(arguments):
    engine = _open_workbook(
        workbook_bytes(
            {
                "B1": f"=NORMDIST({arguments})",
                "C1": f"=NORM.DIST({arguments})",
            }
        ),
        True,
    )
    engine.evaluate_all()
    assert engine.get_value("Model", 1, 2) == engine.get_value("Model", 1, 3)


@pytest.mark.parametrize("targeted", [False, True])
def test_aliases_are_calculated_through_guards_and_keep_original_formulas(targeted):
    formulas = {
        "A1": "=NORMSDIST(0)",
        "B1": "=IFERROR(A1,7)+NORMDIST(0,0,1,FALSE)",
        "C1": "=IFERROR(NORMSDIST(0),7)",
        "D1": "=IFERROR(NORMSDIST(1/0),7)",
        "E1": "=IFERROR(NORMDIST(0,0,0,TRUE),7)",
        "F1": "=UNKNOWN_STATISTIC(0)",
        "G1": "=MissingName+1",
        "H1": "=IFERROR(UNKNOWN_STATISTIC(0),9)",
    }
    graph = analyze_workbook(
        workbook_bytes(formulas),
        targets=[f"Model!{address}" for address in formulas] if targeted else None,
    )["graph"]
    nodes = {n.get("addr"): n for n in graph["nodes"] if n.get("formula")}
    for address, expected in {
        "A1": 0.5,
        "B1": 0.5 + 1 / math.sqrt(2 * math.pi),
        "C1": 0.5,
        "D1": 7,
        "E1": 7,
        "H1": 9,
    }.items():
        assert nodes[address]["value"] == pytest.approx(expected)
        assert nodes[address]["valueSource"] == "engine"
        assert (
            fz.parse(nodes[address]["formula"]).to_dict()
            == fz.parse(formulas[address]).to_dict()
        )
    assert nodes["F1"]["value"] == "#NAME?"
    assert nodes["G1"]["value"] == "#NAME?"
    normal = next(
        s for s in nodes["B1"]["steps"]["children"] if s["label"] == "NORMDIST"
    )
    assert normal["evaluated"]
    assert normal["value"] == pytest.approx(1 / math.sqrt(2 * math.pi))


def test_separate_workbooks_and_repeated_calls_do_not_reuse_stale_arguments():
    data = workbook_bytes({"A1": 0, "B1": "=NORMSDIST(A1)"})
    first, second = _open_workbook(data, True), _open_workbook(data, False)
    for value in [1, -1, 0, 2]:
        first.set_value("Model", 1, 1, value)
        assert first.evaluate_cell("Model", 1, 2) == pytest.approx(
            (1 + math.erf(value / math.sqrt(2))) / 2
        )
        assert second.evaluate_cell("Model", 1, 2) == 0.5


def test_targeting_only_a_dependent_keeps_the_legacy_precedent_in_scope():
    graph = analyze_workbook(
        workbook_bytes(
            {
                "A1": 0,
                "B1": "=NORMSDIST(A1)",
                "C1": "=IFERROR(B1,7)+1",
                "D1": "=NORMDIST(0,0,1,FALSE)",
            }
        ),
        targets=["Model!C1"],
    )["graph"]
    nodes = {n.get("addr"): n for n in graph["nodes"] if n.get("formula")}
    assert nodes["B1"]["value"] == 0.5
    assert nodes["C1"]["value"] == 1.5
    assert nodes["C1"]["valueSource"] == "engine"
    assert "D1" not in nodes


@pytest.mark.parametrize(
    "formula",
    ["=NORMSDIST()", "=NORMSDIST(1,2)", "=NORMDIST(0,0,1)", "=NORMDIST(0,0,1,TRUE,2)"],
)
def test_invalid_arity_remains_an_error(formula):
    engine = _open_workbook(workbook_bytes({"A1": formula}), True)
    engine.evaluate_all()
    assert engine.get_value("Model", 1, 1)["type"] == "Error"
