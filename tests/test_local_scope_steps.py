"""Local variables cannot be evaluated as independent scratch expressions."""

import io
from unittest.mock import Mock

import formualizer as fz
import pytest
from openpyxl import Workbook

from linexcel.analyzer import analyze_workbook
from linexcel.decompose import _collect_step_exprs, _decompose


@pytest.mark.parametrize("targeted", [False, True])
def test_let_keeps_its_result_without_fabricating_name_errors_in_steps(targeted):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Model"
    sheet["A1"], sheet["A2"] = 2, 8
    sheet["B1"] = "=LET(total,SUM(A1:A2),n,COUNT(A1:A2),total/n)+SUM(A1:A2)"
    data = io.BytesIO()
    workbook.save(data)
    graph = analyze_workbook(
        data.getvalue(), targets=["Model!B1"] if targeted else None
    )["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    assert node["value"] == 15
    local, ordinary = node["steps"]["children"]
    assert local["value"] == 5 and local["evaluated"]
    assert ordinary["value"] == 10 and ordinary["evaluated"]
    assert len(local["children"]) == 3
    for step in local["children"]:
        assert step["value"] is None
        assert not step["evaluated"]
        assert step["evaluationReason"] == "local-scope"


@pytest.mark.parametrize(
    "binding", ["LET(x,2,x+1)", "LAMBDA(x,x+1)", "_xlfn.LET(x,2,x+1)"]
)
def test_local_scope_never_reads_shadowed_workbook_names_or_uses_scratch(binding):
    ast = fz.parse("=" + binding).to_dict()
    resolver = Mock()
    tree = _decompose(ast, "Model", resolver, root_value=3)
    assert tree["value"] == 3
    assert _collect_step_exprs(ast, skip_root=True) == []
    resolver.eval_expr.assert_not_called()
    resolver.value.assert_not_called()
    assert tree["children"][0]["value"] is None
    assert not tree["children"][0]["evaluated"]


def test_nested_scope_does_not_leak_into_an_ordinary_sibling():
    ast = fz.parse("=LET(x,2,LET(x,3,x+1)+x)+SUM(1,2)").to_dict()
    expressions = _collect_step_exprs(ast, skip_root=True)
    assert len(expressions) == 2
    assert expressions[-1] == "SUM(1, 2)"
    assert "x + 1" not in expressions
