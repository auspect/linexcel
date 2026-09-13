"""Iteration limits are not convergence: keep unverified results out of reports."""

import io
import re
import zipfile

import pytest
from openpyxl import Workbook
from openpyxl.workbook.properties import CalcProperties

from linexcel.analyzer import analyze_workbook
from linexcel.engine import _open_workbook


def iterative_workbook(
    *, count=3, delta=0.001, iterate=True, cache=False, divergent=False, broken=False
):
    workbook = Workbook()
    workbook.calculation = CalcProperties(
        iterate=iterate, iterateCount=count, iterateDelta=delta
    )
    sheet = workbook.active
    sheet.title = "Model"
    # Financing and interest expense depend on one another; C1 uses the result.
    sheet["A1"] = "=B1+1" if divergent else "=100+B1/2"
    sheet["B1"] = "=A1" if divergent else "=A1*0.1"
    sheet["C1"] = "=A1-B1"
    sheet["D1"] = "=2+3"
    if broken:
        sheet["E1"] = "=IFERROR(NoSheet!A1,7)"
    original = io.BytesIO()
    workbook.save(original)
    if not cache:
        return original.getvalue()
    output = io.BytesIO()
    with zipfile.ZipFile(original) as source, zipfile.ZipFile(output, "w") as dest:
        for part in source.infolist():
            payload = source.read(part.filename)
            if part.filename == "xl/worksheets/sheet1.xml":
                for address, value in {"A1": 105, "B1": 10.5, "C1": 94.5}.items():
                    pattern = rf'(<c r="{address}"[^>]*><f>.*?</f>)<v(?:\s*/>|></v>)'
                    payload = re.sub(
                        pattern.encode(),
                        rf"\g<1><v>{value}</v>".encode(),
                        payload,
                    )
            dest.writestr(part, payload)
    return output.getvalue()


@pytest.mark.parametrize("targeted", [False, True])
@pytest.mark.parametrize("cache", [False, True])
@pytest.mark.parametrize("divergent", [False, True])
def test_iteration_cap_keeps_cycles_and_dependents_unverified(
    targeted, cache, divergent
):
    graph = analyze_workbook(
        iterative_workbook(cache=cache, divergent=divergent),
        targets=["Model!C1", "Model!D1"] if targeted else None,
    )["graph"]
    nodes = {node.get("addr"): node for node in graph["nodes"] if node.get("formula")}
    for address, stored in {"A1": 105, "B1": 10.5, "C1": 94.5}.items():
        assert nodes[address]["value"] == (stored if cache else None)
        assert nodes[address].get("valueSource") == ("file" if cache else None)
        assert nodes[address]["steps"] is None
    assert nodes["D1"]["value"] == 5
    assert nodes["D1"]["valueSource"] == "engine"
    assert any("did not converge in 3" in w for w in graph["meta"]["warnings"])


def test_declared_iteration_count_and_delta_reach_the_engine():
    engine = _open_workbook(iterative_workbook(count=2, delta=0.00001), parallel=False)
    engine.evaluate_all()
    telemetry = engine.last_cycle_telemetry()
    assert telemetry.max_passes_single_scc == 2
    assert telemetry.capped_sccs == 1
    engine = _open_workbook(iterative_workbook(count=100, delta=1000), parallel=False)
    engine.evaluate_all()
    telemetry = engine.last_cycle_telemetry()
    assert telemetry.max_passes_single_scc <= 2
    assert telemetry.converged_sccs == 1


@pytest.mark.parametrize("targeted", [False, True])
def test_a_converged_declared_iteration_can_be_reported(targeted):
    graph = analyze_workbook(
        iterative_workbook(count=100, delta=0.000001),
        targets=["Model!C1"] if targeted else None,
    )["graph"]
    nodes = {node.get("addr"): node for node in graph["nodes"] if node.get("formula")}
    assert nodes["A1"]["value"] == pytest.approx(100 / 0.95, abs=0.000001)
    assert nodes["C1"]["valueSource"] == "engine"
    assert not any("did not converge" in w for w in graph["meta"]["warnings"])


def test_undeclared_iteration_does_not_approve_a_circular_result():
    graph = analyze_workbook(iterative_workbook(count=100, iterate=False, cache=True))[
        "graph"
    ]
    node = next(node for node in graph["nodes"] if node.get("addr") == "A1")
    assert node["valueSource"] == "file"
    assert any(
        "without workbook iteration enabled" in w for w in graph["meta"]["warnings"]
    )


@pytest.mark.parametrize("targeted", [False, True])
def test_failed_evaluation_does_not_recover_a_cycle_as_a_number(targeted):
    graph = analyze_workbook(
        iterative_workbook(broken=True, divergent=True),
        targets=["Model!C1", "Model!D1", "Model!E1"] if targeted else None,
    )["graph"]
    nodes = {node.get("addr"): node for node in graph["nodes"] if node.get("formula")}
    assert nodes["A1"]["value"] is None
    assert nodes["C1"]["value"] is None
    assert nodes["A1"]["steps"] is None
    assert nodes["D1"]["value"] == 5
    assert nodes["E1"]["value"] == 7
