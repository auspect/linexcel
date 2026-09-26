"""OOM recovery inventories source cells without inventing recalculated values."""

import io
import json
import os
import time
import zipfile

import pytest
from openpyxl import Workbook
from test_analysis_limits import replace_part

from linexcel import ExecutionPolicy, analyze
from linexcel.progress import Reporter
from linexcel.recovery import _decode_text, source_inventory


def package(sheets):
    book = Workbook()
    book.remove(book.active)
    for name, cells in sheets.items():
        sheet = book.create_sheet(name)
        for address, value in cells.items():
            sheet[address] = value
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


def recover(data, budget=100, **kwargs):
    return source_inventory(data, kwargs, budget, Reporter())


def test_sparse_cells_do_not_pay_for_empty_coordinates_and_budget_is_fair():
    data = package(
        {
            "Dense": {f"A{r}": r for r in range(1, 100)},
            "Sparse": {"XFD1048576": "=SUM(Dense!A:A)"},
        }
    )
    graph = recover(data, budget=4)
    assert len(graph["nodes"]) == 3
    distant = graph["nodes"][-1]
    assert distant["addr"] == "XFD1048576"
    assert distant["formula"] == "=SUM(Dense!A:A)"
    assert distant["value"] is None
    assert distant["steps"] is None
    assert graph["meta"]["stats"]["totalFormulas"] is None
    assert graph["meta"]["stats"]["sheets"][0]["truncated"]
    assert not graph["edges"]
    assert (
        graph["meta"]["analysisCoverage"]["dependencyCompleteness"] == "not_inspected"
    )


def test_stored_values_and_shared_formulas_are_not_recalculated():
    data = package({"S": {"A1": "=1+1", "A2": "=1+1", "B1": True, "C1": "hello"}})
    data = replace_part(
        data,
        "xl/worksheets/sheet1.xml",
        lambda xml: xml.replace(
            b"<f>1+1</f><v />",
            b'<f t="shared" si="0" ref="A1:A2">B1+1</f><v>99</v>',
            1,
        ).replace(b"<f>1+1</f><v />", b'<f t="shared" si="0"/><v>98</v>', 1),
    )
    nodes = {n["addr"]: n for n in recover(data)["nodes"]}
    assert nodes["A1"]["formula"] == "=B1+1"
    assert nodes["A2"]["formula"] == "=B2+1"
    assert nodes["A1"]["value"] == 99
    assert nodes["A2"]["value"] == 98
    assert nodes["B1"]["value"] is True
    assert nodes["C1"]["value"] == "hello"
    assert all(n["valueSource"] == "file" for n in nodes.values())


@pytest.mark.parametrize("key", ["max_cells_per_sheet", "max_nodes_per_sheet"])
def test_zero_limit_is_not_unlimited(key):
    graph = recover(package({"S": {"A1": "=2+3"}}), **{key: 0})
    assert not graph["nodes"]
    assert graph["meta"]["stats"]["sheets"][0]["truncated"]


@pytest.mark.parametrize("value", [-1, 3, True, 1.5])
def test_retry_limit_is_validated(value):
    with pytest.raises((ValueError, TypeError)):
        ExecutionPolicy(memory_retries=value)


def execution(status):
    return {
        "graph": {
            "nodes": [],
            "edges": [],
            "meta": {
                "warnings": ["original failure"],
                "execution": {
                    "status": status,
                    "failure": {"kind": status},
                    "elapsedSeconds": 0,
                },
            },
        }
    }


def test_retries_reduce_work_keep_memory_cap_and_share_time_budget(monkeypatch):
    from linexcel import execution as module

    calls = []
    effective = []

    def attempt(data, kwargs, policy, *, recovery_cells=None):
        calls.append((policy, recovery_cells))
        if recovery_cells is not None:
            effective.append(
                recover(
                    package({"S": {f"A{r}": r for r in range(1, 600)}}),
                    budget=recovery_cells,
                    **kwargs,
                )
            )
        time.sleep(0.01)
        return execution("completed" if len(calls) == 3 else "memory_limit")

    monkeypatch.setattr(module, "_run_isolated_once", attempt)
    result = module.run_isolated(b"", {}, ExecutionPolicy(seconds=10, memory_mb=256))
    assert len(calls) == 3
    assert calls[1][1] == 2048
    assert calls[2][1] == 512
    assert [len(g["nodes"]) for g in effective] == [400, 100]
    assert all(p.memory_mb == 256 for p, _ in calls)
    assert calls[0][0].seconds > calls[1][0].seconds > calls[2][0].seconds
    evidence = result["graph"]["meta"]["execution"]
    assert evidence["status"] == "memory_limit"
    assert evidence["recovery"]["recalculated"] is False
    assert len(evidence["attempts"]) == 3
    json.dumps(evidence)  # no self-referential attempt history


@pytest.mark.parametrize("status", ["timed_out", "cancelled", "crashed", "completed"])
def test_only_proven_memory_failure_retries(monkeypatch, status):
    from linexcel import execution as module

    calls = []

    def attempt(*args, **kwargs):
        calls.append(kwargs)
        return execution(status)

    monkeypatch.setattr(module, "_run_isolated_once", attempt)
    module.run_isolated(b"", {}, ExecutionPolicy())
    assert len(calls) == 1


@pytest.mark.parametrize(
    "kwargs, policy",
    [
        ({"targets": ["S!A1"]}, ExecutionPolicy()),
        ({}, ExecutionPolicy(memory_retries=0)),
        ({}, ExecutionPolicy(seconds=0)),
    ],
)
def test_target_scope_opt_out_and_expired_budget_prevent_retry(
    monkeypatch, kwargs, policy
):
    from linexcel import execution as module

    calls = []

    def attempt(*args, **kwargs):
        calls.append(kwargs)
        return execution("memory_limit")

    monkeypatch.setattr(module, "_run_isolated_once", attempt)
    module.run_isolated(b"", kwargs, policy)
    assert len(calls) == 1


@pytest.mark.skipif(os.name != "nt", reason="Native OOM under Windows Job memory cap")
def test_real_native_oom_recovers_formulas_without_reentering_evaluation():
    data = package({"S": {"A1": "=SUM(SEQUENCE(100000,100))", "B1": "=2+3"}})
    result = analyze(data, execution=ExecutionPolicy(seconds=30, memory_mb=256))
    evidence = result.graph["meta"]["execution"]
    assert evidence["status"] == "memory_limit"
    assert evidence["recovery"]["status"] == "completed"
    assert len(evidence["attempts"]) == 2
    assert evidence["attempts"][0]["failure"]["kind"] == "memory_limit"
    assert len(result.nodes) == 2
    assert all(n["value"] is None for n in result.nodes)
    assert all(n["steps"] is None for n in result.nodes)
    assert not result.edges


@pytest.mark.parametrize("status", ["cancelled", "timed_out", "error", "crashed"])
def test_failed_recovery_stops_without_hiding_initial_failure(monkeypatch, status):
    from linexcel import execution as module

    calls = []

    def attempt(*args, **kwargs):
        calls.append(kwargs)
        return execution("memory_limit" if len(calls) == 1 else status)

    monkeypatch.setattr(module, "_run_isolated_once", attempt)
    result = module.run_isolated(b"", {}, ExecutionPolicy())
    evidence = result["graph"]["meta"]["execution"]
    assert len(calls) == 2
    assert evidence["status"] == (
        "cancelled" if status == "cancelled" else "memory_limit"
    )
    assert evidence["attempts"][0]["failure"]["kind"] == "memory_limit"
    assert evidence["recovery"]["status"] == status


def test_shared_and_inline_text_exclude_phonetic_guides():
    data = package({"S": {"A1": "placeholder", "B1": "placeholder"}})
    xml = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c>'
        '<c r="B1" t="inlineStr"><is><r><t>東</t></r><r><t>京</t></r>'
        '<rPh sb="0" eb="2"><t>とうきょう</t></rPh></is></c>'
        "</row></sheetData></worksheet>"
    )
    data = replace_part(data, "xl/worksheets/sheet1.xml", lambda _: xml.encode())
    buffer = io.BytesIO(data)
    with zipfile.ZipFile(buffer, "a") as archive:
        archive.writestr(
            "xl/sharedStrings.xml",
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<si><t>東京</t><rPh sb="0" eb="2"><t>とうきょう</t></rPh></si></sst>',
        )
    nodes = recover(buffer.getvalue())["nodes"]
    assert [n["value"] for n in nodes] == ["東京", "東京"]
    assert all(n["valueSource"] == "file" for n in nodes)


@pytest.mark.parametrize(
    "encoded, expected",
    [
        ("line_x000D_break", "line\rbreak"),
        ("_x005F_x0041_", "_x0041_"),
        ("_xD83D__xDE00_", "😀"),
        ("_xD83D_", None),
    ],
)
def test_text_escapes_decode_once_and_remain_serializable(encoded, expected):
    assert _decode_text(encoded) == expected
    json.dumps(_decode_text(encoded), ensure_ascii=False).encode("utf-8")
