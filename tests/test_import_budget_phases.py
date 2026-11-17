"""An import publishes within its time budget, whatever phase runs out (lot 2
ter, budget step).

The budget used to be checked only while sheets were read. When the
dependency analysis of a model with many formulas outlasted the task, the
coordinator stopped the worker: nothing was published and no reimport was
proposed. These tests run the real worker on the ``serve`` path with one
phase slowed down at least ten times beyond the budget, and check states,
never durations.
"""

import base64
import io
import json
import sys
import time

import pytest
from openpyxl import Workbook

from linexcel.web import create_app

#: The import's time budget in these tests: room for a worker to start on a
#: slow Windows machine and to publish within the final margin.
BUDGET = 20
#: Each slowed step sleeps this long; the steps slowed in a test add up to
#: at least ten times ``BUDGET``.
STEP = 1.0


def call(app, method, path, body=None, query=""):
    data = json.dumps(body or {}).encode()
    response = []
    result = b"".join(
        app(
            {
                "REQUEST_METHOD": method,
                "PATH_INFO": path,
                "QUERY_STRING": query,
                "CONTENT_TYPE": "application/json",
                "CONTENT_LENGTH": str(len(data)),
                "wsgi.input": io.BytesIO(data),
                "user": "alice",
            },
            lambda status, headers: response.append(status),
        )
    )
    return int(response[0].split()[0]), json.loads(result)


def finish(app, task_id, seconds=BUDGET + 120):
    # Generous: the coordinator itself ends an import at its budget.
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        owner = app.store.tasks[task_id]["owner"]
        task = app.store.get_task(owner, task_id)
        if task["status"] in {"failed", "succeeded", "cancelled"}:
            return task
        time.sleep(0.05)
    pytest.fail("task did not finish")


def slowed_worker(monkeypatch, patch):
    """Run the import worker after ``patch``, Python source run first."""
    import linexcel._process_budget as process_budget

    def command(module, root, memory_mb):
        bootstrap = (
            f"{patch}\n"
            "import runpy, sys\n"
            f"sys.argv = [{module!r}, sys.argv[1]]\n"
            f"runpy.run_module({module!r}, run_name='__main__')\n"
        )
        return [sys.executable, "-c", bootstrap, str(root)]

    monkeypatch.setattr(process_budget, "worker_command", command)


#: Dependency analysis slowed by ``STEP`` a formula outside the first sheet.
SLOW_ANALYSIS = f"""
import time
import linexcel.lazy as lazy
_signature = lazy._pattern_signature
def _slow(node, nodes):
    if node.get("sheet") != "First":
        time.sleep({STEP})
    return _signature(node, nodes)
lazy._pattern_signature = _slow
"""


def model_book():
    """First reads values and a formula of Third; Slow and Third hold far
    more formulas than the budget can analyse once slowed."""
    book = Workbook()
    first = book.active
    first.title = "First"
    first["A1"] = "=Third!B1*2"
    first["A2"] = "=Third!A1+1"
    first["A3"] = "=SUM(Third!B1:B3)"
    slow = book.create_sheet("Slow")
    third = book.create_sheet("Third")
    # Analysis needs at least 2 x 120 x STEP s: twelve times the budget.
    for row in range(1, 121):
        slow.cell(row, 1, f"=Third!B{row}+{row}")
        third.cell(row, 1, f"=B{row}*3")
        third.cell(row, 2, row * 10)
    stream = io.BytesIO()
    book.save(stream)
    return base64.b64encode(stream.getvalue()).decode()


def upload(app, data, seconds=BUDGET):
    code, created = call(
        app,
        "POST",
        "/api/projects",
        {
            "workbook": {"name": "model.xlsx", "data": data},
            "budget": {"seconds": seconds},
        },
    )
    assert code == 200, created
    return created["project"]["id"], finish(app, created["task"]["id"])


def meta_of(app, project):
    code, snapshot = call(app, "GET", f"/api/projects/{project}")
    assert code == 200, snapshot
    return (snapshot.get("graph") or {}).get("meta") or {}


def evaluate(app, project, node_id):
    code, submitted = call(
        app,
        "POST",
        f"/api/projects/{project}/tasks",
        {"operation": "evaluate", "nodeId": node_id},
    )
    assert code == 200, submitted
    task = finish(app, submitted["task"]["id"])
    assert task["status"] == "succeeded", task
    code, result = call(app, "GET", f"/api/tasks/{submitted['task']['id']}/export")
    assert code == 200, result
    return result


def node(app, project, node_id):
    code, found = call(
        app, "GET", f"/api/projects/{project}/node", query="nodeId=" + node_id
    )
    assert code == 200, found
    return found.get("node", found)


@pytest.fixture
def app(tmp_path):
    application = create_app(tmp_path, resolve_user=lambda env: env.get("user"))
    try:
        yield application
    finally:
        application.close()


def test_an_analysis_past_the_budget_publishes_what_it_analyzed(app, monkeypatch):
    """T1. The analysis stops at the deadline; the index is published with
    every value; the formulas left are marked not analyzed, their sheets
    partial for the time budget, and a calculation reaching one is refused.
    """
    slowed_worker(monkeypatch, SLOW_ANALYSIS)
    project, task = upload(app, model_book())
    assert task["status"] == "succeeded", task
    meta = meta_of(app, project)
    details = {s["name"]: s for s in meta["sheetDetails"]}
    assert details["First"]["status"] == "complete"
    for name in ("Slow", "Third"):
        assert details[name]["status"] == "partial", details[name]
        assert details[name]["reason"] == "time_budget"
        assert details[name]["unanalyzedFormulas"] >= 1
    # Most of Slow and all of Third were left.
    assert details["Third"]["unanalyzedFormulas"] == 120
    assert meta["status"] == "partial"
    # T2. The proposal counts the analysis, not only the reading. Too few
    # formulas were analyzed to measure their rate: it says it is uncertain.
    estimate = meta["importEstimate"]
    assert estimate["cause"] == ["analysis"]
    assert estimate["formulasAnalyzed"] < estimate["formulasPredicted"] == 243
    assert estimate["phases"]["analysis"] > 0
    assert not estimate["formulaRateMeasured"] and estimate["uncertain"]
    assert estimate["seconds"] >= BUDGET + 60
    timings = meta["importTimings"]
    assert timings["formulasNotAnalyzed"] == 243 - estimate["formulasAnalyzed"]
    assert any(
        d.get("code") == "formulas_not_analyzed" and d.get("sheet") == "Third"
        for d in meta["diagnostics"]
    )
    # Values stay, as saved values; a formula not analyzed says so.
    assert node(app, project, "Third!B7")["value"] == 70
    left = node(app, project, "Third!A7")
    assert any("not analyzed" in d for d in left["diagnostics"])
    assert left["dependencies"] == []
    # Refused, with the reason named, never calculated without its sources.
    refused = evaluate(app, project, "Third!A7")
    assert refused["status"] != "completed" and refused["value"] is None
    assert any("not analyzed" in d for d in refused["diagnostics"])
    reaching = evaluate(app, project, "First!A2")
    assert reaching["status"] != "completed" and reaching["value"] is None
    assert any("not analyzed" in d for d in reaching["diagnostics"])
    # An analyzed formula reading only values of that sheet is calculated.
    read = evaluate(app, project, "First!A1")
    assert (read["status"], read["value"]) == ("completed", 20)
    total = evaluate(app, project, "First!A3")
    assert (total["status"], total["value"]) == ("completed", 60)


#: Region description predicted far beyond any budget, and truly that slow:
#: describing the large sheet would sleep ten times the budget. A sheet too
#: small to hold a region is still looked at, which costs nothing.
SLOW_REGIONS = f"""
import time
import linexcel.data_regions as regions
import linexcel.lazy_store as lazy_store
lazy_store.ImportBudget.VALUE_SECONDS = 1.0
_detect = regions.detect_regions
def _slow(store, sheets=None):
    if sheets is None or "Table" in sheets:
        time.sleep({10 * BUDGET})
    return _detect(store, sheets)
regions.detect_regions = _slow
"""


def table_book():
    book = Workbook()
    sheet = book.active
    sheet.title = "Table"
    sheet.append(["Year", "Amount"])
    for row in range(2, 6_002):
        sheet.append([2000 + row % 20, row])
    book.create_sheet("Total")["A1"] = "=SUM(Table!B2:B6001)"
    stream = io.BytesIO()
    book.save(stream)
    return base64.b64encode(stream.getvalue()).decode()


def test_regions_that_do_not_fit_the_time_left_are_skipped(app, monkeypatch):
    """T5. Data regions change no value: when their predicted description
    does not fit the time left, they are skipped and the index says so."""
    slowed_worker(monkeypatch, SLOW_REGIONS)
    project, task = upload(app, table_book())
    assert task["status"] == "succeeded", task.get("error")
    meta = meta_of(app, project)
    regions = meta["dataRegions"]
    assert regions["regions"] == 0
    assert regions["skipped"] == {"reason": "time_budget", "sheets": ["Table"]}
    # Values and formulas are complete: only the descriptions are missing.
    assert {s["status"] for s in meta["sheetDetails"]} == {"complete"}
    total = evaluate(app, project, "Total!A1")
    assert total["status"] == "completed"


#: The VBA phase, which cannot be stopped, sleeps ten times the budget.
SLOW_VBA = f"""
import time
import linexcel.macros as macros
_index_vba = macros.index_vba
def _slow(*args, **kwargs):
    time.sleep({10 * BUDGET})
    return _index_vba(*args, **kwargs)
macros.index_vba = _slow
"""


def test_an_import_that_still_times_out_says_what_to_try(app, monkeypatch):
    """T6. A phase that cannot be stopped can still outlast the budget: the
    task times out and publishes nothing. The error used to be the bare word
    "timeout"; it now says that no new index was published, whether the
    previous one remains, and proposes a larger budget, without claiming an
    estimate. The previous index stays available."""
    project, first = upload(app, model_book(), seconds=600)
    assert first["status"] == "succeeded", first.get("error")
    slowed_worker(monkeypatch, SLOW_VBA)
    code, submitted = call(
        app,
        "POST",
        f"/api/projects/{project}/tasks",
        {"operation": "import", "budget": {"seconds": BUDGET}, "force": True},
    )
    assert code == 200, submitted
    task = finish(app, submitted["task"]["id"])
    assert task["status"] == "failed"
    error = task["error"]
    assert error["kind"] == "timeout"
    assert error["suggestedSeconds"] == 2 * BUDGET
    assert error["previousIndex"] is True
    message = error["message"]
    assert f"budget de {BUDGET} s" in message
    assert "l’index précédent reste consultable" in message
    assert f"le double ({2 * BUDGET} s)" in message
    assert "ce n’est pas une estimation" in message
    # The previous index is still published and read.
    code, nodes = call(app, "GET", f"/api/projects/{project}/nodes")
    assert code == 200, nodes
    assert node(app, project, "First!A1")["formula"] == "=Third!B1*2"
