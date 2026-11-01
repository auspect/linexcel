"""Offline reports freeze selected evidence without evaluating or calling AI."""

import base64
import io
import json
from pathlib import Path

import pytest
from openpyxl import Workbook
from test_web import request, wait

from linexcel.web import APIError, create_app
from linexcel.web_report import normalize_options, prepare_report, write_report


@pytest.fixture
def report_project(tmp_path):
    app = create_app(tmp_path, resolve_user=lambda env: env.get("user"))
    book = Workbook()
    book.active["A1"] = '</script><img src="https://example.org/leak">'
    for row in range(2, 452):
        book.active.cell(row, 1, row)
    book.active["B1"] = "=SUM(A2:A451)"
    stream = io.BytesIO()
    book.save(stream)
    _, raw, _ = request(
        app,
        "POST",
        "/api/projects",
        {
            "workbook": {
                "name": "selected.xlsx",
                "data": base64.b64encode(stream.getvalue()).decode(),
            }
        },
    )
    created = json.loads(raw)
    task = wait(app.store, created["task"]["id"])
    assert task["status"] == "succeeded", task
    owner = app.store.tasks[task["id"]]["owner"]
    pid = created["project"]["id"]
    try:
        yield app, owner, pid
    finally:
        app.close()


@pytest.mark.parametrize(
    "options",
    [
        {},
        {"nodeIds": []},
        {"nodeIds": ["a", "a"]},
        {"nodeIds": [None]},
        {"nodeIds": ["a"], "captureIds": ["../x"]},
        {"nodeIds": ["a"], "language": "xx"},
        {"nodeIds": ["a"], "endpoint": "http://example.org"},
    ],
)
def test_report_rejects_invalid_selection(options):
    with pytest.raises(APIError):
        normalize_options(options)


def test_report_freezes_bounded_evidence_and_escapes_html(report_project):
    app, owner, pid = report_project
    task = app.store.submit(owner, pid, "evaluate", "Sheet!B1", {})
    calculation = wait(app.store, task["id"])
    assert calculation["result"]["value"] == sum(range(2, 452))
    before = set(app.store.tasks)
    with app.store.lock:
        snapshot = prepare_report(
            app.store,
            owner,
            pid,
            {"nodeIds": ["Sheet!A1", "Sheet!B1", "Sheet!A2:A451"], "language": "en"},
        )
    assert set(app.store.tasks) == before
    assert snapshot["records"][1]["calculation"]["id"] == task["id"]
    range_evidence = snapshot["records"][2]["ranges"][0]
    assert range_evidence["representedCellCount"] == 450
    assert len(range_evidence["members"]) == 200
    assert range_evidence["omitted"] == 250
    root = app.store.project_path(owner, pid) / "tasks" / "report-test"
    root.mkdir()
    result = write_report(snapshot, root)
    document = (root / result["file"]).read_text()
    assert "Source input" in document
    assert "No calculation recorded" in document
    assert "&lt;/script&gt;&lt;img" in document
    assert "<script" not in document
    assert '<img src="https://' not in document
    assert "Report" not in result["file"]
    assert result["revision"] == snapshot["project"]["revision"]
    assert result["bytes"] == (root / result["file"]).stat().st_size
    assert "200 / 450 represented cells included" in document
    assert "Sheet!A451" not in document  # outside the explicitly bounded context


def test_report_rejects_missing_node_and_foreign_project(report_project):
    app, owner, pid = report_project
    with pytest.raises(APIError, match="introuvable"):
        prepare_report(app.store, owner, pid, {"nodeIds": ["Sheet!Z999"]})
    with pytest.raises(APIError):
        prepare_report(app.store, "different-owner", pid, {"nodeIds": ["Sheet!A1"]})


def test_report_artifact_endpoint_and_download_ownership(report_project):
    app, owner, pid = report_project
    code, data, _ = request(
        app,
        "POST",
        f"/api/projects/{pid}/tasks",
        {"operation": "export_html", "options": {"nodeIds": ["Sheet!B1"]}},
    )
    assert code == 200, data
    task = wait(app.store, json.loads(data)["task"]["id"])
    assert task["status"] == "succeeded", task
    code, document, headers = request(app, "GET", f"/api/tasks/{task['id']}/artifact")
    assert code == 200 and b"Sheet!B1" in document
    assert "attachment" in headers["Content-Disposition"]
    assert (
        request(app, "GET", f"/api/tasks/{task['id']}/artifact", user="bob")[0] == 404
    )


def test_report_reads_as_file_without_javascript_or_network(report_project, tmp_path):
    playwright = pytest.importorskip("playwright.sync_api")
    app, owner, pid = report_project
    snapshot = prepare_report(app.store, owner, pid, {"nodeIds": ["Sheet!B1"]})
    result = write_report(snapshot, tmp_path)
    with playwright.sync_playwright() as p:
        try:
            browser = p.chromium.launch()
        except playwright.Error as exc:
            pytest.skip(f"Chromium unavailable: {exc}")
        try:
            for width, height in ((1440, 900), (390, 844)):
                context = browser.new_context(
                    java_script_enabled=False,
                    viewport={"width": width, "height": height},
                )
                requests = []
                context.on("request", lambda request: requests.append(request.url))
                page = context.new_page()
                page.goto(Path(tmp_path / result["file"]).as_uri())
                assert page.locator("h2").first.inner_text() == "Sheet!B1"
                assert page.locator("body").bounding_box()["width"] == width
                assert all(url.startswith("file:") for url in requests)
                context.close()
        finally:
            browser.close()
