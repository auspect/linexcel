"""Web contracts: isolated work, private ownership, restart and honest failures."""

from __future__ import annotations

import base64
import io
import json
import os
import threading
import time

import pytest
from openpyxl import Workbook

from linexcel.web import APIError, TaskStore, WebConfig, create_app


def workbook():
    book = Workbook()
    book.active["A1"] = 6
    book.active["B1"] = "=A1*7"
    stream = io.BytesIO()
    book.save(stream)
    return {
        "name": "example.xlsx",
        "data": base64.b64encode(stream.getvalue()).decode(),
    }


def request(app, method, path, body=None, user="alice", **extra):
    data = json.dumps(body).encode() if body is not None else b""
    env = {
        "REQUEST_METHOD": method,
        "PATH_INFO": path,
        "SCRIPT_NAME": "",
        "CONTENT_TYPE": "application/json",
        "CONTENT_LENGTH": str(len(data)),
        "wsgi.input": io.BytesIO(data),
        "wsgi.url_scheme": "http",
        "HTTP_HOST": "127.0.0.1:8765",
        "user": user,
    }
    env.update(extra)
    response = []
    payload = b"".join(
        app(env, lambda status, headers: response.extend([status, headers]))
    )
    return int(response[0].split()[0]), payload, dict(response[1])


def wait(store, task_id):
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        with store.lock:
            task = store.tasks[task_id]
            if task["status"] in {"succeeded", "failed", "cancelled"}:
                return task.copy()
        time.sleep(0.02)
    raise AssertionError("Task did not terminate")


@pytest.fixture
def app(tmp_path):
    app = create_app(tmp_path, resolve_user=lambda env: env.get("user"))
    yield app
    app.close()


def test_full_import_evaluate_reconnect_export_and_delete(app):
    code, body, _ = request(app, "POST", "/api/projects", {"workbook": workbook()})
    assert code == 200
    imported = json.loads(body)
    project = imported["project"]["id"]
    task = wait(app.store, imported["task"]["id"])
    assert task["status"] == "succeeded", task
    code, body, _ = request(app, "GET", f"/api/projects/{project}")
    graph = json.loads(body)["graph"]
    assert graph["meta"]["evaluated"] is False
    assert next(n for n in graph["nodes"] if n["id"] == "Sheet!B1")["value"] is None
    code, body, _ = request(
        app,
        "POST",
        f"/api/projects/{project}/tasks",
        {
            "operation": "evaluate",
            "nodeId": "Sheet!B1",
        },
    )
    assert code == 200
    task_id = json.loads(body)["task"]["id"]
    task = wait(app.store, task_id)
    assert task["status"] == "succeeded", task
    assert task["result"]["value"] == 42
    assert task["result"]["comparison"] == "unknown"
    _, body, _ = request(
        app,
        "POST",
        f"/api/projects/{project}/tasks",
        {
            "operation": "evaluate",
            "nodeId": "Sheet!B1",
        },
    )
    assert json.loads(body)["task"]["id"] == task_id
    _, body, _ = request(
        app,
        "POST",
        f"/api/projects/{project}/tasks",
        {"operation": "evaluate", "nodeId": "Sheet!B1", "force": True},
    )
    forced_id = json.loads(body)["task"]["id"]
    assert forced_id != task_id
    assert wait(app.store, forced_id)["result"]["value"] == 42
    _, body, _ = request(app, "GET", f"/api/projects/{project}/export")
    assert json.loads(body)["graph"] == graph
    assert request(app, "DELETE", f"/api/projects/{project}")[0] == 200
    assert request(app, "GET", f"/api/tasks/{task_id}")[0] == 404


def test_users_cannot_access_projects_tasks_or_exports(app):
    _, body, _ = request(app, "POST", "/api/projects", {"workbook": workbook()})
    imported = json.loads(body)
    project, task_id = imported["project"]["id"], imported["task"]["id"]
    wait(app.store, task_id)
    for method, path in [
        ("GET", f"/api/projects/{project}"),
        ("GET", f"/api/projects/{project}/export"),
        ("DELETE", f"/api/projects/{project}"),
        ("GET", f"/api/tasks/{task_id}"),
        ("POST", f"/api/tasks/{task_id}/cancel"),
    ]:
        assert request(app, method, path, user="bob")[0] == 404
    assert request(app, "GET", "/api/projects", user=None)[0] == 401
    assert json.loads(request(app, "GET", "/api/projects", user="bob")[1]) == {
        "projects": []
    }


def test_restart_reuses_latest_calculation_regardless_of_directory_order(
    tmp_path, monkeypatch
):
    from pathlib import Path

    config = WebConfig(tmp_path)
    first = TaskStore(config)
    try:
        created = first.create("alice", {"workbook": workbook()})
        wait(first, created["task"]["id"])
        project = created["project"]["id"]
        older = first.submit("alice", project, "evaluate", "Sheet!B1", {})
        assert wait(first, older["id"])["status"] == "succeeded"
        newer = first.submit("alice", project, "evaluate", "Sheet!B1", {}, force=True)
        assert wait(first, newer["id"])["status"] == "succeeded"
    finally:
        first.close()

    original_glob = Path.glob

    def newest_files_first(directory, pattern):
        paths = list(original_glob(directory, pattern))
        if pattern == "tasks/*/task.json":
            paths.sort(
                key=lambda path: json.loads(path.read_text("utf-8"))["createdAt"],
                reverse=True,
            )
        return iter(paths)

    monkeypatch.setattr(Path, "glob", newest_files_first)
    recovered = TaskStore(config)
    try:
        reused = recovered.submit("alice", project, "evaluate", "Sheet!B1", {})
        assert reused["id"] == newer["id"]
        assert reused["createdAt"] > older["createdAt"]
    finally:
        recovered.close()


def test_graph_asset_is_packaged_mounted_and_allowlisted(tmp_path):
    from pathlib import Path

    from linexcel import web

    app = create_app(
        tmp_path, mount_path="/tools/linexcel", resolve_user=lambda env: env.get("user")
    )
    expected = (Path(web.__file__).parent / "assets" / "cytoscape.min.js").read_bytes()
    try:
        for path, script_name in [
            ("/tools/linexcel/assets/cytoscape.min.js", ""),
            ("/assets/cytoscape.min.js", "/tools/linexcel"),
        ]:
            code, body, headers = request(app, "GET", path, SCRIPT_NAME=script_name)
            assert code == 200 and body == expected
            assert headers["Content-Type"].startswith("text/javascript")
            assert int(headers["Content-Length"]) == len(expected)
            code, body, headers = request(app, "HEAD", path, SCRIPT_NAME=script_name)
            assert code == 200 and body == b""
            assert int(headers["Content-Length"]) == len(expected)
        for path in [
            "/tools/linexcel/assets/../web.py",
            "/tools/linexcel/assets/app.html",
            "/tools/linexcel/assets/unknown.js",
            "/assets/cytoscape.min.js",
        ]:
            assert request(app, "GET", path)[0] == 404
        assert (
            request(app, "GET", "/tools/linexcel/assets/cytoscape.min.js", user=None)[0]
            == 401
        )
        assert (
            request(app, "POST", "/tools/linexcel/assets/cytoscape.min.js", {})[0]
            == 404
        )
    finally:
        app.close()


def test_uploaded_reference_change_never_reuses_old_calculation(app):
    main = Workbook()
    main.active["B1"] = "='[reference.xlsx]Sheet'!A1*2"
    stream = io.BytesIO()
    main.save(stream)
    primary = {
        "name": "main.xlsx",
        "data": base64.b64encode(stream.getvalue()).decode(),
    }
    revisions = []
    for value in (4, 9):
        reference = Workbook()
        reference.active["A1"] = value
        stream = io.BytesIO()
        reference.save(stream)
        body = {
            "workbook": primary,
            "references": [
                {
                    "name": "reference.xlsx",
                    "data": base64.b64encode(stream.getvalue()).decode(),
                }
            ],
        }
        imported = app.store.create("owner", body)
        assert wait(app.store, imported["task"]["id"])["status"] == "succeeded"
        project = imported["project"]
        assert project["name"] == "main.xlsx"
        revisions.append(project["revision"])
        task = app.store.submit("owner", project["id"], "evaluate", "Sheet!B1", {})
        completed = wait(app.store, task["id"])
        assert completed["status"] == "succeeded", completed
        assert completed["result"]["value"] == value * 2
        assert completed["patternId"].startswith("formula-sha256:")
    assert revisions[0] != revisions[1]


def test_mount_and_local_cookie(tmp_path):
    app = create_app(tmp_path, mount_path="/tools/linexcel")
    try:
        assert request(app, "GET", "/tools/linexcel")[0] == 303
        assert request(app, "GET", "", SCRIPT_NAME="/tools/linexcel")[0] == 303
        code, page, headers = request(app, "GET", "/tools/linexcel/")
        assert code == 200 and b"Linexcel" in page
        assert "HttpOnly; SameSite=Strict" in headers["Set-Cookie"]
        assert (
            request(app, "GET", "/api/config", SCRIPT_NAME="/tools/linexcel")[0] == 200
        )
        assert request(app, "GET", "/elsewhere/")[0] == 404
        assert (
            request(
                app,
                "POST",
                "/tools/linexcel/api/projects",
                {},
                HTTP_ORIGIN="https://attacker.example",
            )[0]
            == 403
        )
        assert (
            request(
                app,
                "POST",
                "/tools/linexcel/api/projects",
                {},
                CONTENT_TYPE="text/plain",
            )[0]
            == 415
        )
    finally:
        app.close()


def test_invalid_inputs_are_rejected_before_task_submission(app):
    for body in [
        [],
        {"workbook": {"name": "../bad.xlsx", "data": "UEs="}},
        {"workbook": workbook(), "budget": {"seconds": float("nan")}},
        {"workbook": workbook(), "budget": {"memoryMb": 4096}},
        {"workbook": workbook(), "references": [workbook()]},
    ]:
        assert request(app, "POST", "/api/projects", body)[0] == 400
    assert app.store.tasks == {}


def test_timeout_preserves_previous_graph_and_diagnostic(app):
    imported = app.store.create("owner", {"workbook": workbook()})
    project = imported["project"]["id"]
    wait(app.store, imported["task"]["id"])
    before = app.store.snapshot("owner", project)["graph"]
    task = app.store.submit(
        "owner", project, "evaluate", "Sheet!B1", {"seconds": 0.001}
    )
    failed = wait(app.store, task["id"])
    assert failed["status"] == "failed"
    assert failed["error"]["kind"] == "timeout"
    assert app.store.snapshot("owner", project)["graph"] == before


def test_global_memory_admission_and_cancellation(tmp_path, monkeypatch):
    entered, release = threading.Event(), threading.Event()
    store = TaskStore(WebConfig(tmp_path, total_memory_mb=512, max_workers=2))

    def worker(task, root, event):
        entered.set()
        while not release.wait(0.01) and not event.is_set():
            pass
        return {"result": {"nodes": [], "edges": [], "meta": {}}}

    monkeypatch.setattr(store, "_worker", worker)
    try:
        first = store.create("owner", {"workbook": workbook()})["task"]
        assert entered.wait(2)
        second = store.create("owner", {"workbook": workbook()})["task"]
        assert store.tasks[second["id"]]["status"] == "queued"
        assert store.used_memory == 512
        store.cancel("owner", second["id"])
        assert wait(store, second["id"])["status"] == "cancelled"
        release.set()
        assert wait(store, first["id"])["status"] == "succeeded"
    finally:
        release.set()
        store.close()


def test_persistent_results_and_single_coordinator(tmp_path):
    store = TaskStore(WebConfig(tmp_path))
    try:
        with pytest.raises(RuntimeError, match="coordinator"):
            TaskStore(WebConfig(tmp_path))
        created = store.create("owner", {"workbook": workbook()})
        task_id = created["task"]["id"]
        assert wait(store, task_id)["status"] == "succeeded"
    finally:
        store.close()
    restored = TaskStore(WebConfig(tmp_path))
    try:
        assert restored.get_task("owner", task_id)["status"] == "succeeded"
        path = restored.project_path("owner", created["project"]["id"])
        (path / "workbook.xlsx").write_bytes(b"modified")
        with pytest.raises(APIError, match="Sources modifiées"):
            restored.submit("owner", created["project"]["id"], "import", None, {})
    finally:
        restored.close()


def test_failed_result_write_releases_worker_capacity(tmp_path, monkeypatch):
    import linexcel.web as web

    original = web._write

    def fail_final_task(path, value):
        if path.name == "task.json" and value.get("status") == "succeeded":
            raise OSError("Disk full")
        return original(path, value)

    monkeypatch.setattr(web, "_write", fail_final_task)
    store = TaskStore(WebConfig(tmp_path))
    try:
        created = store.create("owner", {"workbook": workbook()})
        task = wait(store, created["task"]["id"])
        assert task["status"] == "failed"
        assert task["error"]["kind"] == "persistence_error"
        assert store.used_memory == 0
        assert store.snapshot("owner", created["project"]["id"])["graph"]
    finally:
        store.close()


@pytest.mark.skipif(os.name != "nt", reason="Windows legacy MAX_PATH regression")
def test_nested_windows_data_directory(tmp_path):
    store = TaskStore(WebConfig(tmp_path / ("nested-" * 12)))
    try:
        imported = store.create("a" * 64, {"workbook": workbook()})
        task = wait(store, imported["task"]["id"])
        assert task["status"] == "succeeded", task
        assert store.snapshot("a" * 64, imported["project"]["id"])["graph"]
    finally:
        store.close()
