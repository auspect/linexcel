"""An index of an earlier schema waits, unused, for the user to reimport.

Index schema 5 has its own file name (``index-v5.sqlite``). A project imported
by an earlier linexcel keeps its ``index.sqlite`` on disk, unread: every view
that reads the index says the workbook must be reimported, and the library
leaves the project out of search. Nothing is deleted when the server starts;
the old index and the views derived from it go when the user launches that
project's reimport, which builds the new index from the kept workbook.
"""

import base64
import io
import time

from test_index_reads import _request, reads_workbook

from linexcel.lazy_store import INDEX_FILENAME
from linexcel.web import OUTDATED_INDEX_MESSAGE, create_app

_OLD_FILES = ("index.sqlite", "hierarchy.sqlite", "source-data.sqlite")


def _wait(app, task_id):
    deadline = time.monotonic() + 60
    while app.store.tasks[task_id]["status"] not in {"succeeded", "failed"}:
        assert time.monotonic() < deadline, "task did not finish"
        time.sleep(0.02)
    return app.store.tasks[task_id]


def _app(root):
    return create_app(root, resolve_user=lambda env: env.get("user"))


def _imported(app):
    stream = io.BytesIO()
    reads_workbook().save(stream)
    code, created = _request(
        app,
        "POST",
        "/api/projects",
        {
            "workbook": {
                "name": "reads.xlsx",
                "data": base64.b64encode(stream.getvalue()).decode(),
            }
        },
    )
    assert code in {200, 201, 202}, created
    task = _wait(app, created["task"]["id"])
    assert task["status"] == "succeeded"
    pid = created["project"]["id"]
    return task["owner"], pid, app.store.project_path(task["owner"], pid)


def test_an_earlier_index_waits_unused_until_the_project_is_reimported(tmp_path):
    app = _app(tmp_path)
    try:
        owner, pid, path = _imported(app)
    finally:
        app.close()
    # The project as an earlier linexcel left it, then the new server starts:
    # it reads nothing from the old index and deletes nothing.
    (path / INDEX_FILENAME).rename(path / "index.sqlite")
    for name in _OLD_FILES[1:]:
        (path / name).write_bytes(b"derived from the old index")
    app = _app(tmp_path)
    try:
        assert all((path / name).is_file() for name in _OLD_FILES)
        code, snapshot = _request(app, "GET", f"/api/projects/{pid}")
        assert code == 200
        assert snapshot["outdatedIndex"] is True
        assert snapshot["graph"] is None
        for view in ("nodes", "structure", "graph-view"):
            code, refused = _request(app, "GET", f"/api/projects/{pid}/{view}")
            assert code == 409
            assert OUTDATED_INDEX_MESSAGE in str(refused)
        code, refused = _request(
            app,
            "POST",
            f"/api/projects/{pid}/tasks",
            {"operation": "evaluate", "nodeId": "Calc!B1"},
        )
        assert code == 409 and OUTDATED_INDEX_MESSAGE in str(refused)
        # The library rebuilds its catalog in the background; the project is
        # listed there as needing a rebuild, so search leaves it out.
        deadline = time.monotonic() + 30
        while True:
            status = app.store.library.query(owner, "status", {})
            flags = [
                source["rebuildRequired"]
                for source in status.get("revisions", [])
                if source["projectId"] == pid
            ]
            if flags or time.monotonic() > deadline:
                break
            app.store.library.refresh(owner)
            time.sleep(0.1)
        assert flags == [True]
        assert all((path / name).is_file() for name in _OLD_FILES)
        code, created = _request(
            app,
            "POST",
            f"/api/projects/{pid}/tasks",
            {"operation": "import", "force": True},
        )
        assert code in {200, 201, 202}, created
        # Launching the reimport deletes the old index and its views at once.
        assert not any((path / name).exists() for name in _OLD_FILES)
        assert _wait(app, created["task"]["id"])["status"] == "succeeded"
        assert (path / INDEX_FILENAME).is_file()
        code, snapshot = _request(app, "GET", f"/api/projects/{pid}")
        assert snapshot["outdatedIndex"] is False
        assert snapshot["graph"]["meta"]["nodeCount"] == 357
    finally:
        app.close()


def test_a_current_index_is_kept_until_its_reimport_replaces_it(tmp_path):
    app = _app(tmp_path)
    try:
        owner, pid, path = _imported(app)
        before = (path / INDEX_FILENAME).stat().st_mtime_ns
        code, created = _request(
            app,
            "POST",
            f"/api/projects/{pid}/tasks",
            {"operation": "import", "force": True},
        )
        assert code in {200, 201, 202}, created
        assert (path / INDEX_FILENAME).is_file()
        assert _wait(app, created["task"]["id"])["status"] == "succeeded"
        assert (path / INDEX_FILENAME).stat().st_mtime_ns >= before
    finally:
        app.close()
