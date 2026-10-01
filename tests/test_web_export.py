"""Full project downloads must not inherit the browser's page limits."""

import base64
import io
import json
import sqlite3
from pathlib import Path

import pytest
from openpyxl import Workbook
from test_web import request, wait

from linexcel.web import APIError, create_app
from linexcel.web_export import CHUNK_BYTES, ProjectExport


def test_complete_streamed_export_keeps_all_nodes_edges_and_task_results(
    tmp_path, monkeypatch
):
    app = create_app(tmp_path, resolve_user=lambda env: env.get("user"))
    try:
        book = Workbook()
        for row in range(1, 451):
            book.active.cell(row, 1, row)
        book.active["B1"] = "=SUM(A1:A450)"
        stream = io.BytesIO()
        book.save(stream)
        _, data, _ = request(
            app,
            "POST",
            "/api/projects",
            {
                "workbook": {
                    "name": "many.xlsx",
                    "data": base64.b64encode(stream.getvalue()).decode(),
                }
            },
        )
        created = json.loads(data)
        assert wait(app.store, created["task"]["id"])["status"] == "succeeded"
        pid = created["project"]["id"]
        owner = app.store.tasks[created["task"]["id"]]["owner"]
        task = app.store.submit(owner, pid, "evaluate", "Sheet!B1", {})
        calculated = wait(app.store, task["id"])
        assert calculated["result"]["value"] == 101475
        page = app.store.snapshot(owner, pid)
        assert len(page["graph"]["nodes"]) == 200
        assert all(t["result"] is None for t in page["tasks"])

        def forbidden(*args, **kwargs):
            raise AssertionError(
                "Export must not load an API graph page or task result"
            )

        monkeypatch.setattr(app.store, "_graph", forbidden)
        monkeypatch.setattr(app.store, "_task_result", forbidden)
        exported = ProjectExport(app.store, owner, pid)
        chunks = list(exported)
        assert max(map(len, chunks)) <= CHUNK_BYTES
        assert len(chunks) > 1
        assert exported.closed and not app.store._project_exports
        result = json.loads(b"".join(chunks))
        assert result["export"]["scope"] == "complete"
        graph = result["graph"]
        assert len(graph["nodes"]) == 452  # 450 inputs, one formula, one range
        assert {f"Sheet!A{i}" for i in range(1, 451)} <= {
            n["id"] for n in graph["nodes"]
        }
        assert graph["pagination"]["hasMore"] is False
        assert len(graph["edges"]) == graph["meta"]["edgeCount"]
        assert (
            next(t for t in result["tasks"] if t["id"] == task["id"])["result"]["value"]
            == 101475
        )
        assert request(app, "GET", f"/api/projects/{pid}/export", user="bob")[0] == 404
        code, data, headers = request(app, "GET", f"/api/projects/{pid}/export")
        assert code == 200 and json.loads(data)["graph"]["nodes"] == graph["nodes"]
        assert headers["Content-Disposition"].endswith(f'{pid}.json"')
        assert "Content-Length" not in headers
        assert request(app, "HEAD", f"/api/projects/{pid}/export")[1] == b""
        assert not app.store._project_exports

        result_path = app.store.task_export(owner, task["id"])
        original_open = Path.open
        opened_results = []

        class TrackedResult:
            def __init__(self, source):
                self.source = source

            def __enter__(self):
                return self

            def __exit__(self, *_):
                self.source.close()

            def read(self, size):
                assert 0 < size <= CHUNK_BYTES
                return self.source.read(size)

        def track_open(path, *args, **kwargs):
            source = original_open(path, *args, **kwargs)
            if path == result_path:
                tracked = TrackedResult(source)
                opened_results.append(tracked)
                return tracked
            return source

        monkeypatch.setattr(Path, "open", track_open)
        export = ProjectExport(app.store, owner, pid)
        while not opened_results:
            next(export)
        assert not opened_results[0].source.closed
        export.close()
        assert opened_results[0].source.closed
        assert not app.store._project_exports
        monkeypatch.setattr(Path, "open", original_open)

        for started in (False, True):
            export = ProjectExport(app.store, owner, pid)
            index = export.index
            if started:
                next(export)
            with pytest.raises(APIError) as blocked:
                app.store.delete(owner, pid)
            assert blocked.value.status == 409
            with pytest.raises(APIError) as blocked:
                app.store.submit(owner, pid, "import", None, {})
            assert blocked.value.status == 409
            export.close()
            export.close()
            assert not app.store._project_exports
            with pytest.raises(sqlite3.ProgrammingError):
                index.db.execute("SELECT 1")
        app.store.delete(owner, pid)
    finally:
        app.close()
