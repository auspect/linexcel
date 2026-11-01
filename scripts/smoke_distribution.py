"""Exercise an installed distribution, with no checkout imports or AI runtime."""

from __future__ import annotations

import base64
import io
import json
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook

import linexcel
from linexcel.web import create_app


def request(app, method, path, body=None, *, query=""):
    raw = json.dumps(body).encode() if body is not None else b""
    response = []
    data = b"".join(
        app(
            {
                "REQUEST_METHOD": method,
                "PATH_INFO": path,
                "QUERY_STRING": query,
                "SCRIPT_NAME": "",
                "CONTENT_LENGTH": str(len(raw)),
                "CONTENT_TYPE": "application/json",
                "wsgi.input": io.BytesIO(raw),
                "wsgi.url_scheme": "http",
                "HTTP_HOST": "127.0.0.1:8765",
            },
            lambda status, headers: response.append((status, headers)),
        )
    )
    assert response[0][0].startswith("200"), (path, response[0][0], data)
    return data


def wait(app, task_id):
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        task = json.loads(request(app, "GET", f"/api/tasks/{task_id}"))["task"]
        if task["status"] in {"failed", "cancelled", "succeeded"}:
            assert task["status"] == "succeeded", task
            return task
        time.sleep(0.05)
    raise TimeoutError(task_id)


def main():
    source = Path(__file__).resolve().parents[1] / "src"
    assert not Path(linexcel.__file__).resolve().is_relative_to(source), (
        "Run with an isolated installed wheel/sdist, not the editable checkout"
    )
    assets = Path(linexcel.__file__).parent / "assets"
    assert (assets / "report.css").is_file()
    assert len(list((assets / "prompts").glob("*/review.md"))) >= 9
    with tempfile.TemporaryDirectory() as folder:
        subprocess.run(
            [sys.executable, "-m", "linexcel", "serve", "--help"],
            check=True,
            cwd=folder,
            stdout=subprocess.DEVNULL,
        )
        book = Workbook()
        book.active["A1"] = 6
        book.active["B1"] = "=A1*7"
        stream = io.BytesIO()
        book.save(stream)
        app = create_app(Path(folder) / "workspace", resolve_user=lambda env: "smoke")
        try:
            assert b"Linexcel" in request(app, "GET", "/")
            assert len(request(app, "GET", "/assets/cytoscape.min.js")) > 1000
            imported = json.loads(
                request(
                    app,
                    "POST",
                    "/api/projects",
                    {
                        "workbook": {
                            "name": "smoke.xlsx",
                            "data": base64.b64encode(stream.getvalue()).decode(),
                        }
                    },
                )
            )
            wait(app, imported["task"]["id"])
            project = imported["project"]["id"]
            source_rows = json.loads(
                request(
                    app,
                    "GET",
                    f"/api/projects/{project}/datasette/source/sheet_001.json",
                    query="_shape=objects",
                )
            )["rows"]
            assert source_rows == [{"row": 1, "A": 6, "B": None}], source_rows
            result = json.loads(
                request(
                    app,
                    "POST",
                    f"/api/projects/{project}/tasks",
                    {"operation": "evaluate", "nodeId": "Sheet!B1"},
                )
            )
            assert wait(app, result["task"]["id"])["result"]["value"] == 42
            exported = json.loads(
                request(app, "GET", f"/api/projects/{project}/export")
            )
            assert any(n["id"] == "Sheet!B1" for n in exported["graph"]["nodes"])
            result = json.loads(
                request(
                    app,
                    "POST",
                    f"/api/projects/{project}/tasks",
                    {"operation": "export_html", "options": {"nodeIds": ["Sheet!B1"]}},
                )
            )
            task = wait(app, result["task"]["id"])
            assert b"Sheet!B1" in request(
                app, "GET", f"/api/tasks/{task['id']}/artifact"
            )
        finally:
            app.close()
    print(
        f"Installed linexcel {linexcel.__version__}: CLI, HTTP, import, source data, calculation, exports passed"
    )


if __name__ == "__main__":
    main()
