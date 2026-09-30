"""Capture pixels stay on disk; requests and AI read only the selected image."""

from __future__ import annotations

import base64
import io
import json
from pathlib import Path

from test_web import request, wait, workbook

from linexcel.web import create_app

PNG = b"\x89PNG\r\n\x1a\nselected image"


def test_capture_urls_stream_owned_images_and_ai_reads_only_selection(
    tmp_path, monkeypatch
):
    app = create_app(
        tmp_path, ai_enabled=True, resolve_user=lambda env: env.get("user")
    )
    try:
        _, body, _ = request(app, "POST", "/api/projects", {"workbook": workbook()})
        imported = json.loads(body)
        assert wait(app.store, imported["task"]["id"])["status"] == "succeeded"
        project = imported["project"]["id"]
        seen = []

        def worker(task, root, event):
            if task["operation"] == "describe_capture":
                seen.append(json.loads((root / "ai_input.json").read_text("utf-8")))
                return {"result": {"markdown": "Selected image"}}
            assert task["operation"] == "capture"
            images = root / "capture-files"
            images.mkdir()
            (images / "0.png").write_bytes(PNG)
            (images / "0.thumb.png").write_bytes(PNG + b"thumb")
            with (images / "1.png").open("wb") as stream:
                stream.write(PNG)
                stream.seek(9 * 1048576)
                stream.write(b"0")
            return {
                "result": {
                    "screenshots": [
                        {
                            "sheet": "Sheet",
                            "name": "one.png",
                            "file": "0.png",
                            "thumbnail": "0.thumb.png",
                        },
                        {"sheet": "Sheet", "name": "large.png", "file": "1.png"},
                    ],
                    "notice": "Rendered workbook",
                }
            }

        monkeypatch.setattr(app.store, "_worker", worker)
        _, body, _ = request(
            app, "POST", f"/api/projects/{project}/tasks", {"operation": "capture"}
        )
        capture = wait(app.store, json.loads(body)["task"]["id"])
        assert capture["status"] == "succeeded", capture
        shots = capture["result"]["screenshots"]
        assert len(json.dumps(capture)) < 4000
        assert "data" not in shots[0] and "file" not in shots[0]
        assert shots[0]["thumbnailUrl"].endswith("?thumbnail=1")
        url = "/" + shots[0]["url"]
        code, body, headers = request(app, "GET", url)
        assert code == 200 and body == PNG
        assert headers["Content-Type"] == "image/png"
        assert request(app, "HEAD", url)[1] == b""
        assert request(app, "GET", url, QUERY_STRING="thumbnail=1")[1] == PNG + b"thumb"
        assert request(app, "GET", url, user="bob")[0] == 404
        assert request(app, "GET", url.rsplit("/", 1)[0] + "/99")[0] == 404
        assert (
            request(app, "GET", url.rsplit("/", 1)[0] + "/../workbook.xlsx")[0] == 404
        )

        def describe(index):
            return request(
                app,
                "POST",
                f"/api/projects/{project}/tasks",
                {
                    "operation": "describe_capture",
                    "options": {"captureTaskId": capture["id"], "captureIndex": index},
                },
            )

        code, body, _ = describe(0)
        assert code == 200, body
        assert wait(app.store, json.loads(body)["task"]["id"])["status"] == "succeeded"
        assert base64.b64decode(seen[0]["image"].split(",", 1)[1]) == PNG
        assert describe(1)[0] == 413
        assert len(seen) == 1
        assert request(app, "DELETE", f"/api/projects/{project}")[0] == 200
        assert request(app, "GET", url)[0] == 404
    finally:
        app.close()


def test_capture_worker_writes_metadata_and_small_thumbnail(tmp_path, monkeypatch):
    from PIL import Image

    from linexcel import insights, web_worker

    root = tmp_path / "project" / "tasks" / "task"
    root.mkdir(parents=True)
    (root / "request.json").write_text(json.dumps({"operation": "capture"}), "utf-8")
    actual = io.BytesIO()
    Image.new("RGB", (1000, 800), "white").save(actual, "PNG")

    def render(source, filename, outdir):
        assert isinstance(source, Path)
        outdir.mkdir()
        result = outdir / "sheet.png"
        result.write_bytes(actual.getvalue())
        return {"Sheet": [result]}

    monkeypatch.setattr(insights, "render_workbook_screenshots", render)
    monkeypatch.setattr("sys.argv", ["web_worker", str(root)])
    web_worker.main()
    payload = json.loads((root / "result.json").read_text("utf-8"))
    shot = payload["result"]["screenshots"][0]
    assert set(shot) == {"sheet", "name", "file", "thumbnail"}
    assert (root / "capture-files" / shot["file"]).read_bytes() == actual.getvalue()
    with Image.open(root / "capture-files" / shot["thumbnail"]) as thumb:
        assert thumb.width <= 320 and thumb.height <= 240
    assert len(json.dumps(payload)) < 1000
