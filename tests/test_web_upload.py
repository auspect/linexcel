"""Large uploads stay bounded, resumable and isolated from other owners."""

from __future__ import annotations

import base64
import io
import json
import threading
import time

import pytest
from test_web import request, wait, workbook

from linexcel.web import TaskStore, WebConfig, create_app
from linexcel.web_upload import CHUNK_BYTES, COPY_BYTES


@pytest.fixture
def app(tmp_path):
    instance = create_app(tmp_path, resolve_user=lambda env: env.get("user"))
    yield instance
    instance.close()


def begin(app, size=None, **options):
    raw = base64.b64decode(workbook()["data"])
    code, response, _ = request(
        app,
        "POST",
        "/api/uploads",
        {
            "workbook": {"name": "example.xlsx", "size": size or len(raw)},
            **options,
        },
    )
    assert code == 200, response
    return json.loads(response)["upload"], raw


def put(app, upload, raw, offset=0, index=0, user="alice", **extra):
    environ = {
        "CONTENT_TYPE": "application/octet-stream",
        "CONTENT_LENGTH": str(len(raw)),
        "HTTP_X_UPLOAD_OFFSET": str(offset),
        "wsgi.input": io.BytesIO(raw),
        **extra,
    }
    return request(
        app,
        "PUT",
        f"/api/uploads/{upload['id']}/files/{index}",
        user=user,
        **environ,
    )


def test_upload_chunks_complete_idempotently_and_evaluate(app):
    upload, raw = begin(app)
    midpoint = len(raw) // 2
    assert put(app, upload, raw[:midpoint])[0] == 200
    assert put(app, upload, raw[midpoint:], offset=midpoint)[0] == 200
    path = f"/api/uploads/{upload['id']}/complete"
    code, body, _ = request(app, "POST", path, {})
    assert code == 200, body
    imported = json.loads(body)
    assert wait(app.store, imported["task"]["id"])["status"] == "succeeded"
    assert (
        json.loads(request(app, "POST", path, {})[1])["project"] == imported["project"]
    )
    assert request(app, "POST", path, {}, user="bob")[0] == 404
    project = imported["project"]["id"]
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
    task = wait(app.store, json.loads(body)["task"]["id"])
    assert task["result"]["value"] == 42
    assert not app.store.uploads.sessions


def test_search_scoped_upload_stays_out_of_library_until_kept(app):
    from openpyxl import load_workbook

    book = load_workbook(io.BytesIO(base64.b64decode(workbook()["data"])))
    book.active["A1"] = "SearchScopedNeedle"
    stream = io.BytesIO()
    book.save(stream)
    raw = stream.getvalue()
    _, body, _ = request(
        app,
        "POST",
        "/api/uploads",
        {
            "workbook": {"name": "scoped.xlsx", "size": len(raw)},
            "searchScoped": True,
        },
    )
    upload = json.loads(body)["upload"]
    assert put(app, upload, raw)[0] == 200
    code, body, _ = request(app, "POST", f"/api/uploads/{upload['id']}/complete", {})
    assert code == 200, body
    imported = json.loads(body)
    assert imported["project"]["searchScopeId"] == upload["id"]
    assert wait(app.store, imported["task"]["id"])["status"] == "succeeded"
    project_id = imported["project"]["id"]

    projects = json.loads(request(app, "GET", "/api/projects")[1])["projects"]
    assert project_id not in {project["id"] for project in projects}
    assert (
        request(
            app,
            "GET",
            "/api/library/search",
            QUERY_STRING="q=SearchScopedNeedle&kinds=cell",
        )[0]
        == 200
    )
    default_search = json.loads(
        request(
            app,
            "GET",
            "/api/library/search",
            QUERY_STRING="q=SearchScopedNeedle&kinds=cell",
        )[1]
    )
    assert all(
        hit["revision"]["projectId"] != project_id for hit in default_search["results"]
    )

    scoped = {}
    for _ in range(100):
        code, body, _ = request(
            app,
            "GET",
            "/api/library/search",
            QUERY_STRING=f"scope={upload['id']}&project={project_id}&q=SearchScopedNeedle&kinds=cell",
        )
        assert code == 200, body
        scoped = json.loads(body)
        if scoped["results"]:
            break
        time.sleep(0.05)
    assert scoped["results"], scoped
    assert scoped["results"][0]["revision"]["projectId"] == project_id
    recovered = json.loads(request(app, "GET", "/api/library/attachments")[1])
    assert recovered["attachments"][0]["scopeId"] == upload["id"]
    assert recovered["attachments"][0]["expiresAt"] > time.time()
    scoped_status = json.loads(
        request(
            app,
            "GET",
            "/api/library/status",
            QUERY_STRING=f"scope={upload['id']}",
        )[1]
    )
    assert any(item["projectId"] == project_id for item in scoped_status["revisions"])
    assert (
        request(
            app,
            "POST",
            f"/api/library/attachments/{upload['id']}/keep",
            {},
        )[0]
        == 200
    )
    projects = json.loads(request(app, "GET", "/api/projects")[1])["projects"]
    assert project_id in {project["id"] for project in projects}
    owner = app.store.tasks[imported["task"]["id"]]["owner"]
    kept_path = app.store.root / owner / "projects" / project_id
    kept = json.loads((kept_path / "project.json").read_text())
    assert "searchScopeId" not in kept and "searchScopeExpiresAt" not in kept
    kept["createdAt"] = 0
    kept["searchScopeExpiresAt"] = 0
    (kept_path / "project.json").write_text(json.dumps(kept), encoding="utf-8")
    app.store.library.cleanup_expired_attachments(owner)
    assert kept_path.exists()

    second, second_raw = begin(app, searchScoped=True)
    assert put(app, second, second_raw)[0] == 200
    second_import = json.loads(
        request(app, "POST", f"/api/uploads/{second['id']}/complete", {})[1]
    )
    assert wait(app.store, second_import["task"]["id"])["status"] == "succeeded"
    assert request(app, "DELETE", f"/api/library/attachments/{second['id']}")[0] == 200
    assert (
        request(
            app,
            "GET",
            "/api/library/status",
            QUERY_STRING=f"scope={second['id']}",
        )[0]
        == 404
    )


def test_failed_search_scoped_import_can_be_removed_from_attachment_route(app):
    broken = b"PK\x03\x04not-a-zip"
    upload, _ = begin(app, size=len(broken), searchScoped=True)
    assert put(app, upload, broken)[0] == 200
    code, body, _ = request(app, "POST", f"/api/uploads/{upload['id']}/complete", {})
    assert code == 200, body
    imported = json.loads(body)
    task = wait(app.store, imported["task"]["id"])
    assert task["status"] == "failed"
    project_id = imported["project"]["id"]
    owner = app.store.tasks[imported["task"]["id"]]["owner"]
    project_path = app.store.root / owner / "projects" / project_id
    assert (project_path / "project.json").exists()
    assert request(app, "DELETE", f"/api/library/attachments/{upload['id']}")[0] == 200
    assert not project_path.exists()
    assert (
        json.loads(request(app, "GET", "/api/library/attachments")[1])["attachments"]
        == []
    )


def test_expired_search_attachment_is_listed_for_cleanup_and_freed(app):
    upload, raw = begin(app, searchScoped=True)
    assert put(app, upload, raw)[0] == 200
    imported = json.loads(
        request(app, "POST", f"/api/uploads/{upload['id']}/complete", {})[1]
    )
    assert wait(app.store, imported["task"]["id"])["status"] == "succeeded"
    project_id = imported["project"]["id"]
    owner = app.store.tasks[imported["task"]["id"]]["owner"]
    project_path = app.store.root / owner / "projects" / project_id
    manifest_path = project_path / "project.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["searchScopeExpiresAt"] = time.time() - 1
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
    listed = json.loads(request(app, "GET", "/api/library/attachments")[1])
    assert listed["attachments"] == []
    assert not project_path.exists()


def test_expired_attachment_cannot_be_kept_and_upload_cleans_its_quota(app):
    upload, raw = begin(app, searchScoped=True)
    assert put(app, upload, raw)[0] == 200
    imported = json.loads(
        request(app, "POST", f"/api/uploads/{upload['id']}/complete", {})[1]
    )
    assert wait(app.store, imported["task"]["id"])["status"] == "succeeded"
    task = app.store.tasks[imported["task"]["id"]]
    path = app.store.root / task["owner"] / "projects" / imported["project"]["id"]
    manifest = json.loads((path / "project.json").read_text())
    manifest["searchScopeExpiresAt"] = time.time() - 1
    (path / "project.json").write_text(json.dumps(manifest), encoding="utf-8")
    assert (
        request(app, "POST", f"/api/library/attachments/{upload['id']}/keep", {})[0]
        == 410
    )
    begin(app)  # Cleanup also runs before a new upload's quota reservation.
    assert not path.exists()


def test_attachment_cleanup_is_bounded_and_never_deletes_a_kept_project(
    app, monkeypatch
):
    import hashlib

    import linexcel.web_library as library_module

    owner = hashlib.sha256(b"alice").hexdigest()
    root = app.store.root / owner / "projects"
    root.mkdir(parents=True, exist_ok=True)
    for number in range(6):
        path = root / f"{number:032x}"
        path.mkdir()
        manifest = {
            "id": path.name,
            "filename": "example.xlsx",
            "createdAt": 0,
            "searchScopeExpiresAt": 0,
        }
        if number < 5:
            manifest["searchScopeId"] = f"{number:032x}"
        (path / "project.json").write_text(json.dumps(manifest), encoding="utf-8")
    monkeypatch.setattr(library_module, "SEARCH_ATTACHMENT_CLEANUP_LIMIT", 2)
    for _ in range(4):
        result = app.store.library.cleanup_expired_attachments(owner)
        assert result["scanned"] <= 2
    assert [p.parent.name for p in root.glob("*/project.json")] == [f"{5:032x}"]


def test_failed_chunk_rolls_back_and_resumes_without_duplicate_bytes(app):
    upload, raw = begin(app)
    assert put(app, upload, raw[:100])[0] == 200
    assert put(app, upload, raw[100:120], offset=100, CONTENT_LENGTH="40")[0] == 400
    stored = json.loads(request(app, "GET", f"/api/uploads/{upload['id']}")[1])[
        "upload"
    ]
    assert stored["files"][0]["received"] == 100
    session = app.store.uploads.sessions[upload["id"]]
    assert (app.store.uploads._path(session) / "workbook.xlsx").stat().st_size == 100
    assert put(app, upload, raw[100:], offset=99)[0] == 409
    assert put(app, upload, raw[100:], offset=100)[0] == 200
    assert request(app, "POST", f"/api/uploads/{upload['id']}/complete", {})[0] == 200


def test_cancel_during_chunk_removes_partial_file_and_reservation(app):
    upload, raw = begin(app)
    entered, release = threading.Event(), threading.Event()

    class SlowStream:
        def read(self, size):
            entered.set()
            assert release.wait(5)
            return raw[:size]

    result = []
    thread = threading.Thread(
        target=lambda: result.append(
            put(app, upload, raw, **{"wsgi.input": SlowStream()})
        )
    )
    thread.start()
    try:
        assert entered.wait(5)
        assert request(app, "DELETE", f"/api/uploads/{upload['id']}")[0] == 200
    finally:
        release.set()
        thread.join(5)
    assert result[0][0] == 409
    assert not app.store.uploads.sessions
    assert not list(app.store.root.glob("*/uploads/*"))


def test_upload_access_size_type_origin_and_cleanup(app):
    upload, raw = begin(app)
    path = f"/api/uploads/{upload['id']}"
    for method, suffix, body in [
        ("GET", "", None),
        ("DELETE", "", None),
        ("POST", "/complete", {}),
    ]:
        assert request(app, method, path + suffix, body, user="bob")[0] == 404
    assert put(app, upload, raw, user="bob")[0] == 404
    assert put(app, upload, raw, HTTP_ORIGIN="https://evil.example")[0] == 403
    assert put(app, upload, raw + b"too much")[0] == 413
    assert put(app, upload, b"not Excel")[0] == 400
    assert request(app, "POST", path + "/complete", {})[0] == 409
    assert request(app, "DELETE", path)[0] == 200
    assert not list(app.store.root.glob("*/uploads/*"))
    assert app.store.uploads.remaining_bytes() == 0
    assert request(app, "GET", path)[0] == 404


def test_declared_sizes_reserve_storage_for_concurrent_uploads(tmp_path):
    instance = create_app(tmp_path, max_storage_mb=4, resolve_user=lambda env: "alice")
    try:
        upload, _ = begin(instance, size=3 * 1048576)
        code, _, _ = request(
            instance,
            "POST",
            "/api/uploads",
            {
                "workbook": {"name": "other.xlsx", "size": 2 * 1048576},
            },
        )
        assert code == 413
        assert request(instance, "DELETE", f"/api/uploads/{upload['id']}")[0] == 200
        begin(instance, size=2 * 1048576)
    finally:
        instance.close()


def test_large_stream_is_only_read_in_bounded_buffers(app):
    size = 20 * 1048576
    upload, _ = begin(app, size=size)

    class GeneratedStream:
        def __init__(self, offset):
            self.offset = offset

        def read(self, length):
            assert 0 < length <= COPY_BYTES
            result = (
                (b"PK" + b"0" * (length - 2)) if self.offset == 0 else b"0" * length
            )
            self.offset += length
            return result

    for offset in range(0, size, CHUNK_BYTES):
        length = min(CHUNK_BYTES, size - offset)
        code, body, _ = request(
            app,
            "PUT",
            f"/api/uploads/{upload['id']}/files/0",
            CONTENT_TYPE="application/octet-stream",
            CONTENT_LENGTH=str(length),
            HTTP_X_UPLOAD_OFFSET=str(offset),
            **{"wsgi.input": GeneratedStream(offset)},
        )
        assert code == 200, body
    session = app.store.uploads.sessions[upload["id"]]
    assert (app.store.uploads._path(session) / "workbook.xlsx").stat().st_size == size
    assert app.store.uploads.remaining_bytes() == 0


def test_upload_resume_and_complete_after_restart(tmp_path):
    config = WebConfig(tmp_path)
    first = TaskStore(config)
    raw = base64.b64decode(workbook()["data"])
    try:
        upload = first.uploads.begin(
            "owner",
            {
                "workbook": {
                    "name": "book.xlsx",
                    "size": len(raw),
                }
            },
        )["upload"]
        first.uploads.append("owner", upload["id"], 0, 0, 100, io.BytesIO(raw[:100]))
        session = first.uploads.sessions[upload["id"]]
        # A crash may leave unacknowledged tail bytes beyond the durable offset.
        with (first.uploads._path(session) / "workbook.xlsx").open("ab") as target:
            target.write(b"uncommitted")
    finally:
        first.close()
    second = TaskStore(config)
    try:
        second.uploads.append(
            "owner", upload["id"], 0, 100, len(raw) - 100, io.BytesIO(raw[100:])
        )
        imported = second.uploads.complete("owner", upload["id"])
        assert wait(second, imported["task"]["id"])["status"] == "succeeded"
    finally:
        second.close()
    third = TaskStore(config)
    try:
        assert (
            third.uploads.complete("owner", upload["id"])["project"]
            == imported["project"]
        )
    finally:
        third.close()


def test_expired_upload_releases_disk_and_reservation(app):
    upload, raw = begin(app)
    put(app, upload, raw)
    app.store.uploads.sessions[upload["id"]]["updatedAt"] = time.time() - 90000
    replacement, _ = begin(app)
    assert upload["id"] not in app.store.uploads.sessions
    assert replacement["id"] in app.store.uploads.sessions


def test_local_import_is_loopback_only_and_copies_immutable_source(tmp_path):
    raw = base64.b64decode(workbook()["data"])
    source = tmp_path / "source.xlsx"
    source.write_bytes(raw)
    instance = create_app(tmp_path / "data", allow_local_import=True)
    cookie = "linexcel_session=" + "a" * 64
    try:
        for environ in [
            {},
            {"REMOTE_ADDR": "192.0.2.1"},
            {"REMOTE_ADDR": "127.0.0.1", "HTTP_HOST": "attacker.example:8765"},
        ]:
            assert (
                request(
                    instance,
                    "POST",
                    "/api/projects/local",
                    {"path": str(source)},
                    HTTP_COOKIE=cookie,
                    **environ,
                )[0]
                == 403
            )
        code, body, _ = request(
            instance,
            "POST",
            "/api/projects/local",
            {"path": str(source)},
            HTTP_COOKIE=cookie,
            REMOTE_ADDR="127.0.0.1",
        )
        assert code == 200, body
        imported = json.loads(body)
        source.write_bytes(b"modified")
        task = wait(instance.store, imported["task"]["id"])
        assert task["status"] == "succeeded", task
        stored = next(instance.store.root.glob("*/projects/*/workbook.xlsx"))
        assert stored.read_bytes() == raw
    finally:
        instance.close()
    embedded = create_app(
        tmp_path / "embedded", allow_local_import=True, resolve_user=lambda env: "owner"
    )
    try:
        assert (
            request(
                embedded,
                "POST",
                "/api/projects/local",
                {"path": str(source)},
                REMOTE_ADDR="127.0.0.1",
            )[0]
            == 403
        )
    finally:
        embedded.close()


def test_metadata_and_legacy_json_are_bounded_before_reading(app):
    class Unreadable:
        def read(self, length):
            raise AssertionError("Oversized requests must be rejected before reading")

    for path, length in [("/api/uploads", 65537), ("/api/projects", 8 * 1048576 + 1)]:
        assert (
            request(
                app,
                "POST",
                path,
                CONTENT_LENGTH=str(length),
                **{"wsgi.input": Unreadable()},
            )[0]
            == 413
        )


def test_serve_has_professional_upload_defaults_and_configurable_quota():
    from linexcel.cli import _build_parser

    args = _build_parser().parse_args(["serve"])
    assert args.max_upload_mb == 256
    assert args.max_storage_mb == 2048
    assert args.no_local_import is False
    args = _build_parser().parse_args(
        ["serve", "--max-upload-mb", "512", "--no-local-import"]
    )
    assert args.max_upload_mb == 512 and args.no_local_import
