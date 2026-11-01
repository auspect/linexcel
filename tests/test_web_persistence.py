"""Local workspaces survive browsers/restarts and restore with isolated identities."""

from __future__ import annotations

import hashlib
import io
import json
import time
import uuid
import zipfile
from urllib.parse import urlsplit

import pytest
from test_web import request, wait, workbook

from linexcel.cli import _build_parser, main
from linexcel.web import APIError, TaskStore, WebConfig, _write, create_app
from linexcel.web_backup import backup, receive_restore, restore
from linexcel.web_local import COOKIE_SECONDS, _process_exists, session_link


def test_session_link_process_probe_never_accepts_invalid_ids():
    import os

    assert _process_exists(os.getpid())
    assert not _process_exists(0)
    assert not _process_exists(-1)
    assert not _process_exists(True)
    assert not _process_exists("123")


def test_close_withdraws_profile_before_unlock_and_retry_preserves_new_server(
    tmp_path, monkeypatch
):
    app = create_app(tmp_path, local_profile=True)
    app.local_profile.publish("http://127.0.0.1:8765/")
    server = tmp_path / ".access" / "server.json"
    original_browser_close = app.data_browser.close
    original_profile_close = app.local_profile.close

    def failed_browser_close():
        original_browser_close()
        raise RuntimeError("shutdown failure")

    def close_profile_with_lock():
        assert not app.store._lock_file.closed
        original_profile_close()

    with monkeypatch.context() as patch:
        patch.setattr(app.data_browser, "close", failed_browser_close)
        patch.setattr(app.local_profile, "close", close_profile_with_lock)
        with pytest.raises(RuntimeError, match="shutdown failure"):
            app.close()
    assert app.store._lock_file.closed
    assert not server.exists()
    replacement = create_app(tmp_path, local_profile=True)
    try:
        replacement.local_profile.publish("http://127.0.0.1:8766/")
        publication = server.read_bytes()
        app.close()
        assert server.read_bytes() == publication
        assert session_link(tmp_path).startswith("http://127.0.0.1:8766/")
    finally:
        replacement.close()


def test_profile_shutdown_error_still_releases_workspace_lock(tmp_path, monkeypatch):
    app = create_app(tmp_path, local_profile=True)
    app.local_profile.publish("http://127.0.0.1:8765/")

    def fail():
        raise OSError("publication unavailable")

    with monkeypatch.context() as patch:
        patch.setattr(app.local_profile, "close", fail)
        with pytest.raises(OSError, match="publication unavailable"):
            app.close()
    assert app.store._lock_file.closed
    replacement = create_app(tmp_path, local_profile=True)
    try:
        replacement.local_profile.publish("http://127.0.0.1:8766/")
        server = tmp_path / ".access" / "server.json"
        publication = server.read_bytes()
        app.close()
        assert server.read_bytes() == publication
    finally:
        replacement.close()


def local_request(app, method, path, body=None, **kwargs):
    return request(app, method, path, body, REMOTE_ADDR="127.0.0.1", **kwargs)


def login(app):
    url = app.local_profile.publish("http://127.0.0.1:8765/")
    token = urlsplit(url).fragment.removeprefix("session=")
    code, _, headers = local_request(app, "POST", "/api/session", {"token": token})
    assert code == 200
    return headers["Set-Cookie"].split(";", 1)[0], token


def imported(store, owner="alice"):
    created = store.create(owner, {"workbook": workbook()})
    assert wait(store, created["task"]["id"])["status"] == "succeeded"
    return created["project"]["id"]


def test_local_identity_requires_one_use_link_and_survives_restart(tmp_path):
    app = create_app(tmp_path, local_profile=True, retention_days=None)
    try:
        assert local_request(app, "GET", "/api/projects")[0] == 401
        config = json.loads(local_request(app, "GET", "/api/config")[1])
        assert config["session"] == {"localProfile": True, "required": True}
        cookie, token = login(app)
        assert local_request(app, "POST", "/api/session", {"token": token})[0] == 401
        pid = imported(app.store, app.local_profile.owner)
        identity = app.local_profile.identity
        code, _, headers = local_request(
            app, "GET", "/api/projects", HTTP_COOKIE=cookie
        )
        assert code == 200 and f"Max-Age={COOKIE_SECONDS}" in headers["Set-Cookie"]
        assert (
            request(
                app, "GET", "/api/projects", HTTP_COOKIE=cookie, REMOTE_ADDR="192.0.2.1"
            )[0]
            == 403
        )
        assert (
            local_request(
                app, "GET", "/api/projects", HTTP_COOKIE="linexcel_session=invalid"
            )[0]
            == 401
        )
    finally:
        app.close()
    app = create_app(tmp_path, local_profile=True, retention_days=None)
    try:
        assert app.local_profile.identity == identity
        projects = json.loads(
            local_request(app, "GET", "/api/projects", HTTP_COOKIE=cookie)[1]
        )
        assert [p["id"] for p in projects["projects"]] == [pid]
        another_cookie, _ = login(app)
        assert another_cookie == cookie
    finally:
        app.close()


def test_expired_link_and_origin_are_rejected_without_consuming_current_link(tmp_path):
    app = create_app(tmp_path, local_profile=True, retention_days=None)
    try:
        app.local_profile.publish("http://127.0.0.1:8765/")
        link = session_link(tmp_path)
        token = urlsplit(link).fragment.removeprefix("session=")
        assert (
            local_request(
                app,
                "POST",
                "/api/session",
                {"token": token},
                HTTP_ORIGIN="https://example.org",
            )[0]
            == 403
        )
        path = (
            tmp_path
            / ".access"
            / (hashlib.sha256(token.encode()).hexdigest() + ".json")
        )
        value = json.loads(path.read_text())
        value["expiresAt"] = time.time() - 1
        path.write_text(json.dumps(value))
        assert local_request(app, "POST", "/api/session", {"token": token})[0] == 401
        assert not path.exists()
    finally:
        app.close()
    with pytest.raises(ValueError, match="Start linexcel"):
        session_link(tmp_path)


def test_local_profile_never_overrides_host_authentication(tmp_path):
    with pytest.raises(ValueError, match="host authentication"):
        create_app(tmp_path, local_profile=True, resolve_user=lambda _: "alice")
    app = create_app(tmp_path, resolve_user=lambda env: env.get("user"))
    try:
        assert request(app, "GET", "/api/recovery")[0] == 403
        assert request(app, "POST", "/api/session", {"token": "a" * 64})[0] == 401
    finally:
        app.close()


def test_durable_default_is_launcher_only_and_expiration_is_explicit(tmp_path):
    assert _build_parser().parse_args(["serve"]).retention_days is None
    assert WebConfig(tmp_path).retention_days == 7
    store = TaskStore(WebConfig(tmp_path, retention_days=None))
    try:
        pid = imported(store)
        path = store.project_path("alice", pid) / "project.json"
        project = json.loads(path.read_text())
        project["createdAt"] = time.time() - 40 * 86400
        _write(path, project)
    finally:
        store.close()
    store = TaskStore(WebConfig(tmp_path, retention_days=None))
    assert store.list_projects("alice")[0]["id"] == pid
    store.close()
    store = TaskStore(WebConfig(tmp_path, retention_days=30))
    try:
        assert store.list_projects("alice") == []
    finally:
        store.close()


def test_corrupt_project_and_task_are_preserved_without_blocking_healthy_data(tmp_path):
    store = TaskStore(WebConfig(tmp_path, retention_days=None))
    try:
        healthy = imported(store)
        damaged = imported(store)
        path = store.project_path("alice", damaged) / "project.json"
        path.write_text('{"partial":')
        task_id = uuid.uuid4().hex
        task_path = (
            store.project_path("alice", healthy) / "tasks" / task_id / "task.json"
        )
        task_path.parent.mkdir()
        task_path.write_text("[]")
    finally:
        store.close()
    store = TaskStore(WebConfig(tmp_path, retention_days=None))
    try:
        assert [p["id"] for p in store.list_projects("alice")] == [healthy]
        assert {i["kind"] for i in store.recovery_issues} == {
            "project_manifest_unreadable",
            "task_manifest_unreadable",
        }
        assert path.read_text() == '{"partial":'
        assert task_path.read_text() == "[]"
        assert (path.parent / "workbook.xlsx").exists()
        with pytest.raises(APIError, match="conservés"):
            store.project_path("alice", damaged)
    finally:
        store.close()


@pytest.fixture
def workspace(tmp_path):
    store = TaskStore(WebConfig(tmp_path, retention_days=None))
    pid = imported(store)
    task = store.submit("alice", pid, "evaluate", "Sheet!B1", {})
    assert wait(store, task["id"])["result"]["value"] == 42
    try:
        yield store, pid, task["id"]
    finally:
        store.close()


def test_backup_restore_rebuilds_index_keeps_history_and_assigns_new_owner(workspace):
    store, pid, task_id = workspace
    archive = backup(store, "alice", pid)
    try:
        with zipfile.ZipFile(archive) as zf:
            manifest = json.loads(zf.read("manifest.json"))
            assert all(
                "owner" not in t and "projectId" not in t for t in manifest["tasks"]
            )
            assert not any(
                name.endswith("sqlite") or ".access" in name for name in zf.namelist()
            )
        restored = restore(store, "bob", archive)
        new_id = restored["project"]["id"]
        assert new_id != pid
        assert (
            restored["project"]["revision"]
            == store.list_projects("alice")[0]["revision"]
        )
        assert wait(store, restored["task"]["id"])["status"] == "succeeded"
        assert (store.project_path("bob", new_id) / "index.sqlite").is_file()
        history = [
            t for t in store.tasks.values() if t.get("restoredFromTask") == task_id
        ]
        assert len(history) == 1 and history[0]["owner"] == "bob"
        assert (
            history[0]["historical"]
            and history[0]["cacheVersion"] == "restored-historical"
        )
        assert store.get_task("bob", history[0]["id"])["result"]["value"] == 42
        with pytest.raises(APIError):
            store.get_task("alice", history[0]["id"])
    finally:
        archive.unlink(missing_ok=True)


def rewrite_archive(path, transform):
    with zipfile.ZipFile(path) as archive:
        members = {item.filename: archive.read(item) for item in archive.infolist()}
    manifest = json.loads(members.pop("manifest.json"))
    transform(manifest, members)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("manifest.json", json.dumps(manifest))
        for name, value in members.items():
            archive.writestr(name, value)


@pytest.mark.parametrize(
    "fault", ["version", "hash", "traversal", "missing", "nan", "duplicate"]
)
def test_corrupt_archives_never_publish_a_partial_workspace(workspace, fault):
    store, pid, _ = workspace
    archive = backup(store, "alice", pid)

    def corrupt(manifest, members):
        if fault == "version":
            manifest["version"] = 999
        elif fault == "hash":
            manifest["files"][0]["sha256"] = "0" * 64
        elif fault == "traversal":
            original = manifest["files"][0]["path"]
            manifest["files"][0]["path"] = "../escaped.xlsx"
            members["../escaped.xlsx"] = members.pop(original)
        elif fault == "missing":
            members.pop(manifest["files"][0]["path"])
        elif fault == "nan":
            manifest["tasks"][0]["createdAt"] = float("nan")
        else:
            manifest["files"].append(dict(manifest["files"][0]))

    try:
        rewrite_archive(archive, corrupt)
        with pytest.raises(APIError):
            restore(store, "bob", archive)
        assert store.list_projects("bob") == []
        assert not list((store.root / ".archives").glob("restore-stage-*"))
        assert not (store.root / "escaped.xlsx").exists()
    finally:
        archive.unlink(missing_ok=True)


def test_restore_rolls_back_if_reindex_cannot_be_scheduled(workspace, monkeypatch):
    store, pid, _ = workspace
    archive = backup(store, "alice", pid)
    before = set(store.tasks)

    def unavailable(*args, **kwargs):
        raise RuntimeError("scheduler unavailable")

    monkeypatch.setattr(store, "submit", unavailable)
    try:
        with pytest.raises(RuntimeError, match="scheduler unavailable"):
            restore(store, "bob", archive)
        assert store.list_projects("bob") == [] and set(store.tasks) == before
    finally:
        archive.unlink(missing_ok=True)


def test_raw_restore_enforces_body_and_disk_limits(workspace):
    store, pid, _ = workspace
    with pytest.raises(APIError, match="incomplet"):
        receive_restore(store, "bob", io.BytesIO(b"zip"), 20)
    assert not list((store.root / ".archives").glob("restore-*.zip"))
    with pytest.raises(APIError) as error:
        receive_restore(
            store, "bob", io.BytesIO(), store.config.max_upload_mb * 1048576 + 1
        )
    assert error.value.status == 413


def test_backup_http_ownership_download_cleanup_and_restore_endpoint(tmp_path):
    app = create_app(tmp_path, resolve_user=lambda env: env.get("user"))
    try:
        pid = imported(app.store, hashlib.sha256(b"alice").hexdigest())
        code, payload, headers = request(app, "GET", f"/api/projects/{pid}/backup")
        assert code == 200 and headers["Content-Type"] == "application/zip"
        assert not list((tmp_path / ".archives").glob("backup-*.zip"))
        assert request(app, "GET", f"/api/projects/{pid}/backup", user="bob")[0] == 404
        code, raw, _ = request(
            app,
            "POST",
            "/api/projects/restore",
            user="bob",
            CONTENT_TYPE="application/octet-stream",
            CONTENT_LENGTH=str(len(payload)),
            **{"wsgi.input": io.BytesIO(payload)},
        )
        assert code == 200, raw
        restored = json.loads(raw)
        assert wait(app.store, restored["task"]["id"])["status"] == "succeeded"
    finally:
        app.close()


def test_doctor_does_not_create_storage_or_request_generation(tmp_path, capsys):
    destination = tmp_path / "absent"
    assert main(["doctor", "--data-dir", str(destination)]) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["generationRequested"] is False
    assert result["checks"]["ai"]["status"] == "disabled"
    assert not destination.exists()


def test_legacy_recovery_is_an_explicit_copy_into_local_identity(tmp_path):
    old_owner = hashlib.sha256(b"old-browser-cookie").hexdigest()
    store = TaskStore(WebConfig(tmp_path, retention_days=None))
    old_id = imported(store, old_owner)
    old_path = store.project_path(old_owner, old_id)
    old_bytes = (old_path / "workbook.xlsx").read_bytes()
    store.close()
    app = create_app(tmp_path, local_profile=True, retention_days=None)
    try:
        cookie, _ = login(app)
        code, raw, _ = local_request(app, "GET", "/api/recovery", HTTP_COOKIE=cookie)
        assert code == 200
        entries = json.loads(raw)["projects"]
        assert len(entries) == 1 and "owner" not in entries[0]
        path = "/api/recovery/" + entries[0]["id"]
        code, raw, _ = local_request(app, "POST", path, {}, HTTP_COOKIE=cookie)
        assert code == 200, raw
        recovered = json.loads(raw)
        assert recovered["project"]["id"] != old_id
        assert wait(app.store, recovered["task"]["id"])["status"] == "succeeded"
        new_path = app.store.project_path(
            app.local_profile.owner, recovered["project"]["id"]
        )
        assert (new_path / "workbook.xlsx").read_bytes() == old_bytes
        assert (old_path / "workbook.xlsx").read_bytes() == old_bytes
        assert local_request(app, "POST", path, {}, HTTP_COOKIE=cookie)[0] == 409
    finally:
        app.close()


@pytest.mark.parametrize(
    "listing,expected",
    [
        ({"data": [{"id": "qwen3.8:latest"}]}, "ready"),
        ({"data": [{"id": "another-model"}]}, "unavailable"),
        ([], "unavailable"),
        ({"data": [None]}, "unavailable"),
    ],
)
def test_health_lists_models_without_generation_or_secret_disclosure(
    tmp_path, monkeypatch, listing, expected
):
    httpx = pytest.importorskip("httpx")
    from linexcel.web_health import directory_health

    original = httpx.Client
    calls = []

    def respond(request):
        calls.append(request)
        return httpx.Response(200, json=listing)

    monkeypatch.setattr(
        httpx,
        "Client",
        lambda **kwargs: original(**kwargs, transport=httpx.MockTransport(respond)),
    )
    result = directory_health(
        WebConfig(tmp_path, ai_enabled=True, ai_api_key="private-test-key")
    )
    assert result["checks"]["ai"]["status"] == expected
    assert len(calls) == 1 and calls[0].method == "GET"
    assert calls[0].url.path == "/v1/models" and calls[0].content == b""
    assert "private-test-key" not in json.dumps(result)
