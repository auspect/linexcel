"""Real task infrastructure and HTTP consent, with an offline fake provider."""

import hashlib
import json
import threading
import time
from dataclasses import replace

import pytest
from test_ask import FakeProvider, call, select_handle
from test_web import request, wait, workbook

from linexcel.web import APIError, _read, _write, create_app


def imported(app, user="alice"):
    code, raw, _ = request(
        app, "POST", "/api/projects", {"workbook": workbook()}, user=user
    )
    assert code == 200, raw
    project = json.loads(raw)["project"]
    owner = hashlib.sha256(user.encode()).hexdigest()
    task = app.store.submit(owner, project["id"], "import", None, {})
    assert wait(app.store, task["id"])["status"] == "succeeded"
    return project


def prepare(app, **body):
    code, raw, _ = request(
        app,
        "POST",
        "/api/library/ask/prepare",
        {"question": "What is the saved input?", **body},
    )
    assert code == 200, raw
    return json.loads(raw)


def start(app, prepared, **body):
    code, raw, _ = request(
        app, "POST", "/api/library/ask", {"token": prepared["token"], **body}
    )
    assert code == 200, raw
    return json.loads(raw)["task"]


def finished(app, task):
    wait(app.store, task["id"])
    code, raw, _ = request(app, "GET", f"/api/library/ask/{task['id']}")
    assert code == 200, raw
    return json.loads(raw)


def proposal(project, rationale="Selected saved source."):
    return FakeProvider(
        [
            [
                call(
                    "read_range",
                    {"project": project["id"], "sheet": "Sheet", "range": "A1:B1"},
                    "read",
                )
            ],
            lambda messages: [
                call(
                    "submit_answer",
                    {
                        "status": "proposed",
                        "handles": [select_handle(messages, "A1")],
                        "rationale": rationale,
                    },
                    "end",
                )
            ],
        ]
    )


@pytest.fixture
def app(tmp_path):
    instance = create_app(
        tmp_path / "data", resolve_user=lambda env: env.get("user"), ai_enabled=True
    )
    yield instance
    instance.close()


def test_question_task_rechecks_evidence_records_feedback_and_recovers(app):
    project = imported(app)
    app.store.ask.provider_factory = lambda _: proposal(project)
    prepared = prepare(app, project=project["id"])
    assert prepared["local"] and prepared["categories"]
    task = start(app, prepared)
    result = finished(app, task)
    assert result["task"]["status"] == "succeeded"
    assert result["run"]["status"] == "answer_proposed"
    assert result["run"]["answer"]["cells"][0]["value"] == 6
    for decision in ["confirm", "reject"]:
        code, raw, _ = request(
            app,
            "POST",
            f"/api/library/ask/{task['id']}/feedback",
            {"decision": decision},
        )
        assert code == 200, raw
    assert [f["decision"] for f in json.loads(raw)["feedback"]] == ["confirm", "reject"]
    code, _, _ = request(app, "GET", f"/api/library/ask/{task['id']}", user="bob")
    assert code == 404
    root = app.store.ask.directory(hashlib.sha256(b"alice").hexdigest(), task["id"])
    assert _read(root / "trace.json")["schemaVersion"] == "linexcel.ask.v1"
    assert "Selected saved source" not in (root / "trace.json").read_text("utf-8")
    app.store.ask._recover()
    assert app.store.ask.get(hashlib.sha256(b"alice").hexdigest(), task["id"])["run"][
        "answer"
    ]
    app.store.delete(hashlib.sha256(b"alice").hexdigest(), project["id"])
    assert finished(app, task)["run"]["answer"] is None


def test_remote_consent_is_per_question_bound_to_owner_and_single_use(app):
    project = imported(app)
    app.store.config = replace(
        app.store.config, ai_base_url="https://remote.example.test/v1"
    )
    app.store.ask.provider_factory = lambda _: proposal(project)
    prepared = prepare(app)
    assert (
        not prepared["local"]
        and prepared["endpoint"] == "https://remote.example.test/v1"
    )
    code, _, _ = request(app, "POST", "/api/library/ask", {"token": prepared["token"]})
    assert code == 409
    code, _, _ = request(
        app,
        "POST",
        "/api/library/ask",
        {"token": prepared["token"], "remoteConsent": True},
        user="bob",
    )
    assert code == 409
    task = start(app, prepared, remoteConsent=True)
    assert finished(app, task)["run"]["status"] == "answer_proposed"
    code, _, _ = request(
        app,
        "POST",
        "/api/library/ask",
        {"token": prepared["token"], "remoteConsent": True},
    )
    assert code == 409
    new = prepare(app)
    code, _, _ = request(app, "POST", "/api/library/ask", {"token": new["token"]})
    assert code == 409


def test_no_model_search_remains_useful_and_ask_bypass_rejected(app):
    project = imported(app)
    app.store.config = replace(app.store.config, ai_enabled=False)
    code, raw, _ = request(app, "GET", "/api/library/ask/config")
    assert code == 200 and not json.loads(raw)["available"]
    code, _, _ = request(
        app, "POST", "/api/library/ask/prepare", {"question": "amount"}
    )
    assert code == 409
    code, _, _ = request(app, "GET", "/api/library/search", QUERY_STRING="q=6")
    assert code == 200
    with pytest.raises(APIError, match="préparation"):
        app.store.submit(
            hashlib.sha256(b"alice").hexdigest(), project["id"], "ask", None, {}
        )


def test_model_call_never_holds_global_task_lock_and_cancels(app):
    imported(app)
    entered, release = threading.Event(), threading.Event()

    def blocked(_):
        entered.set()
        assert release.wait(5)
        return [call("submit_answer", {"status": "not_found"})]

    app.store.ask.provider_factory = lambda _: FakeProvider([blocked])
    task = start(app, prepare(app))
    assert entered.wait(5)
    with app.store.lock:
        assert app.store.tasks[task["id"]]["status"] == "running"
    code, _, _ = request(app, "GET", "/api/projects")
    assert code == 200
    request(app, "POST", f"/api/tasks/{task['id']}/cancel", {})
    release.set()
    result = finished(app, task)
    assert result["task"]["status"] == "cancelled"
    assert result["run"]["trace"]["reason"] == "cancelled"


def test_expiry_and_count_retention_are_owner_local(app):
    imported(app)
    app.store.ask.provider_factory = lambda _: FakeProvider(
        [[call("submit_answer", {"status": "not_found"})]]
    )
    task = start(app, prepare(app))
    finished(app, task)
    owner = hashlib.sha256(b"alice").hexdigest()
    root = app.store.ask.directory(owner, task["id"])
    stored = _read(root / "task.json")
    stored["createdAt"] = time.time() - 8 * 86400
    _write(root / "task.json", stored)
    app.store.ask.cleanup(owner)
    assert not root.exists()
    assert task["id"] not in app.store.tasks
