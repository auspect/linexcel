"""Real task infrastructure and HTTP consent, with an offline fake provider."""

import hashlib
import importlib.util
import json
import threading
import time
from dataclasses import replace

import pytest
from test_ask import FakeProvider, call, select_handle
from test_web import request, wait, workbook

from linexcel.ask import encoded
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
def app(tmp_path, monkeypatch):
    # These tests inject a fake provider, so they stay in the required CI, which
    # deliberately runs without the optional OpenAI SDK.
    find_spec = importlib.util.find_spec
    monkeypatch.setattr(
        importlib.util,
        "find_spec",
        lambda name, *args: object() if name == "openai" else find_spec(name, *args),
    )
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


def test_actual_foreign_owner_handle_cannot_enter_question_scope(app):
    imported(app)
    foreign = imported(app, user="bob")
    from linexcel.catalog import cell_handle, revision
    from linexcel.lazy_store import WorkbookIndex

    bob = hashlib.sha256(b"bob").hexdigest()
    with WorkbookIndex(
        app.store.project_path(bob, foreign["id"]) / "index.sqlite"
    ) as index:
        foreign_handle = cell_handle(
            revision(foreign["id"], index.meta["sha256"]), "Sheet!A1"
        )
    app.store.ask.provider_factory = lambda _: FakeProvider(
        [[call("submit_answer", {"status": "proposed", "handles": [foreign_handle]})]]
    )
    run = finished(app, start(app, prepare(app)))["run"]
    assert run["status"] == "not_found" and run["answer"] is None
    assert foreign["id"] not in encoded(run["trace"]["scope"])
    app.store.ask.provider_factory = lambda _: FakeProvider(
        [
            [
                call(
                    "read_range",
                    {"project": foreign["id"], "sheet": "Sheet", "range": "A1"},
                )
            ]
        ]
    )
    run = finished(app, start(app, prepare(app)))["run"]
    assert run["status"] == "failed" and run["answer"] is None


def test_queued_cancellation_keeps_terminal_trace(app):
    imported(app)
    gate = threading.Event()
    entered = threading.Barrier(3)

    def blocked(_):
        entered.wait(timeout=5)
        assert gate.wait(5)
        return [call("submit_answer", {"status": "not_found"})]

    app.store.ask.provider_factory = lambda _: FakeProvider([blocked])
    first = start(app, prepare(app))
    second = start(app, prepare(app))
    try:
        entered.wait(timeout=5)
        third = start(app, prepare(app))
        request(app, "POST", f"/api/tasks/{third['id']}/cancel", {})
    finally:
        gate.set()
    finished(app, first)
    finished(app, second)
    result = finished(app, third)
    assert result["task"]["status"] == "cancelled"
    assert result["run"]["trace"]["state"] == "terminal"
    assert result["run"]["trace"]["steps"][-1]["tool"] == "finish"


@pytest.mark.parametrize("change_at", ["queued", "construction"])
def test_question_uses_approved_snapshot_and_rechecks_before_request(
    app, monkeypatch, change_at
):
    imported(app)
    app.store.config = replace(
        app.store.config,
        ai_base_url="https://approved.example/v1",
        ai_model="approved-model",
    )
    jobs, constructed = [], []
    provider = FakeProvider([[call("submit_answer", {"status": "not_found"})]])

    def factory(config):
        constructed.append((config.ai_base_url, config.ai_model))
        if change_at == "construction":
            app.store.config = replace(
                app.store.config,
                ai_base_url="https://unapproved.example/v1",
                ai_model="unapproved-model",
            )
        return provider

    app.store.ask.provider_factory = factory
    monkeypatch.setattr(
        app.store.executor,
        "submit",
        lambda function, *args: jobs.append((function, args)),
    )
    task = start(app, prepare(app), remoteConsent=True)
    if change_at == "queued":
        app.store.config = replace(
            app.store.config,
            ai_base_url="https://unapproved.example/v1",
            ai_model="unapproved-model",
        )
    function, arguments = jobs[0]
    function(*arguments)
    result = finished(app, task)["run"]
    assert constructed == [("https://approved.example/v1", "approved-model")]
    assert provider.messages == []
    assert result["status"] == "failed" and result["answer"] is None
    assert result["trace"]["reason"] == "approved_configuration_changed"
    root = app.store.ask.directory(hashlib.sha256(b"alice").hexdigest(), task["id"])
    approved = _read(root / "input.json")
    assert (approved["endpoint"], approved["model"]) == constructed[0]


@pytest.mark.parametrize("mutation", ["pin", "catalog", "unrelated_pin"])
def test_completed_answer_is_invalidated_when_pin_or_catalog_changes(app, mutation):
    import base64

    from test_external import _link, linked_workbook

    from linexcel.catalog import Catalog
    from linexcel.catalog_links import set_pin
    from linexcel.lazy_store import WorkbookIndex

    def add(name, cells, links=()):
        code, raw, _ = request(
            app,
            "POST",
            "/api/projects",
            {
                "workbook": {
                    "name": name,
                    "data": base64.b64encode(
                        linked_workbook(cells, links, "S")
                    ).decode(),
                }
            },
        )
        assert code == 200, raw
        project = json.loads(raw)["project"]
        owner = hashlib.sha256(b"alice").hexdigest()
        assert (
            wait(
                app.store,
                app.store.submit(owner, project["id"], "import", None, {})["id"],
            )["status"]
            == "succeeded"
        )
        return project

    target = add("target.xlsx", {"A1": 7})
    reader = add("reader.xlsx", {"B1": "='[1]S'!A1"}, [_link("target.xlsx", "S")])
    add("unrelated.xlsx", {"C1": "='[1]S'!A1"}, [_link("target.xlsx", "S")])
    app.store.ask.provider_factory = lambda _: FakeProvider(
        [
            [
                call(
                    "read_range",
                    {"project": reader["id"], "sheet": "S", "range": "B1"},
                    "read",
                )
            ],
            lambda messages: [
                call(
                    "submit_answer",
                    {"status": "proposed", "handles": [select_handle(messages, "B1")]},
                    "end",
                )
            ],
        ]
    )
    task = start(app, prepare(app))
    result = finished(app, task)["run"]
    assert (
        result["status"] == "answer_proposed"
        and result["answer"]["lineages"][0]["coverage"]["complete"]
    )
    owner = hashlib.sha256(b"alice").hexdigest()
    with app.store.library.mutation(owner), app.store.lock:
        catalog = app.store.library.catalog(owner)
        sources = app.store.library.sources(owner)
        catalog.ensure(sources)
        from linexcel.catalog_links import list_links

        links = list_links(catalog, sources)["links"]
        link = next(
            link
            for link in links
            if (link["sourceRevision"]["projectId"] == reader["id"])
            == (mutation != "unrelated_pin")
        )
        with WorkbookIndex(
            app.store.project_path(owner, target["id"]) / "index.sqlite"
        ) as index:
            sha = index.meta["sha256"]
        if mutation in {"pin", "unrelated_pin"}:
            set_pin(catalog, sources, link["id"], target["id"], sha)
        else:
            directory = app.store.ask.directory(owner, task["id"]) / "catalog"
            Catalog(directory, pin_directory=app.store.root / owner).ensure(
                sources, force=True
            )
    changed = finished(app, task)["run"]
    if mutation != "pin":
        assert changed["status"] == "answer_proposed"
        assert changed["answer"]["lineages"][0]["coverage"]["complete"]
        return
    assert changed["status"] == "failed" and changed["answer"] is None
    assert changed["trace"]["reason"] == "lineage_state_changed"
    assert not changed["trace"]["steps"][-1]["summary"]["searchCompleted"]


def test_question_budget_is_not_capped_by_the_documentation_budget(tmp_path):
    # 2026-11-10 trial: a 16k documentation budget stopped real questions after
    # three or four tool steps with "token_budget", since each step resends the
    # history. Questions now use their own cumulative budget.
    from linexcel.ask import AskLimits
    from linexcel.cli import _build_parser

    args = _build_parser().parse_args(["serve"])
    assert args.ai_question_token_budget == AskLimits().tokens == 64000
    instance = create_app(
        tmp_path / "data",
        resolve_user=lambda env: env.get("user"),
        ai_enabled=True,
        ai_token_budget=50,
        ask_limits=AskLimits(tokens=64000),
    )
    try:
        seen = []
        original = instance.store.ask.run

        def spy(task, event):
            import linexcel.web_ask as web_ask

            real = web_ask.QuestionRun

            class Recording(real):
                def __init__(self, *args, **kwargs):
                    super().__init__(*args, **kwargs)
                    seen.append(self.limits.tokens)

            web_ask.QuestionRun = Recording
            try:
                return original(task, event)
            finally:
                web_ask.QuestionRun = real

        instance.store.ask.run = spy
        find_spec = importlib.util.find_spec
        importlib.util.find_spec = lambda name, *a: (
            object() if name == "openai" else find_spec(name, *a)
        )
        try:
            project = imported(instance)
            instance.store.ask.provider_factory = lambda _: proposal(project)
            result = finished(
                instance, start(instance, prepare(instance, project=project["id"]))
            )
        finally:
            importlib.util.find_spec = find_spec
        assert seen == [64000]
        assert result["run"]["status"] == "answer_proposed"
    finally:
        instance.close()
