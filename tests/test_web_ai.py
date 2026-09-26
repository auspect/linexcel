"""Opt-in web AI contracts, with fake providers only (no network)."""

from __future__ import annotations

import base64
import json
from dataclasses import replace

import pytest
from test_web import request, wait, workbook

from linexcel import aidoc
from linexcel.web import APIError, TaskStore, WebConfig
from linexcel.web_ai import WebAIError, capture_evidence, evidence, generate

PNG = b"\x89PNG\r\n\x1a\n" + b"selected-image"
CONFIG = {
    "baseUrl": "http://private.test/v1",
    "model": "test-model",
    "visionModel": "test-vision",
    "apiKey": "secret-test-key",
}


def graph():
    return {
        "nodes": [
            {
                "id": "S!A1",
                "kind": "input",
                "sheet": "S",
                "cell": "A1",
                "value": 6,
                "cachedValue": 6,
                "valueSource": "saved_input",
                "dependencies": [],
            },
            {
                "id": "S!B1",
                "kind": "cell",
                "sheet": "S",
                "cell": "B1",
                "formula": "=A1*7",
                "cachedValue": 40,
                "valueSource": "saved_cache",
                "dependencies": ["S!A1"],
                "patternId": "pattern1",
                "patternMemberCount": 2,
            },
        ],
        "meta": {"sha256": "abc", "sheets": ["S"], "limits": ["No global evaluation"]},
        "formulaPatterns": [
            {
                "id": "pattern1",
                "memberCount": 2,
                "representativeId": "S!B1",
                "representativeFormula": "=A1*7",
            }
        ],
    }


class FakeProvider:
    def __init__(self, text="Source : `=A1*7`.", usage=None):
        self.text = text
        self.usage = usage or aidoc.TokenUsage(100, 50, 1)
        self.calls = []

    def generate_with_usage(self, system, user, **kwargs):
        self.calls.append((system, user, kwargs))
        return self.text, self.usage

    def generate_with_image(self, system, user, image, **kwargs):
        self.calls.append((system, user, image, kwargs))
        return self.text, self.usage


def test_node_dossier_distinguishes_cache_and_latest_calculation():
    recent = {
        "id": "calc1",
        "finishedAt": 123,
        "result": {
            "status": "completed",
            "value": 42,
            "valueKind": "scalar",
            "coverage": {"complete": True},
            "comparison": "different",
            "provenance": "targeted_recalculation",
            "steps": [],
        },
    }
    snapshot = evidence(graph(), "S!B1", [recent])
    assert snapshot["savedCache"]["value"] == 40
    assert snapshot["latestCalculation"]["value"] == 42
    assert snapshot["latestCalculation"]["taskId"] == "calc1"
    assert snapshot["latestCalculation"]["valueKind"] == "scalar"
    assert snapshot["latestCalculation"]["coverage"] == {"complete": True}
    assert snapshot["node"]["patternMemberCount"] == 2
    assert snapshot["formulaGroup"]["memberCount"] == 2
    assert len(json.dumps(snapshot)) < 12000


def test_older_graph_types_remain_unknown_and_overview_preserves_error_kinds():
    sample = graph()
    sample["nodes"][1]["cachedValue"] = "#N/A"
    assert evidence(sample, "S!B1", [])["savedCache"]["kind"] == "unknown"
    recent = {
        "id": "calc-error",
        "nodeId": "S!B1",
        "result": {
            "status": "completed",
            "value": "#N/A",
            "valueKind": "error",
            "coverage": {"complete": True},
        },
    }
    result = evidence(sample, None, [recent])["recentCalculations"][0]["result"]
    assert result["valueKind"] == "error"
    assert result["coverage"] == {"complete": True}


def test_overview_contains_formulas_beyond_headers_and_explicit_omissions():
    sample = graph()
    sample["nodes"] = [
        dict(sample["nodes"][0], id=f"S!A{i}") for i in range(100)
    ] + sample["nodes"][1:]
    snapshot = evidence(sample, None, [])
    assert any(n.get("formula") == "=A1*7" for n in snapshot["sample"])
    assert snapshot["omittedNodes"] == len(sample["nodes"]) - len(snapshot["sample"])


@pytest.mark.parametrize("target", ["S!B1", None])
def test_exact_sent_formula_is_supported_and_invented_formula_warns(
    monkeypatch, target
):
    snapshot = evidence(graph(), target, [])
    provider = FakeProvider()
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    result = generate(snapshot, CONFIG, {"language": "fr"}, 16000)
    assert result["validation"]["counts"]["source_supported"] == 1
    assert result["validation"]["status"] == "unverified"
    provider.text = "Source : `=Z99+1`."
    result = generate(snapshot, CONFIG, {"language": "fr"}, 16000)
    assert result["validation"]["counts"]["not_source_supported"] == 1
    assert any("citations" in warning for warning in result["warnings"])


def test_only_selected_image_is_sent_and_usage_marks_estimate(monkeypatch):
    task = {
        "id": "capture1",
        "result": {
            "screenshots": [
                {
                    "name": "one.png",
                    "sheet": "S",
                    "data": "data:image/png;base64," + base64.b64encode(PNG).decode(),
                },
                {
                    "name": "two.png",
                    "sheet": "T",
                    "data": "data:image/png;base64,"
                    + base64.b64encode(PNG + b"second").decode(),
                },
            ]
        },
    }
    snapshot = capture_evidence(task, 1)
    provider = FakeProvider(
        "Image visible.", aidoc.TokenUsage(100, 50, 1, estimated=True)
    )
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    result = generate(snapshot, CONFIG, {"language": "fr"}, 16000)
    assert provider.calls[0][2] == PNG + b"second"
    assert "data:image" not in provider.calls[0][1]
    assert "one.png" not in provider.calls[0][1]
    assert result["sources"]["captureIndex"] == 1
    assert result["model"] == "test-vision"
    assert any("tokens image" in w for w in result["warnings"])


@pytest.mark.parametrize(
    "error,kind",
    [
        (aidoc.AiDocError("incomplete finish_reason=length"), "ai_truncated"),
        (aidoc.AiDocError("empty response"), "ai_empty"),
        (ImportError("openai"), "ai_dependency"),
        (RuntimeError("http://private.test/v1 secret-test-key"), "ai_provider"),
    ],
)
def test_provider_failures_are_real_and_do_not_expose_configuration(
    monkeypatch, error, kind
):
    def broken(**kwargs):
        raise error

    monkeypatch.setattr(aidoc, "_resolve_provider", broken)
    with pytest.raises(WebAIError) as caught:
        generate(evidence(graph(), "S!B1", []), CONFIG, {"language": "fr"}, 16000)
    assert caught.value.kind == kind
    assert "secret-test-key" not in str(caught.value)
    assert "private.test" not in str(caught.value)


def test_token_budget_checks_input_before_call_and_actual_usage_after(monkeypatch):
    provider = FakeProvider(usage=aidoc.TokenUsage(16000, 100, 1))
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    snapshot = evidence(graph(), "S!B1", [])
    with pytest.raises(WebAIError, match="insuffisant"):
        generate(snapshot, CONFIG, {"language": "fr"}, 100)
    assert provider.calls == []
    with pytest.raises(WebAIError, match="dépasse"):
        generate(snapshot, CONFIG, {"language": "fr"}, 16000)
    assert provider.calls[0][-1]["max_tokens"] < 16000


@pytest.fixture
def store(tmp_path):
    result = TaskStore(
        WebConfig(
            tmp_path,
            ai_enabled=True,
            ai_base_url=CONFIG["baseUrl"],
            ai_api_key=CONFIG["apiKey"],
        )
    )
    yield result
    result.close()


def test_ai_cache_tracks_language_config_and_exact_recent_calculation(
    store, monkeypatch
):
    created = store.create("alice", {"workbook": workbook()})
    wait(store, created["task"]["id"])
    project = created["project"]["id"]
    original = store._worker
    calls = []

    def fake(task, root, event):
        if task["operation"] == "document":
            calls.append(json.loads((root / "ai_input.json").read_text("utf-8")))
            return {
                "result": {
                    "markdown": "Fake",
                    "sources": {},
                    "language": task["options"]["language"],
                }
            }
        return original(task, root, event)

    monkeypatch.setattr(store, "_worker", fake)
    first = store.submit("alice", project, "document", "Sheet!B1", {})
    assert wait(store, first["id"])["status"] == "succeeded"
    assert calls[0]["latestCalculation"] is None
    assert (
        store.submit("alice", project, "document", "Sheet!B1", {})["id"] == first["id"]
    )
    calc = store.submit("alice", project, "evaluate", "Sheet!B1", {})
    wait(store, calc["id"])
    second = store.submit("alice", project, "document", "Sheet!B1", {})
    wait(store, second["id"])
    assert first["id"] != second["id"]
    assert calls[-1]["latestCalculation"]["value"] == 42
    english = store.submit(
        "alice", project, "document", "Sheet!B1", {}, options={"language": "en"}
    )
    wait(store, english["id"])
    assert english["id"] != second["id"]
    store.config = replace(store.config, ai_model="different-model")
    changed = store.submit("alice", project, "document", "Sheet!B1", {})
    wait(store, changed["id"])
    assert changed["id"] != second["id"]
    # Upgrading calculation semantics must not ground a fresh explanation in
    # an obsolete engine observation, even when source bytes are unchanged.
    store.cache_version = "updated-calculation-implementation"
    upgraded = store.submit("alice", project, "document", "Sheet!B1", {})
    wait(store, upgraded["id"])
    assert upgraded["id"] != changed["id"]
    assert calls[-1]["latestCalculation"] is None
    fresh = store.submit("alice", project, "evaluate", "Sheet!B1", {})
    wait(store, fresh["id"])
    refreshed = store.submit("alice", project, "document", "Sheet!B1", {})
    wait(store, refreshed["id"])
    assert calls[-1]["latestCalculation"]["taskId"] == fresh["id"]
    public = json.dumps(store.snapshot("alice", project))
    assert CONFIG["apiKey"] not in public and CONFIG["baseUrl"] not in public
    for path in store.root.rglob("*.json"):
        assert CONFIG["apiKey"] not in path.read_text("utf-8")
        assert CONFIG["baseUrl"] not in path.read_text("utf-8")


@pytest.mark.parametrize(
    "options",
    [
        {"language": []},
        {"language": {}},
        {"language": "xx"},
        {"baseUrl": "http://evil"},
        {"model": "evil"},
        {"apiKey": "evil"},
    ],
)
def test_options_do_not_configure_provider_or_accept_invalid_languages(store, options):
    created = store.create("alice", {"workbook": workbook()})
    wait(store, created["task"]["id"])
    with pytest.raises(APIError) as error:
        store.submit(
            "alice", created["project"]["id"], "document", None, {}, options=options
        )
    assert error.value.status == 400


def test_capture_description_checks_owner_project_revision_status_and_index(
    store, monkeypatch
):
    created = store.create("alice", {"workbook": workbook()})
    wait(store, created["task"]["id"])
    project = created["project"]["id"]
    original = store._worker
    seen = []
    shot = {
        "sheet": "S",
        "name": "one.png",
        "data": "data:image/png;base64," + base64.b64encode(PNG).decode(),
    }

    def fake(task, root, event):
        if task["operation"] == "capture":
            return {"result": {"screenshots": [shot]}}
        if task["operation"] == "describe_capture":
            seen.append(json.loads((root / "ai_input.json").read_text("utf-8")))
            return {"result": {"markdown": "Fake image"}}
        return original(task, root, event)

    monkeypatch.setattr(store, "_worker", fake)
    capture = store.submit("alice", project, "capture", None, {})
    wait(store, capture["id"])
    options = {"captureTaskId": capture["id"], "captureIndex": 0}
    accepted = store.submit(
        "alice", project, "describe_capture", None, {}, options=options
    )
    wait(store, accepted["id"])
    assert seen[0]["image"] == shot["data"]
    for wrong in (True, -1, 1):
        with pytest.raises(APIError):
            store.submit(
                "alice",
                project,
                "describe_capture",
                None,
                {},
                options={**options, "captureIndex": wrong},
            )
    task = store.tasks[capture["id"]]
    for field, wrong in (
        ("owner", "bob"),
        ("projectId", "different"),
        ("revision", "stale"),
        ("status", "failed"),
        ("operation", "evaluate"),
    ):
        saved = task[field]
        task[field] = wrong
        try:
            with pytest.raises(APIError):
                store.submit(
                    "alice", project, "describe_capture", None, {}, options=options
                )
        finally:
            task[field] = saved


def test_config_is_nonsecret_and_ai_default_is_optional(tmp_path):
    from linexcel.web import create_app

    app = create_app(tmp_path, resolve_user=lambda env: "alice")
    try:
        code, raw, headers = request(app, "GET", "/api/config")
        config = json.loads(raw)
        assert code == 200 and not config["ai"]["configured"]
        assert "baseUrl" not in raw.decode() and "apiKey" not in raw.decode()
        _, _, headers = request(app, "GET", "/")
        assert "img-src 'self' data: blob:" in headers["Content-Security-Policy"]
        with pytest.raises(APIError, match="Aucun fournisseur"):
            app.store.submit("alice", "unused", "document", None, {})
    finally:
        app.close()
