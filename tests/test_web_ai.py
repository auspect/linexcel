"""Opt-in web AI contracts, with fake providers only (no network)."""

from __future__ import annotations

import base64
import json
from dataclasses import replace

import pytest
from test_web import request, wait, workbook
from vision_helpers import image_bytes, region_response

from linexcel import aidoc
from linexcel.web import APIError, TaskStore, WebConfig
from linexcel.web_ai import WebAIError, capture_evidence, evidence, generate

PNG = image_bytes()
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

    def generate(self, system, user, **kwargs):
        return self.generate_with_usage(system, user, **kwargs)[0]

    def generate_with_image(self, system, user, image, **kwargs):
        self.calls.append((system, user, image, kwargs))
        return region_response(user, self.text), self.usage


def test_omitted_steps_count_includes_sample_dropped_by_character_budget():
    recent = {
        "id": "large",
        "finishedAt": 123,
        "result": {
            "stepTotal": 302,
            "steps": [{"formula": "=" + "1+" * 400 + "1"} for _ in range(12)],
        },
    }
    dossier = evidence(graph(), "S!B1", [recent])
    calculation = dossier["latestCalculation"]
    assert calculation["dependencySteps"]["omitted"]
    assert calculation["omittedSteps"] == 302


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
    assert any(n.get("formula") == "=A1*7" for n in snapshot["formula_patterns"])
    assert snapshot["omittedNodes"] == len(sample["nodes"]) - (
        len(snapshot["formula_patterns"]) + len(snapshot["value_samples"])
    )


def test_overview_formula_tokens_do_not_borrow_a_sample_formula(monkeypatch):
    raw = "{{SOURCE_FORMULA}}\n{{STORED_XML_FORMULA}}\n{{CANONICAL_FORMULA}}"
    provider = FakeProvider(raw)
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    result = generate(evidence(graph(), None, []), CONFIG, {"language": "en"}, 16000)
    assert "{{" not in result["markdown"] and "=A1*7" not in result["markdown"]
    assert result["markdown"].count("not_supplied_in_sent_dossier") == 3
    assert result["rawResponse"] == raw
    assert all(
        r["status"] == "not_supplied"
        for r in result["validation"]["deterministic_formula_insertions"]
    )


@pytest.mark.parametrize("target", ["S!B1", None])
def test_exact_sent_formula_is_supported_and_invented_formula_warns(
    monkeypatch, target
):
    snapshot = evidence(graph(), target, [])
    provider = FakeProvider()
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    result = generate(snapshot, CONFIG, {"language": "fr"}, 16000)
    assert result["validation"]["counts"]["source_supported"] == 1
    # A correct formula quote cannot satisfy the separately required ledger.
    assert result["validation"]["status"] == "qualified"
    assert {"kind": "missing_required_claim_block"} in result["validation"]["issues"]
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


def test_live_capture_uses_every_native_region_and_exposes_bounded_evidence(
    monkeypatch,
):
    import hashlib

    from linexcel.vision_evidence import prepare_image

    payload = image_bytes(2300, 40)
    snapshot = {
        "scope": "capture",
        "sheet": "Wide table",
        "image": "data:image/png;base64," + base64.b64encode(payload).decode(),
    }
    provider = FakeProvider("Visible text")
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    result = generate(snapshot, CONFIG, {"language": "en"}, 16000)
    native = prepare_image(payload)
    assert len(provider.calls) == len(native.regions) == 3
    assert [hashlib.sha256(call[2]).hexdigest() for call in provider.calls] == [
        region.sha256 for region in native.regions
    ]
    proof = result["sources"]["visionEvidence"]
    assert proof["status"] == "complete"
    assert proof["pages"][0]["region_coverage_complete"]
    assert proof["pages"][0]["evidence"]["original_sha256"] == native.sha256
    assert not result["validation"]["factual_certification"]
    assert "r0-c2" in result["markdown"]


def test_invalid_later_native_region_cannot_return_partial_live_description(
    monkeypatch,
):
    class InvalidSecondRegion(FakeProvider):
        def generate_with_image(self, system, user, image, **kwargs):
            if self.calls:
                self.calls.append((system, user, image, kwargs))
                return '{"regions":[]}', self.usage
            return super().generate_with_image(system, user, image, **kwargs)

    payload = image_bytes(2300, 40)
    snapshot = {
        "scope": "capture",
        "sheet": "Wide",
        "image": "data:image/png;base64," + base64.b64encode(payload).decode(),
    }
    provider = InvalidSecondRegion()
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    with pytest.raises(WebAIError) as caught:
        generate(snapshot, CONFIG, {"language": "en"}, 16000)
    assert caught.value.usage["totalTokens"] == 450
    assert len(provider.calls) == 3


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


def test_text_document_keeps_both_review_phases_and_cumulative_usage(monkeypatch):
    class EditorialProvider(FakeProvider):
        def generate_with_usage(self, system, user, **kwargs):
            self.calls.append((system, user, kwargs))
            text = (
                "Draft: the engine gave40."
                if len(self.calls) == 1
                else "Saved file cache40; no targeted calculation was supplied."
            )
            return text, self.usage

    provider = EditorialProvider()
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    result = generate(evidence(graph(), "S!B1", []), CONFIG, {"language": "en"}, 16000)
    assert len(provider.calls) == 2
    phases = result["validation"]["specialized_review"]["stages"]
    assert [phase["phase"] for phase in phases] == ["draft", "review"]
    assert phases[0]["raw_response"] == "Draft: the engine gave40."
    assert result["rawResponse"] == phases[1]["raw_response"]
    assert result["usage"]["totalTokens"] == 300
    assert result["validation"]["specialized_review"]["factual_certification"] is False
    assert "secret-test-key" not in json.dumps(result)


def test_failed_review_preserves_draft_and_never_presents_it_as_reviewed(monkeypatch):
    class FailingReviewer(FakeProvider):
        def generate_with_usage(self, system, user, **kwargs):
            self.calls.append((system, user, kwargs))
            if len(self.calls) == 2:
                raise aidoc.AiDocError(
                    "AI returned empty response", usage=aidoc.TokenUsage(10, 2, 1)
                )
            return "Private draft about the supplied source.", self.usage

    provider = FailingReviewer()
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    with pytest.raises(WebAIError) as caught:
        generate(evidence(graph(), "S!B1", []), CONFIG, {"language": "en"}, 16000)
    assert caught.value.kind == "ai_empty"
    assert caught.value.usage["totalTokens"] == 162
    review = caught.value.review_report
    assert review["status"] == "failed"
    assert review["stages"][0]["raw_response"].startswith("Private draft")
    assert review["stages"][1]["status"] == "failed"


def test_uniform_capture_uses_no_model_receipt_and_reports_local_origin(monkeypatch):
    import io

    from PIL import Image

    image = io.BytesIO()
    Image.new("RGB", (40, 30), "white").save(image, format="PNG")
    provider = FakeProvider()
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    snapshot = {
        "scope": "capture",
        "sheet": "Blank",
        "image": "data:image/png;base64," + base64.b64encode(image.getvalue()).decode(),
    }
    result = generate(snapshot, CONFIG, {"language": "en"}, 16000)
    assert provider.calls == []
    assert result["rawResponse"] == "[]"
    assert result["usage"]["totalTokens"] == 0
    assert result["model"] is None and result["configuredModel"] == "test-vision"
    assert result["observationOrigin"] == "deterministic_pixels"
    assert "Aucun appel IA" in result["warnings"][0]


def test_failed_worker_keeps_review_proof_out_of_error_metadata(monkeypatch, tmp_path):
    import hashlib
    import sys

    from linexcel import web_ai, web_worker

    root = tmp_path / "project" / "tasks" / "task"
    root.mkdir(parents=True)
    (root.parent.parent / "project.json").write_text('{"filename":"source.xlsx"}')
    (root / "request.json").write_text(
        json.dumps(
            {"operation": "document", "options": {"language": "en"}, "tokens": 1000}
        )
    )
    (root / "ai_input.json").write_text('{"scope":"node"}')
    proof = {"status": "failed", "stages": [{"raw_response": "private draft"}]}

    def failed(*args, **kwargs):
        raise WebAIError("ai_empty", "Review failed", review_report=proof)

    monkeypatch.setenv("LINEXCEL_WEB_AI_CONFIG", '{"apiKey":"synthetic-secret"}')
    monkeypatch.setattr(web_ai, "generate", failed)
    monkeypatch.setattr(sys, "argv", ["web_worker", str(root)])
    web_worker.main()
    output = json.loads((root / "result.json").read_text())
    recorded = (root / "ai-review-failed.json").read_bytes()
    assert json.loads(recorded) == proof
    assert (
        output["error"]["reviewEvidence"]["sha256"]
        == hashlib.sha256(recorded).hexdigest()
    )
    assert "private draft" not in json.dumps(output)
    assert "synthetic-secret" not in json.dumps(output)


@pytest.mark.parametrize(
    "url", ["http://localhost:11434/v1", "http://127.0.0.2/v1", "http://[::1]/v1"]
)
def test_cli_gateway_key_is_not_forwarded_to_local_defaults(monkeypatch, url):
    from linexcel.web import _serve_ai_key

    monkeypatch.delenv("LINEXCEL_AI_API_KEY", raising=False)
    monkeypatch.setenv("OPENAI_API_KEY", "gateway-only-secret")
    assert _serve_ai_key(url) is None
    assert _serve_ai_key("https://gateway.example/v1") == "gateway-only-secret"
    monkeypatch.setenv("LINEXCEL_AI_API_KEY", "explicit-endpoint-key")
    assert _serve_ai_key(url) == "explicit-endpoint-key"


def test_web_provider_options_are_explicit_and_two_phase_usage_is_retained(monkeypatch):
    provider = FakeProvider()
    configured = []

    def factory(**kwargs):
        configured.append(kwargs)
        return provider

    monkeypatch.setattr(aidoc, "OpenAICompatProvider", factory)
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: kwargs["provider"])
    config = {
        **CONFIG,
        "maxTokens": 32768,
        "stream": True,
        "reasoningSplit": True,
        "thinking": "disabled",
    }
    result = generate(evidence(graph(), "S!B1", []), config, {"language": "en"}, 100000)
    assert len(provider.calls) == 2
    assert configured[0]["streaming"] is True
    assert configured[0]["request_kwargs"] == {
        "extra_body": {"reasoning_split": True, "thinking": {"type": "disabled"}}
    }
    assert all(call[-1]["max_tokens"] == 32768 for call in provider.calls)
    assert result["usage"]["totalTokens"] == 300
    assert CONFIG["apiKey"] not in json.dumps(result)


def test_server_response_policy_changes_invalidate_completed_ai_cache(
    store, monkeypatch
):
    created = store.create("alice", {"workbook": workbook()})
    wait(store, created["task"]["id"])
    project = created["project"]["id"]
    monkeypatch.setattr(
        store, "_worker", lambda *args: {"result": {"markdown": "Reviewed source."}}
    )
    first = store.submit("alice", project, "document", "Sheet!B1", {})
    wait(store, first["id"])
    assert (
        store.submit("alice", project, "document", "Sheet!B1", {})["id"] == first["id"]
    )
    store.config = replace(store.config, ai_max_tokens=32768, ai_thinking="disabled")
    second = store.submit("alice", project, "document", "Sheet!B1", {})
    wait(store, second["id"])
    assert second["id"] != first["id"]


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
