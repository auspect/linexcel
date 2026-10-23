"""The manual gate must fail for incomplete languages and preserve evidence."""

import hashlib
import importlib.util
import json
import sys
from argparse import Namespace
from contextlib import nullcontext
from pathlib import Path
from types import SimpleNamespace

import pytest


@pytest.fixture
def manual(monkeypatch):
    root = Path(__file__).resolve().parents[1]
    monkeypatch.syspath_prepend(str(root))
    spec = importlib.util.spec_from_file_location(
        "manual_validation_test", root / "validate_manual.py"
    )
    module = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, spec.name, module)
    spec.loader.exec_module(module)
    return module


def test_sample_is_reproducible_and_spans_sheets(manual):
    nodes = [
        {"id": f"{sheet}!A{i}", "sheet": sheet, "kind": "cell"}
        for sheet in ("A", "B", "C")
        for i in range(10)
    ]
    selected = manual.selected_nodes(nodes, 6, 123)
    assert selected == manual.selected_nodes(list(reversed(nodes)), 6, 123)
    assert {nid.split("!")[0] for nid in selected} == {"A", "B", "C"}


def test_sample_includes_sheetless_vba_alongside_formula_nodes(manual):
    nodes = [
        {"id": "S!A1", "sheet": "S", "kind": "cell"},
        {"id": "vp:M.P", "sheet": None, "kind": "vba"},
        {"id": "vp:M.Q", "kind": "vba"},
    ]
    selected = manual.selected_nodes(nodes, None, 123)
    assert set(selected) == {node["id"] for node in nodes}
    assert selected == manual.selected_nodes(list(reversed(nodes)), None, 123)


def test_failure_trace_never_copies_provider_messages(manual):
    cause = ConnectionError("synthetic-secret in provider body")
    error = manual.AiDocError("sanitized connection failure")
    error.__cause__ = cause
    evidence = manual.ai_failure_evidence(error, stage="overview", language="en")
    assert evidence["cause_chain"] == [
        {"type": "AiDocError"},
        {"type": "ConnectionError"},
    ]
    assert "synthetic-secret" not in json.dumps(evidence)
    assert evidence["failed_transport_attempt_usage"] == "not reported by SDK"
    assert not evidence["response_usage_available"]
    assert not evidence["response_usage_reported"]
    assert evidence["response_usage_estimated"] is None


def test_partial_stream_usage_is_distinguished_from_reported_counts(manual):
    error = manual.AiDocError("partial stream")
    error.usage = SimpleNamespace(estimated=True)
    evidence = manual.ai_failure_evidence(error, stage="nodes", language="en")
    assert evidence["response_usage_available"]
    assert evidence["response_usage_estimated"]
    assert not evidence["response_usage_reported"]


def test_vision_transport_failure_keeps_a_sanitized_stage_trace(manual, monkeypatch):
    def describe_failure(*args, **kwargs):
        raise manual.AiDocError("vision transport failed") from ConnectionError(
            "synthetic-secret in provider message"
        )

    monkeypatch.setattr(manual, "acceptance_provider", lambda *a, **k: nullcontext())
    checks = {}
    result = SimpleNamespace(describe_screenshots=describe_failure)
    args = Namespace(
        vision_model="vision-model", model="text-model", max_tokens=42, token_budget=100
    )
    assert (
        manual.describe(
            result, {"S": "capture.png"}, args, "en", validation_results=checks
        )
        is None
    )
    failure = checks["provider_failures"][0]
    assert failure["stage"] == "vision" and failure["language"] == "en"
    assert failure["cause_chain"][-1]["type"] == "ConnectionError"
    assert "synthetic-secret" not in json.dumps(checks)


@pytest.mark.parametrize("reason,accepted", [("stop", True), ("length", False)])
def test_completion_receipts_preserve_finish_and_reported_usage(
    manual, tmp_path, reason, accepted
):
    receipts = []
    journal = tmp_path / "receipts.jsonl"
    earlier = {"response_id": "earlier-completion", "status": "complete"}
    journal.write_text(json.dumps(earlier) + "\n", encoding="utf-8")
    provider = manual.AcceptanceProvider(
        model="requested-model",
        client=SimpleNamespace(),
        response_receipts=receipts,
        response_journal=journal,
    )
    response = SimpleNamespace(
        model="resolved-model",
        id="response-1",
        choices=[SimpleNamespace(finish_reason=reason)],
        usage=SimpleNamespace(prompt_tokens=7, completion_tokens=11),
        secret_body="synthetic-secret",
    )
    if accepted:
        assert (
            provider._checked_response(
                response, "source dossier", "reviewed text", require_stop=True
            )[0]
            == "reviewed text"
        )
    else:
        with pytest.raises(manual.AiDocError):
            provider._checked_response(
                response, "source dossier", "reviewed text", require_stop=True
            )
    assert len(receipts) == 1
    receipt = receipts[0]
    assert receipt["status"] == ("complete" if accepted else "failed")
    assert receipt["requested_model"] == "requested-model"
    assert receipt["resolved_model"] == "resolved-model"
    assert receipt["finish_reason"] == reason
    assert receipt["streaming"] is True
    assert receipt["usage"]["input_tokens"] == 7
    assert receipt["usage"]["output_tokens"] == 11
    assert receipt["content_sha256"] == hashlib.sha256(b"reviewed text").hexdigest()
    assert receipt["raw_response"] == "reviewed text"
    assert "synthetic-secret" not in json.dumps(receipts)
    # The completed/rejected receipt is durable before the language finishes,
    # and does not overwrite a previous provider's completion in the same run.
    assert [json.loads(line) for line in journal.read_text("utf-8").splitlines()] == [
        earlier,
        receipt,
    ]
    assert "synthetic-secret" not in journal.read_text("utf-8")


def test_rejected_response_is_preserved_before_completion_gate(manual, monkeypatch):
    receipts = []
    provider = manual.AcceptanceProvider(
        model="requested-model", client=SimpleNamespace(), response_receipts=receipts
    )
    text = 'Partial source review\n```json_claims\n{"claims": ['
    response = SimpleNamespace(
        model="resolved-model",
        id="response-length",
        choices=[SimpleNamespace(finish_reason="length")],
    )

    def reject(self, response, prompt, received, *, require_stop=False):
        assert receipts[0]["status"] == "pending"
        assert receipts[0]["raw_response"] == received == text
        raise manual.AiDocError("completion rejected")

    monkeypatch.setattr(manual.OpenAICompatProvider, "_checked_response", reject)
    with pytest.raises(manual.AiDocError, match="completion rejected"):
        provider._checked_response(response, "source dossier", text)
    assert len(receipts) == 1
    assert receipts[0]["status"] == "failed"
    assert receipts[0]["raw_response"] == text
    assert receipts[0]["content_sha256"] == hashlib.sha256(text.encode()).hexdigest()
    assert receipts[0]["usage"] is None


def test_corpus_files_are_verified_and_do_not_replace_fixtures(manual, tmp_path):
    workbook, addin = tmp_path / "real.xlsx", tmp_path / "tool.xlam"
    workbook.write_bytes(b"PK public source")
    addin.write_bytes(b"PK public VBA")
    manifest = tmp_path / "manifest.json"
    entries = [
        {
            "path": path.name,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            "source_url": "https://official.example/" + path.name,
            "role": role,
        }
        for path, role in [(workbook, "workbook"), (addin, "addin")]
    ]
    manifest.write_text(json.dumps({"version": 1, "files": entries}), encoding="utf-8")
    cases, addins = manual.corpus_files(manifest)
    assert len(cases) == len(addins) == 1
    assert cases[0].name == "public-real" and not cases[0].generated
    assert cases[0].workbook == workbook
    assert cases[0].source_receipt["sha256"] == entries[0]["sha256"]
    assert set(manual.CASES) == {"sales", "stress"}
    workbook.write_bytes(b"changed")
    with pytest.raises(ValueError, match="fingerprint mismatch"):
        manual.corpus_files(manifest)


def test_corpus_manifest_refuses_unrecorded_or_non_https_sources(manual, tmp_path):
    path = tmp_path / "real.xlsx"
    path.write_bytes(b"PK source")
    manifest = tmp_path / "manifest.json"
    entry = {
        "path": path.name,
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "source_url": "http://official.example/real.xlsx",
        "role": "workbook",
    }
    manifest.write_text(json.dumps({"version": 1, "files": [entry]}), encoding="utf-8")
    with pytest.raises(ValueError, match="HTTPS"):
        manual.corpus_files(manifest)
    entry["source_url"] = "https://official.example/real.xlsx"
    entry.pop("sha256")
    manifest.write_text(json.dumps({"version": 1, "files": [entry]}), encoding="utf-8")
    with pytest.raises(ValueError, match="SHA-256"):
        manual.corpus_files(manifest)


def test_corpus_refuses_addin_artifact_collisions(manual, tmp_path):
    entries = []
    for folder in ["one", "two"]:
        path = tmp_path / folder / "Tool.xlam"
        path.parent.mkdir()
        path.write_bytes(b"PK source")
        entries.append(
            {
                "path": str(path.relative_to(tmp_path)),
                "role": "addin",
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "source_url": "https://official.example/Tool.xlam",
            }
        )
    manifest = tmp_path / "manifest.json"
    manifest.write_text(json.dumps({"version": 1, "files": entries}), encoding="utf-8")
    with pytest.raises(ValueError, match="Duplicate corpus add-in"):
        manual.corpus_files(manifest)


@pytest.mark.parametrize(
    "missing", ["nodes", "overview", "vision", "images", "sheet", "regions", None]
)
def test_each_language_and_stage_must_complete(manual, monkeypatch, tmp_path, missing):
    image = tmp_path / "sheet.png"
    image_module = pytest.importorskip("PIL.Image")
    image_module.new("RGB", (40, 30), "white").save(image)
    result = SimpleNamespace(
        graph={"meta": {"execution": {"status": "completed"}}},
        nodes=[{"id": "S!A1", "kind": "cell", "sheet": "S"}],
        sheets=["S"],
        token_usage=SimpleNamespace(estimated=False),
        to_json=lambda **kw: "{}",
        save_html=lambda path, **kw: path.write_text("report", encoding="utf-8"),
    )
    monkeypatch.setattr(manual.linexcel, "analyze", lambda *args, **kwargs: result)
    monkeypatch.setattr(manual, "report_structure", lambda result: None)
    monkeypatch.setattr(manual, "report_context", lambda result: None)
    monkeypatch.setattr(
        manual,
        "_sheet_names",
        lambda data: ["S", "Lost"] if missing == "sheet" else ["S"],
    )
    monkeypatch.setattr(
        manual,
        "render_screenshots",
        lambda *args: None if missing == "images" else {"S": [image]},
    )

    def document(result, args, language, *, validation_results=None):
        from test_documentation_receipts import reviewed

        receipts = []
        validation_results["S!A1"] = {
            **reviewed("S!A1", receipts),
            "status": "qualified",
        }
        validation_results["workbook"] = reviewed("workbook", receipts)
        validation_results["request_receipts"] = receipts
        return (
            {} if language == "en" and missing == "nodes" else {"S!A1": "Complete"},
            " " if language == "en" and missing == "overview" else "Overview",
        )

    monkeypatch.setattr(manual, "document", document)

    def describe(result, shots, args, language, **kwargs):
        from linexcel.vision_evidence import persist_image, prepare_image

        native = prepare_image(image.read_bytes())
        metadata = persist_image(native, tmp_path / "vision_crops")
        kwargs["vision_results"]["S"] = {
            "status": "complete",
            "expected_pages": 1,
            "pages": [
                {
                    "page": 1,
                    "evidence": metadata,
                    "region_coverage_complete": True,
                    "completed_regions": [
                        region["id"] for region in metadata["regions"]
                    ],
                    "batches": [
                        {
                            "id": "batch",
                            "prompt_text_sha256": "prompt",
                            "content_sha256": "content",
                            "original_sha256": native.sha256,
                            "json_prompt_sha256": "json",
                            "images": metadata["regions"],
                            "region_ids": [
                                region["id"] for region in metadata["regions"]
                            ],
                            "status": "complete",
                        }
                    ],
                }
            ],
        }
        if missing != "regions":
            kwargs["response_receipts"].append(
                {
                    "prompt_text_sha256": "prompt",
                    "content_sha256": "content",
                    "status": "complete",
                    "finish_reason": "stop",
                    "response_id": "response",
                }
            )
        return (
            {} if language == "en" and missing == "vision" else {"S": "Visible image"}
        )

    monkeypatch.setattr(manual, "describe", describe)
    (tmp_path / "source.xlsx").write_bytes(b"workbook")
    case = manual.Case(
        "sample",
        "test",
        lambda: b"workbook",
        tmp_path / "source.xlsx",
        ("fr", "en"),
        generated=False,
        root=tmp_path,
    )
    args = Namespace(no_recalc=True, max_nodes=None, seed=1, check_max_tokens=False)
    outcome = manual.run_case(case, args, True, True)
    assert outcome["passed"] is (missing is None)
    assert (tmp_path / "sample-en-ai.json").exists()
    evidence = json.loads((tmp_path / "sample-en-ai.json").read_text())
    assert evidence["documentation_checks"]["S!A1"]["status"] == "qualified"
    assert evidence["documentation_checks"]["S!A1"]["raw_response"].startswith(
        "S!A1: review response"
    )
    assert json.loads((tmp_path / "sample-validation.json").read_text())["passed"] is (
        missing is None
    )


def test_missing_provider_is_nonzero_and_recorded(manual, monkeypatch, tmp_path):
    output = tmp_path / "run"
    monkeypatch.setattr(
        sys, "argv", ["validate_manual.py", "--output-dir", str(output)]
    )
    monkeypatch.setattr(manual, "check_local_provider", lambda *args: False)
    seen = []

    def run(case, args, ai, vision):
        seen.append((ai, vision))
        return {"case": case.name, "passed": ai and vision}

    monkeypatch.setattr(manual, "run_case", run)
    assert manual.main() == 1
    assert seen == [(False, False), (False, False)]
    assert json.loads((output / "validation.json").read_text())["status"] == "failed"
    with pytest.raises(SystemExit) as caught:
        manual.main()
    assert caught.value.code == 2


def test_no_screenshots_and_no_ai_never_start_optional_stages(
    manual, monkeypatch, tmp_path
):
    (tmp_path / "source.xlsx").write_bytes(b"synthetic")
    result = SimpleNamespace(
        graph={"meta": {"execution": {"status": "completed"}}},
        nodes=[],
        sheets=["S"],
        token_usage=SimpleNamespace(estimated=False),
        to_json=lambda **kw: "{}",
        save_html=lambda path, **kw: path.write_text("report", encoding="utf-8"),
    )
    monkeypatch.setattr(manual.linexcel, "analyze", lambda *a, **k: result)
    for name in ("report_structure", "report_context"):
        monkeypatch.setattr(manual, name, lambda *a: None)
    monkeypatch.setattr(manual, "_sheet_names", lambda *a: ["S"])

    def forbidden(*a, **k):
        pytest.fail("Explicitly skipped stage was started")

    for name in ("render_screenshots", "document", "describe", "check_local_provider"):
        monkeypatch.setattr(manual, name, forbidden)
    monkeypatch.setattr(
        manual,
        "CASES",
        {
            "sales": manual.Case(
                "sample",
                "test",
                lambda: b"synthetic",
                tmp_path / "source.xlsx",
                ("en",),
                generated=False,
                root=tmp_path,
            )
        },
    )
    output = tmp_path / "run"
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "validate_manual.py",
            "--no-ai",
            "--no-vision",
            "--no-screenshots",
            "--no-recalc",
            "--output-dir",
            str(output),
        ],
    )
    assert manual.main() == 1
    evidence = json.loads((output / "validation.json").read_text())
    assert evidence["status"] == "incomplete"
    assert evidence["cases"][0]["execution"]["status"] == "completed"
    assert evidence["cases"][0]["screenshots"] == []


def test_placeholder_and_unexpected_keys_are_not_coverage(manual):
    assert not manual.coverage(["a"], {"a": "(AI returned empty response)"})["passed"]
    assert not manual.coverage(["a"], {"a": "ok", "b": "extra"})["passed"]


def test_header_only_png_does_not_pass_image_validation(manual, tmp_path):
    pytest.importorskip("PIL.Image")
    image = tmp_path / "truncated.png"
    image.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + b"\0\0\0\rIHDR"
        + (40).to_bytes(4, "big")
        + (30).to_bytes(4, "big")
        + b"\0"
    )
    assert not manual.screenshot_evidence({"sheet": image})[0]["passed"]


def test_gateway_discovery_authenticates_without_forwarding_key_to_loopback(
    manual, monkeypatch
):
    import io

    monkeypatch.delenv("LINEXCEL_AI_API_KEY", raising=False)
    monkeypatch.setenv("OPENAI_API_KEY", "synthetic-gateway-secret")
    requests = []

    def open_request(request, *, timeout):
        requests.append(request)
        return io.BytesIO(b'{"data": [{"id": "selected-model"}]}')

    monkeypatch.setattr(
        manual.urllib.request,
        "build_opener",
        lambda *handlers: SimpleNamespace(open=open_request),
    )
    assert manual.probe_local_models("https://gateway.example/v1") == ["selected-model"]
    assert requests[-1].get_header("Authorization") == (
        "Bearer synthetic-gateway-secret"
    )
    assert manual.probe_local_models("http://localhost:11434/v1") == ["selected-model"]
    assert requests[-1].get_header("Authorization") == "Bearer not-needed"


def test_authenticated_discovery_refuses_cross_origin_redirect(manual):
    import io
    from urllib.response import addinfourl

    requests = []

    class RedirectingGateway(manual.urllib.request.HTTPHandler):
        handler_order = 0

        def https_open(self, request):
            requests.append(request)
            response = addinfourl(
                io.BytesIO(),
                {"Location": "http://localhost:11434/v1/models"},
                request.full_url,
                code=302,
            )
            response.msg = "Found"
            return response

        def http_open(self, request):
            pytest.fail("Discovery followed a redirect to another origin")

    request = manual.urllib.request.Request(
        "https://gateway.example/v1/models",
        headers={"Authorization": "Bearer synthetic-secret"},
    )
    opener = manual.urllib.request.build_opener(
        RedirectingGateway(), manual._NoModelRedirect()
    )
    with pytest.raises(manual.urllib.error.HTTPError) as caught:
        opener.open(request)
    assert caught.value.code == 302
    assert len(requests) == 1
    assert requests[0].get_header("Authorization") == "Bearer synthetic-secret"


def test_provider_request_options_apply_to_documentation_and_vision(
    manual, monkeypatch
):
    captured = []
    monkeypatch.setattr(
        manual, "AcceptanceProvider", lambda **kwargs: captured.append(kwargs)
    )
    monkeypatch.setenv("OPENAI_API_KEY", "synthetic-gateway-secret")
    monkeypatch.delenv("LINEXCEL_AI_API_KEY", raising=False)
    args = Namespace(
        base_url="https://gateway.example/v1",
        model="doc-model",
        vision_base_url="http://localhost:11434/v1",
        vision_model="vision-model",
        reasoning_split=True,
        api_retries=2,
        api_stream=True,
    )
    manual.acceptance_provider(args)
    manual.acceptance_provider(args, vision=True)
    assert captured == [
        {
            "response_receipts": None,
            "response_journal": None,
            "base_url": "https://gateway.example/v1",
            "model": "doc-model",
            "api_key": "synthetic-gateway-secret",
            "client_kwargs": {"max_retries": 2},
            "streaming": True,
            "request_kwargs": {"extra_body": {"reasoning_split": True}},
        },
        {
            "response_receipts": None,
            "response_journal": None,
            "base_url": "http://localhost:11434/v1",
            "model": "vision-model",
            "api_key": "not-needed",
            "client_kwargs": {"max_retries": 2},
            "streaming": True,
            "request_kwargs": {"extra_body": {"reasoning_split": True}},
        },
    ]


@pytest.mark.parametrize("thinking", [None, "disabled", "adaptive"])
@pytest.mark.parametrize("reasoning_split", [False, True])
@pytest.mark.parametrize("chat_template_thinking", [None, "enabled", "disabled"])
def test_explicit_thinking_merges_with_reasoning_split(
    manual, monkeypatch, thinking, reasoning_split, chat_template_thinking
):
    captured = []
    monkeypatch.setattr(
        manual, "AcceptanceProvider", lambda **kwargs: captured.append(kwargs)
    )
    args = Namespace(
        base_url="http://localhost:11434/v1",
        model="MiniMax-M3",
        vision_base_url=None,
        vision_model=None,
        reasoning_split=reasoning_split,
        api_thinking=thinking,
        chat_template_thinking=chat_template_thinking,
    )
    expected = {}
    if reasoning_split:
        expected["reasoning_split"] = True
    if thinking is not None:
        expected["thinking"] = {"type": thinking}
    if chat_template_thinking is not None:
        expected["chat_template_kwargs"] = {
            "enable_thinking": chat_template_thinking == "enabled"
        }
    manual.acceptance_provider(args)
    manual.acceptance_provider(args, vision=True)
    assert [entry["request_kwargs"] for entry in captured] == [
        {"extra_body": expected} if expected else None
    ] * 2
    assert all(entry["client_kwargs"] == {"max_retries": 0} for entry in captured)
    assert all(entry["streaming"] is False for entry in captured)


def test_env_file_is_explicit_and_credentials_are_not_in_manifest(
    manual, monkeypatch, tmp_path
):
    pytest.importorskip("dotenv")
    monkeypatch.setattr(manual.os, "environ", manual.os.environ.copy())
    for name in ("OPENAI_BASE_URL", "OPENAI_API_KEY", "LINEXCEL_AI_BASE_URL"):
        monkeypatch.delenv(name, raising=False)
    env_file = tmp_path / ".env"
    env_file.write_text(
        "OPENAI_BASE_URL=https://gateway.example/v1\nOPENAI_API_KEY=synthetic-secret\n"
    )
    seen = []
    monkeypatch.setattr(
        manual,
        "check_local_provider",
        lambda base, model: seen.append((base, model)) or True,
    )
    monkeypatch.setattr(
        manual,
        "run_case",
        lambda case, args, ai, vision: {"case": case.name, "passed": ai and vision},
    )
    output = tmp_path / "explicit-run"
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "validate_manual.py",
            "--env-file",
            str(env_file),
            "--model",
            "selected-model",
            "--vision-model",
            "vision-model",
            "--vision-base-url",
            "http://localhost:11434/v1",
            "--api-thinking",
            "disabled",
            "--chat-template-thinking",
            "disabled",
            "--output-dir",
            str(output),
        ],
    )
    assert manual.main() == 0
    assert seen == [
        ("https://gateway.example/v1", "selected-model"),
        ("http://localhost:11434/v1", "vision-model"),
    ]
    manifest_text = (output / "validation.json").read_text()
    assert "synthetic-secret" not in manifest_text
    assert "api_key" not in json.loads(manifest_text)["configuration"]
    assert json.loads(manifest_text)["configuration"]["api_thinking"] == "disabled"
    assert (
        json.loads(manifest_text)["configuration"]["chat_template_thinking"]
        == "disabled"
    )
    # Even an already populated environment must not switch a default run.
    seen.clear()
    monkeypatch.setattr(
        sys, "argv", ["validate_manual.py", "--output-dir", str(tmp_path / "local-run")]
    )
    assert manual.main() == 0
    assert seen == [(manual.DEFAULT_BASE_URL, manual.DEFAULT_MODEL)]
    local_manifest = json.loads(
        (tmp_path / "local-run" / "validation.json").read_text()
    )
    assert local_manifest["configuration"]["api_thinking"] is None
    assert local_manifest["configuration"]["chat_template_thinking"] is None


@pytest.mark.parametrize(
    "damage",
    [
        None,
        "missing_page",
        "missing_batch",
        "duplicate_batch",
        "crop_sha",
        "completed_regions",
        "original_sha",
        "crop_bytes",
        "receipt_hash",
        "duplicate_receipt",
        "length",
    ],
)
def test_native_coverage_requires_every_exact_page_region_and_receipt(
    manual, tmp_path, damage
):
    from vision_helpers import image_bytes, region_response

    from linexcel.aidoc import TokenUsage, describe_images

    class NativeProvider:
        def generate(self, *args, **kwargs):
            raise AssertionError("Pixels required")

        def generate_with_images(self, system, user, images, **kwargs):
            return region_response(user), TokenUsage(20, 10, 1)

    checks = {}
    describe_images(
        {"S": [image_bytes(4600, 30), image_bytes()]},
        provider=NativeProvider(),
        validation_results=checks,
        evidence_dir=tmp_path,
    )
    pages = checks["S"]["pages"]
    receipts = [
        {
            "prompt_text_sha256": attempt["prompt_text_sha256"],
            "content_sha256": attempt["content_sha256"],
            "raw_response": attempt.get("raw_response"),
            "usage": attempt.get("usage"),
            "status": "complete",
            "finish_reason": "stop",
            "response_id": f"{batch['id']}-attempt-{attempt.get('number', 1)}",
        }
        for page in pages
        for batch in page["batches"]
        for attempt in batch.get("attempts", [batch])
    ]
    first = pages[0]
    if damage == "missing_page":
        pages.pop()
    elif damage == "missing_batch":
        first["batches"].pop()
    elif damage == "duplicate_batch":
        first["batches"].append(first["batches"][0])
    elif damage == "crop_sha":
        first["batches"][0]["images"][0]["sha256"] = "different"
    elif damage == "completed_regions":
        first["completed_regions"].pop()
    elif damage == "original_sha":
        first["batches"][0]["original_sha256"] = "different"
    elif damage == "crop_bytes":
        Path(first["evidence"]["regions"][0]["path"]).write_bytes(b"changed")
    elif damage == "receipt_hash":
        receipts[0]["content_sha256"] = "different"
    elif damage == "duplicate_receipt":
        receipts.append(receipts[0])
    elif damage == "length":
        receipts[0]["finish_reason"] = "length"
    manual.link_vision_receipts(checks, receipts)
    coverage = manual.vision_region_coverage(["S"], checks)
    assert coverage["passed"] is (damage is None)
    assert not coverage["factual_certification"]


@pytest.mark.parametrize(
    "damage",
    [
        None,
        "missing_prior_receipt",
        "prior_length",
        "prior_raw",
        "prior_usage",
        "duplicate_response_id",
        "duplicate_prompt",
        "prior_not_contract_failure",
        "missing_contract_error",
        "missing_attempt",
        "final_attempt",
        "top_prompt",
        "top_content",
        "top_usage",
        "prompt_hash",
        "system_prompt",
        "usage_omits_prior",
        "usage_counts_twice",
    ],
)
def test_retry_coverage_checks_every_attempt_and_final_selection(
    manual, tmp_path, damage
):
    from vision_helpers import image_bytes

    from linexcel.vision_evidence import persist_image, prepare_image

    native = prepare_image(image_bytes())
    metadata = persist_image(native, tmp_path / "crops")
    ids = [region["id"] for region in metadata["regions"]]
    usage = {
        "input_tokens": 20,
        "output_tokens": 10,
        "requests": 1,
        "estimated": False,
        "model": "selected-model",
        "provider": "openai-compatible",
    }
    attempts, receipts = [], []
    for number in (1, 2):
        prompt = json.dumps({"regions": ids, "contract_retry": number})
        raw = "invalid JSON" if number == 1 else '{"observations": []}'
        digest = hashlib.sha256(raw.encode()).hexdigest()
        receipt = {
            "status": "complete",
            "finish_reason": "stop",
            "response_id": f"response-{number}",
            "prompt_text_sha256": f"prompt-{number}",
            "content_sha256": digest,
            "raw_response": raw,
            "usage": usage.copy(),
        }
        attempts.append(
            {
                "number": number,
                "status": "invalid_response" if number == 1 else "complete",
                "contract_error": "invalid JSON" if number == 1 else None,
                "user_prompt": prompt,
                "json_prompt_sha256": hashlib.sha256(prompt.encode()).hexdigest(),
                "prompt_text_sha256": f"prompt-{number}",
                "system_prompt_sha256": "same-system",
                "raw_response": raw,
                "content_sha256": digest,
                "usage": usage.copy(),
                "provider_receipt": receipt.copy(),
            }
        )
        receipts.append(receipt)
    batch = {
        "id": "page-1-batch-0",
        "original_sha256": native.sha256,
        "images": metadata["regions"],
        "region_ids": ids,
        "attempts": attempts,
        "final_attempt": 2,
        "usage_total": dict(usage, input_tokens=40, output_tokens=20, requests=2),
        **{key: value for key, value in attempts[-1].items() if key != "number"},
    }
    checks = {
        "S": {
            "status": "complete",
            "expected_pages": 1,
            "pages": [
                {
                    "page": 1,
                    "evidence": metadata,
                    "region_coverage_complete": True,
                    "completed_regions": ids,
                    "batches": [batch],
                }
            ],
        }
    }
    if damage == "missing_prior_receipt":
        receipts.pop(0)
    elif damage == "prior_length":
        receipts[0]["finish_reason"] = "length"
    elif damage == "prior_raw":
        attempts[0]["raw_response"] = "changed"
    elif damage == "prior_usage":
        attempts[0]["usage"] = dict(usage, output_tokens=11)
    elif damage == "duplicate_response_id":
        receipts[0]["response_id"] = receipts[1]["response_id"]
    elif damage == "duplicate_prompt":
        attempts[0]["prompt_text_sha256"] = attempts[1]["prompt_text_sha256"]
    elif damage == "prior_not_contract_failure":
        attempts[0]["status"] = "failed"
    elif damage == "missing_contract_error":
        attempts[0].pop("contract_error")
    elif damage == "missing_attempt":
        attempts.pop(0)
    elif damage == "final_attempt":
        batch["final_attempt"] = 1
    elif damage == "top_prompt":
        batch["prompt_text_sha256"] = "another-prompt"
    elif damage == "top_content":
        batch["raw_response"] = "earlier response"
    elif damage == "top_usage":
        batch["usage"] = batch["usage_total"]
    elif damage == "prompt_hash":
        attempts[0]["json_prompt_sha256"] = "wrong"
    elif damage == "system_prompt":
        attempts[0]["system_prompt_sha256"] = "other-system"
    elif damage == "usage_omits_prior":
        batch["usage_total"] = usage
    elif damage == "usage_counts_twice":
        batch["usage_total"]["requests"] = 3
    manual.link_vision_receipts(checks, receipts)
    coverage = manual.vision_region_coverage(["S"], checks)
    accepted = damage in {None, "duplicate_response_id"}
    assert coverage["passed"] is accepted
    assert not coverage["factual_certification"]
    if accepted:
        assert [receipt["vision_evidence"]["selected"] for receipt in receipts] == [
            False,
            True,
        ]
        assert [receipt["vision_evidence"]["attempt"] for receipt in receipts] == [1, 2]
        assert [attempt["provider_receipt_index"] for attempt in attempts] == [0, 1]
        assert [attempt["response_id"] for attempt in attempts] == [
            receipt["response_id"] for receipt in receipts
        ]


@pytest.mark.parametrize(
    "damage", [None, "reused_index", "ambiguous_receipt", "prompt", "content", "id"]
)
def test_vision_receipts_require_unique_local_evidence(manual, damage):
    batches, receipts = [], []
    for index in range(2):
        batches.append(
            {
                "id": f"batch-{index}",
                "prompt_text_sha256": f"prompt-{index}",
                "content_sha256": f"content-{index}",
                "original_sha256": "source-image",
                "json_prompt_sha256": f"json-{index}",
                "images": [],
            }
        )
        receipts.append(
            {
                "response_id": "chatcmpl-51",
                "prompt_text_sha256": f"prompt-{index}",
                "content_sha256": f"content-{index}",
                "status": "complete",
                "finish_reason": "stop",
            }
        )
    if damage == "reused_index":
        batches[1].update(prompt_text_sha256="prompt-0", content_sha256="content-0")
        receipts.pop()
    elif damage == "ambiguous_receipt":
        receipts.append(receipts[1].copy())
    elif damage == "prompt":
        receipts[1]["prompt_text_sha256"] = "different"
    elif damage == "content":
        receipts[1]["content_sha256"] = "different"
    elif damage == "id":
        receipts[1]["response_id"] = ""
    checks = {"S": {"pages": [{"page": 1, "batches": batches}]}}
    manual.link_vision_receipts(checks, receipts)
    assert batches[0]["receipt_link_complete"]
    assert batches[1]["receipt_link_complete"] is (damage is None)
    assert batches[0]["provider_receipt_index"] == 0
    if damage is None:
        assert batches[1]["provider_receipt_index"] == 1
        assert [batch["response_id"] for batch in batches] == [
            "chatcmpl-51",
            "chatcmpl-51",
        ]


def test_real_acceptance_provider_retry_receipts_pass_native_gate(manual, tmp_path):
    from vision_helpers import image_bytes, region_response

    from linexcel.aidoc import TokenUsage, describe_images

    calls = []

    def create(**kwargs):
        calls.append(kwargs)
        prompt = kwargs["messages"][-1]["content"][0]["text"]
        text = "invalid JSON" if len(calls) == 1 else region_response(prompt)
        return SimpleNamespace(
            model="MiniMax-M3",
            id=f"response-{len(calls)}",
            choices=[
                SimpleNamespace(
                    finish_reason="stop", message=SimpleNamespace(content=text)
                )
            ],
            usage=SimpleNamespace(prompt_tokens=20, completion_tokens=10),
        )

    receipts, checks = [], {}
    provider = manual.AcceptanceProvider(
        model="MiniMax-M3",
        client=SimpleNamespace(
            chat=SimpleNamespace(completions=SimpleNamespace(create=create))
        ),
        response_receipts=receipts,
    )
    usage = TokenUsage()
    described = describe_images(
        {"S": image_bytes()},
        provider=provider,
        usage=usage,
        validation_results=checks,
        evidence_dir=tmp_path,
    )
    assert set(described) == {"S"}
    assert usage.requests == len(receipts) == len(calls) == 2
    assert usage.total == 60
    manual.link_vision_receipts(checks, receipts)
    assert manual.vision_region_coverage(["S"], checks)["passed"]
    batch = checks["S"]["pages"][0]["batches"][0]
    assert batch["usage_total"]["requests"] == 2
    assert batch["usage"]["requests"] == 1
    assert batch["provider_receipt"]["raw_response"] == batch["raw_response"]
    assert batch["attempts"][0]["provider_receipt"]["raw_response"] == "invalid JSON"


@pytest.mark.parametrize(
    "damage",
    [
        None,
        "fake_receipt",
        "fake_response",
        "fake_attempt",
        "tokens",
        "missing_decision",
        "extra_observation",
        "uniform_metadata",
        "nonuniform_crop",
        "nonuniform_original",
        "missing_usage",
        "missing_bbox",
    ],
)
def test_native_uniform_shortcut_requires_exact_pixels_and_no_invented_receipt(
    manual, tmp_path, damage
):
    import io

    from PIL import Image

    from linexcel.aidoc import describe_images

    def encode(source):
        stream = io.BytesIO()
        source.save(stream, format="PNG")
        return stream.getvalue()

    original = Image.new("RGB", (40, 30), "white")
    source_path = tmp_path / "original.png"
    source_path.write_bytes(encode(original))
    checks = {}
    describe_images(
        {"S": source_path}, validation_results=checks, evidence_dir=tmp_path / "crops"
    )
    page = checks["S"]["pages"][0]
    batch = page["batches"][0]
    if damage == "fake_receipt":
        batch["provider_receipt"] = {"response_id": "not-a-request"}
    elif damage == "fake_response":
        batch["raw_response"] = '{"regions": []}'
    elif damage == "fake_attempt":
        batch["attempts"] = [{"number": 1, "status": "complete"}]
    elif damage == "tokens":
        batch["usage_total"]["requests"] = 1
    elif damage == "missing_usage":
        batch["usage_total"] = None
    elif damage == "missing_bbox":
        batch["images"][0].pop("bbox")
    elif damage == "missing_decision":
        batch.pop("observations")
    elif damage == "extra_observation":
        page["observations"].append({"region_id": "r0-c0", "text": "Quarter"})
    elif damage == "uniform_metadata":
        batch["images"][0]["pixel_evidence"]["rgba_extrema"][0] = [0, 255]
    elif damage == "nonuniform_crop":
        original.putpixel((20, 15), (254, 254, 254))
        changed = encode(original)
        stored = page["evidence"]["regions"][0]
        Path(stored["path"]).write_bytes(changed)
        digest = hashlib.sha256(changed).hexdigest()
        stored["sha256"] = batch["images"][0]["sha256"] = digest
        batch["observations"][0]["crop_sha256"] = digest
    elif damage == "nonuniform_original":
        original.putpixel((20, 15), (254, 254, 254))
        changed = encode(original)
        source_path.write_bytes(changed)
        # Uniform saved crop with a valid hash cannot replace changed source pixels.
        page["evidence"]["original_sha256"] = batch["original_sha256"] = hashlib.sha256(
            changed
        ).hexdigest()
    manual.link_vision_receipts(checks, [])
    assert not batch.get("receipt_link_complete")
    result = manual.vision_region_coverage(["S"], checks)
    assert result["passed"] is (damage is None)
    assert result["factual_certification"] is False


def test_nonuniform_crop_cannot_gain_coverage_by_claiming_a_local_pixel_check(
    manual, tmp_path
):
    from vision_helpers import image_bytes, region_response

    from linexcel.aidoc import TokenUsage, describe_images

    class Provider:
        def generate(self, *args, **kwargs):
            raise AssertionError("A nonuniform capture requires pixels")

        def generate_with_images(self, system, user, images, **kwargs):
            return region_response(user), TokenUsage(20, 10, 1)

    checks = {}
    describe_images(
        {"S": image_bytes()},
        provider=Provider(),
        validation_results=checks,
        evidence_dir=tmp_path,
    )
    batch = checks["S"]["pages"][0]["batches"][0]
    batch["resolution"] = "deterministic_pixels"
    batch["provider_request_required"] = False
    manual.link_vision_receipts(checks, [])
    assert not manual.vision_region_coverage(["S"], checks)["passed"]
