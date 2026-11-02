"""Capture evidence must distinguish optional AI from deterministic screenshots."""

import copy
import importlib.util
import json
import struct
from pathlib import Path
from types import SimpleNamespace

import pytest


@pytest.fixture(scope="module")
def capture_app():
    path = Path(__file__).parents[1] / "scripts" / "capture_app.py"
    spec = importlib.util.spec_from_file_location("app_capture_tests", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture
def document(capture_app):
    return {
        "id": "documentation-task",
        "operation": "document",
        "nodeId": capture_app.NODE,
        "status": "succeeded",
        "result": {
            "language": "en",
            "model": "test-model",
            "markdown": "The cell uses the source formula.",
            "usage": {"inputTokens": 10, "outputTokens": 12},
            "validation": {
                "specialized_review": {"stages": [{"usage": {"requests": 2}}]}
            },
        },
    }


def manifest_for(capture_app, directory, document=None, *, no_ai=True):
    images = {}
    for name, caption in capture_app.shot_captions(bool(document)).items():
        # This test exercises manifest validation, not image decoding.
        data = b"\x89PNG\r\n\x1a\n" + b"\x00\x00\x00\rIHDR"
        data += struct.pack(">II", capture_app.WIDTH * 2, capture_app.HEIGHT * 2)
        path = directory / name
        path.write_bytes(data)
        images[name] = {
            "caption": caption,
            "sha256": capture_app.digest(path),
            "pixels": [capture_app.WIDTH * 2, capture_app.HEIGHT * 2],
        }
    return {
        "schema": capture_app.SCHEMA,
        "uiLanguage": "en",
        "viewport": {"width": 1440, "height": 900, "scale": 2},
        "sources": capture_app.source_hashes(),
        "fixture": {
            "name": "fy26_commercial_review.xlsx",
            "generator": "validation_workbooks.build_sales_workbook",
            "generatorSha256": capture_app.source_digest(
                capture_app.ROOT / "validation_workbooks.py"
            ),
            "sha256": "a" * 64,
        },
        "images": images,
        "capture": {"taskId": "capture-task", "imageCount": 4},
        "modelRequestsDuringCapture": 0,
        **capture_app.document_evidence(
            document, SimpleNamespace(no_ai=no_ai, generate_doc=False)
        ),
    }


def write_manifest(capture_app, directory, manifest):
    (directory / capture_app.MANIFEST_NAME).write_text(json.dumps(manifest), "utf-8")


@pytest.mark.parametrize("no_ai", [False, True])
def test_no_document_is_valid_and_records_why(capture_app, tmp_path, no_ai):
    manifest = manifest_for(capture_app, tmp_path, no_ai=no_ai)
    assert manifest["document"] is None
    assert manifest["documentLanguage"] is None
    assert manifest["documentation"]["status"] == "skipped"
    assert manifest["documentation"]["reason"] == (
        "--no-ai" if no_ai else "no existing English document"
    )
    write_manifest(capture_app, tmp_path, manifest)
    assert capture_app.check(tmp_path) == 0


@pytest.mark.parametrize(
    "change",
    [
        {"modelRequestsDuringCapture": 1},
        {"documentLanguage": "en"},
        {"document": {"model": "unrequested-model"}},
        {"documentation": {"status": "failed", "requested": True}},
        {"documentation": {"status": "skipped", "requested": True}},
        {"capture": {"taskId": "capture-task", "imageCount": 3}},
    ],
)
def test_skipped_ai_does_not_hide_missing_or_requested_work(
    capture_app, tmp_path, change
):
    manifest = manifest_for(capture_app, tmp_path)
    manifest.update(change)
    write_manifest(capture_app, tmp_path, manifest)
    assert capture_app.check(tmp_path) == 1


def test_existing_document_keeps_strict_provenance(capture_app, tmp_path, document):
    manifest = manifest_for(capture_app, tmp_path, document)
    write_manifest(capture_app, tmp_path, manifest)
    assert capture_app.check(tmp_path) == 0
    for field in ("model", "usage", "providerRequests", "markdownSha256"):
        broken = copy.deepcopy(manifest)
        del broken["document"][field]
        write_manifest(capture_app, tmp_path, broken)
        assert capture_app.check(tmp_path) == 1


def test_requested_generation_cannot_be_recorded_as_skipped(capture_app):
    with pytest.raises(ValueError, match="did not complete"):
        capture_app.document_evidence(
            None, SimpleNamespace(no_ai=False, generate_doc=True)
        )


def test_conflicting_ai_selection_fails_before_capture(capture_app):
    with pytest.raises(SystemExit) as exc:
        capture_app.main(["--no-ai", "--generate-doc"])
    assert exc.value.code == 2


def test_hashes_and_captions_still_bind_deterministic_images(capture_app, tmp_path):
    manifest = manifest_for(capture_app, tmp_path)
    name = "app_graph_documented.png"
    manifest["images"][name]["caption"] = capture_app.SHOTS[name]
    write_manifest(capture_app, tmp_path, manifest)
    assert capture_app.check(tmp_path) == 1
    manifest["images"][name]["caption"] = capture_app.DETERMINISTIC_SHOTS[name]
    write_manifest(capture_app, tmp_path, manifest)
    (tmp_path / name).write_bytes((tmp_path / name).read_bytes() + b"changed")
    assert capture_app.check(tmp_path) == 1


@pytest.mark.parametrize("operation", ["document", "describe_capture"])
def test_default_network_guard_blocks_ai_before_send(capture_app, operation):
    calls, requests = [], []
    route = SimpleNamespace(
        request=SimpleNamespace(
            method="POST",
            url="http://localhost/review/api/projects/p/tasks",
            post_data_json={"operation": operation},
        ),
        abort=lambda: calls.append("abort"),
        continue_=lambda: calls.append("continue"),
    )
    capture_app.guard_model_request(route, allowed=False, requests=requests)
    assert calls == ["abort"]
    with pytest.raises(ValueError, match="Unexpected model"):
        capture_app.check_model_requests(requests, allowed=False)


def test_opt_in_allows_only_one_document_request(capture_app):
    calls, requests = [], []
    route = SimpleNamespace(
        request=SimpleNamespace(
            method="POST",
            url="http://localhost/review/api/projects/p/tasks",
            post_data_json={"operation": "document"},
        ),
        abort=lambda: calls.append("abort"),
        continue_=lambda: calls.append("continue"),
    )
    capture_app.guard_model_request(route, allowed=True, requests=requests)
    capture_app.check_model_requests(requests, allowed=True)
    capture_app.guard_model_request(route, allowed=True, requests=requests)
    assert calls == ["continue", "abort"]
    with pytest.raises(ValueError, match="Unexpected model"):
        capture_app.check_model_requests(requests, allowed=True)
