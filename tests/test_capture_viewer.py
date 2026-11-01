"""README capture receipts must describe complete, unchanged image evidence."""

import importlib.util
import json
import sys
from contextlib import nullcontext
from pathlib import Path
from types import SimpleNamespace

import pytest


@pytest.fixture
def capture(monkeypatch, tmp_path):
    script = Path(__file__).resolve().parents[1] / "scripts" / "capture_viewer.py"
    spec = importlib.util.spec_from_file_location("capture_viewer_tests", script)
    module = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, spec.name, module)
    spec.loader.exec_module(module)
    assert Path("validation_workbooks.py") in module.WATCHED_SOURCES
    assert Path("scripts/capture_viewer.py") in module.WATCHED_SOURCES
    sources = (Path("validation_workbooks.py"), Path("capture.py"))
    for relative in sources:
        (tmp_path / relative).write_text("source\n", encoding="utf-8")
    images = tmp_path / "imgs"
    images.mkdir()
    for shot in module.SHOTS:
        (images / shot.name).write_bytes(b"original image")
    report = tmp_path / "report.html"
    report.write_text('  var GRAPH = {"nodes": [], "meta": {}};\n', encoding="utf-8")
    monkeypatch.setattr(module, "ROOT", tmp_path)
    monkeypatch.setattr(module, "IMAGES_DIR", images)
    monkeypatch.setattr(module, "MANIFEST", images / "manifest.json")
    monkeypatch.setattr(module, "WATCHED_SOURCES", sources)
    return module, report


def test_generator_changes_and_changed_images_invalidate_capture(capture):
    module, report = capture
    module.write_manifest(report, {"aiNodeDocumentation": False})
    assert module.check() == 0
    before = json.loads(module.MANIFEST.read_text(encoding="utf-8"))
    (module.ROOT / "validation_workbooks.py").write_text("new fixture\n")
    assert module.check() == 1
    assert module.accept() == 0
    accepted = json.loads(module.MANIFEST.read_text(encoding="utf-8"))
    assert accepted["report"] == before["report"]
    assert accepted["captured"] == before["captured"]
    assert accepted["content"] == {"aiNodeDocumentation": False}
    assert "accepted_without_recapture" in accepted
    assert module.check() == 0
    (module.IMAGES_DIR / module.SHOTS[0].name).write_bytes(b"changed image")
    assert module.check() == 1
    assert module.accept() == 1


def test_accept_cannot_invent_capture_receipt(capture):
    module, _report = capture
    assert module.accept() == 1
    assert not module.MANIFEST.exists()


def test_content_records_actual_ai_presence(capture):
    module, report = capture
    assert not any(module.report_content(report).values())
    graph = {
        "nodes": [{"doc": "A node card"}],
        "meta": {
            "workbookDoc": "Overview",
            "screenshotDocs": {"Sales": "Description"},
        },
    }
    report.write_text("  var GRAPH = " + json.dumps(graph) + ";\n", encoding="utf-8")
    assert all(module.report_content(report).values())


def test_missing_capture_preserves_all_existing_images_and_receipt(
    capture, monkeypatch
):
    module, report = capture
    module.write_manifest(report)
    original = module.MANIFEST.read_bytes()

    class Page:
        def goto(self, *_args):
            pass

        def wait_for_selector(self, *_args, **_kwargs):
            pass

        def wait_for_timeout(self, *_args):
            pass

        def evaluate(self, *_args):
            return {"aiNodeDocumentation": False}

        def screenshot(self, path):
            Path(path).write_bytes(b"new capture")

    browser = SimpleNamespace(new_page=lambda **_kw: Page(), close=lambda: None)
    playwright = SimpleNamespace(chromium=SimpleNamespace(launch=lambda **_kw: browser))
    monkeypatch.setitem(
        sys.modules,
        "playwright.sync_api",
        SimpleNamespace(
            sync_playwright=lambda: nullcontext(playwright),
        ),
    )
    attempts = iter([True, False])
    monkeypatch.setattr(module, "apply_shot", lambda *_: next(attempts))
    assert module.capture(report) == 1
    assert module.MANIFEST.read_bytes() == original
    assert all(
        (module.IMAGES_DIR / shot.name).read_bytes() == b"original image"
        for shot in module.SHOTS
    )
