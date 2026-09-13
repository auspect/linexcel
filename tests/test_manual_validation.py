"""The manual gate must fail for incomplete languages and preserve evidence."""

import importlib.util
import json
import sys
from argparse import Namespace
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


@pytest.mark.parametrize(
    "missing", ["nodes", "overview", "vision", "images", "sheet", None]
)
def test_each_language_and_stage_must_complete(manual, monkeypatch, tmp_path, missing):
    image = tmp_path / "sheet.png"
    image_module = pytest.importorskip("PIL.Image")
    image_module.new("RGB", (40, 30), "white").save(image)
    result = SimpleNamespace(
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

    def document(result, args, language):
        return (
            {} if language == "en" and missing == "nodes" else {"S!A1": "Complete"},
            " " if language == "en" and missing == "overview" else "Overview",
        )

    monkeypatch.setattr(manual, "document", document)
    monkeypatch.setattr(
        manual,
        "describe",
        lambda result, shots, args, language: (
            {} if language == "en" and missing == "vision" else {"S": "Visible image"}
        ),
    )
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
