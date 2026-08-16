import json

import pytest

from linexcel.cli import main


@pytest.fixture()
def workbook_path(lineage_excel: bytes, tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(lineage_excel)
    return path


def test_analyze_writes_html_next_to_workbook(workbook_path):
    assert main(["analyze", str(workbook_path)]) == 0
    html = workbook_path.with_name("book_lineage.html")
    assert "<html" in html.read_text(encoding="utf-8").lower()


def test_analyze_json_output(workbook_path, tmp_path):
    out = tmp_path / "graph.json"
    assert main(["analyze", str(workbook_path), "--no-html", "--json", str(out)]) == 0
    graph = json.loads(out.read_text(encoding="utf-8"))
    assert graph["nodes"] and graph["edges"]


def test_deterministic_only_is_the_default(workbook_path, tmp_path, monkeypatch):
    # An AI call in the default mode would reach aidoc; make that fatal.
    def _fail(*args, **kwargs):
        raise AssertionError("no AI call expected without --ai-docs")

    monkeypatch.setattr("linexcel.aidoc.document_nodes", _fail)
    argv = ["analyze", str(workbook_path), "-o", str(tmp_path / "o.html")]
    assert main(argv) == 0
    assert main([*argv, "--deterministic-only"]) == 0


def test_ai_docs_without_provider_exits_2(workbook_path, capsys, monkeypatch):
    for var in ("LINEXCEL_AI_BASE_URL", "OPENAI_BASE_URL"):
        monkeypatch.delenv(var, raising=False)
    code = main(["analyze", str(workbook_path), "--no-html", "--ai-docs"])
    assert code == 2
    assert "provider" in capsys.readouterr().err.lower()


def test_bad_workbook_exits_2(tmp_path, capsys):
    path = tmp_path / "not_excel.xlsx"
    path.write_text("plain text", encoding="utf-8")
    assert main(["analyze", str(path), "--no-html"]) == 2
    assert "not an Excel file" in capsys.readouterr().err
