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


def test_vision_docs_without_screenshots_exits_2(workbook_path, capsys):
    """There is nothing to look at until the sheets have been rendered."""
    code = main(["analyze", str(workbook_path), "--no-html", "--vision-docs"])
    assert code == 2
    assert "--screenshots" in capsys.readouterr().err


def test_vision_docs_is_refused_in_deterministic_mode(workbook_path, tmp_path, capsys):
    code = main(
        [
            "analyze",
            str(workbook_path),
            "--no-html",
            "--screenshots",
            str(tmp_path / "shots"),
            "--vision-docs",
            "--deterministic-only",
        ]
    )
    assert code == 2
    assert "--deterministic-only" in capsys.readouterr().err


def test_screenshots_are_rendered_and_embedded(workbook_path, tmp_path, monkeypatch):
    """The flag renders the sheets and puts them in the report, no AI involved."""
    shots = {"Sales": [tmp_path / "Sales.png"]}
    shots["Sales"][0].write_bytes(b"\x89PNG\r\n\x1a\n")
    monkeypatch.setattr(
        "linexcel.result.LineageResult.save_screenshots",
        lambda self, output_dir, **kwargs: shots,
    )
    out = tmp_path / "o.html"
    argv = [
        "analyze",
        str(workbook_path),
        "-o",
        str(out),
        "--screenshots",
        str(tmp_path),
    ]
    assert main(argv) == 0
    assert "data:image/png;base64," in out.read_text(encoding="utf-8")


def test_bad_workbook_exits_2(tmp_path, capsys):
    path = tmp_path / "not_excel.xlsx"
    path.write_text("plain text", encoding="utf-8")
    assert main(["analyze", str(path), "--no-html"]) == 2
    assert "not an Excel file" in capsys.readouterr().err


class TestDryRun:
    """`--dry-run` answers "is this going to be long?" without finding out."""

    def test_it_writes_no_report(self, workbook_path, capsys):
        assert main(["analyze", str(workbook_path), "--dry-run"]) == 0
        assert not workbook_path.with_name("book_lineage.html").exists()

    def test_it_names_the_sheets_and_their_declared_size(self, workbook_path, capsys):
        main(["analyze", str(workbook_path), "--dry-run"])
        out = capsys.readouterr().out
        assert "Sales" in out
        assert "cells declared" in out

    def test_it_goes_to_stdout_so_it_can_be_piped(self, workbook_path, capsys):
        main(["analyze", str(workbook_path), "--dry-run"])
        captured = capsys.readouterr()
        assert captured.out
        assert "cells declared" not in captured.err

    def test_it_says_which_ceilings_will_apply(self, workbook_path, capsys):
        main(["analyze", str(workbook_path), "--dry-run"])
        assert "ceilings:" in capsys.readouterr().out


def test_cli_no_heuristic_duration(workbook_path, capsys):
    assert main(["analyze", str(workbook_path), "--no-html"]) == 0
    assert "should take" not in capsys.readouterr().err


def test_dry_run_states_unknown_duration(workbook_path, capsys):
    assert main(["analyze", str(workbook_path), "--dry-run"]) == 0
    assert "duration is unknown" in capsys.readouterr().out


def test_budget_expiry_returns_partial_report(workbook_path, tmp_path):
    out = tmp_path / "partial.json"
    assert (
        main(
            [
                "analyze",
                str(workbook_path),
                "--no-html",
                "--json",
                str(out),
                "--analysis-seconds",
                "0",
            ]
        )
        == 3
    )
    assert json.loads(out.read_text())["meta"]["execution"]["status"] == "timed_out"


class TestTargetOption:
    """--target SHEET!A1: the analysis is the upstream subgraph, and nothing else."""

    def test_the_graph_is_limited_to_the_upstream_subgraph(
        self, workbook_path, tmp_path
    ):
        out = tmp_path / "g.json"
        argv = [
            "analyze",
            str(workbook_path),
            "--no-html",
            "--json",
            str(out),
            "--target",
            "Summary!B1",
        ]
        assert main(argv) == 0
        graph = json.loads(out.read_text(encoding="utf-8"))
        assert graph["targets"] == ["Summary!B1"]
        ids = {n["id"] for n in graph["nodes"]}
        assert "c:Summary!B1" in ids
        assert "g:Sales!D2#100" in ids
        # The neighbouring formulas the target does not read are omitted.
        assert "c:Summary!B2" not in ids
        assert "c:Summary!B3" not in ids

    def test_several_targets_comma_separated_and_repeated(
        self, workbook_path, tmp_path
    ):
        out = tmp_path / "g.json"
        argv = [
            "analyze",
            str(workbook_path),
            "--no-html",
            "--json",
            str(out),
            "--target",
            "Summary!B1,Summary!B2",
            "--target",
            "Summary!B3",
        ]
        assert main(argv) == 0
        graph = json.loads(out.read_text(encoding="utf-8"))
        assert graph["targets"] == ["Summary!B1", "Summary!B2", "Summary!B3"]
        ids = {n["id"] for n in graph["nodes"]}
        assert {"c:Summary!B1", "c:Summary!B2", "c:Summary!B3"} <= ids

    def test_an_invalid_target_exits_2(self, workbook_path, capsys):
        code = main(["analyze", str(workbook_path), "--no-html", "--target", "B1"])
        assert code == 2
        assert "Invalid target" in capsys.readouterr().err

    def test_a_target_on_an_unknown_sheet_exits_2(self, workbook_path, capsys):
        code = main(["analyze", str(workbook_path), "--no-html", "--target", "Gone!A1"])
        assert code == 2
        assert "does not have" in capsys.readouterr().err
