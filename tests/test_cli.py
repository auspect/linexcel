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
    shots = {"Ventes": [tmp_path / "Ventes.png"]}
    shots["Ventes"][0].write_bytes(b"\x89PNG\r\n\x1a\n")
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
        assert "Ventes" in out
        assert "cells declared" in out

    def test_it_goes_to_stdout_so_it_can_be_piped(self, workbook_path, capsys):
        main(["analyze", str(workbook_path), "--dry-run"])
        captured = capsys.readouterr()
        assert captured.out
        assert "cells declared" not in captured.err

    def test_it_says_which_ceilings_will_apply(self, workbook_path, capsys):
        main(["analyze", str(workbook_path), "--dry-run"])
        assert "ceilings:" in capsys.readouterr().out


class TestSayingHowLongItWillTake:
    """A big workbook announces itself before the wait, not after.

    The estimate comes from the uncompressed size of the sheet parts, which
    the zip index carries — so asking costs about a twentieth of a
    millisecond, and every run can afford to ask.
    """

    def test_the_index_gives_the_weight_without_unpacking_anything(self, workbook_path):
        from linexcel.analyzer import sheet_bytes

        assert sheet_bytes(workbook_path.read_bytes()) > 0

    def test_something_that_is_not_a_package_weighs_nothing(self):
        from linexcel.analyzer import sheet_bytes

        assert sheet_bytes(b"not a zip") == 0

    def test_a_small_workbook_says_nothing(self, workbook_path, capsys):
        main(["analyze", str(workbook_path), "--no-html"])
        assert "should take" not in capsys.readouterr().err

    def test_a_large_one_says_how_long_before_it_starts(
        self, workbook_path, monkeypatch, capsys
    ):
        from linexcel import analyzer

        monkeypatch.setattr(analyzer, "sheet_bytes", lambda data: 200 * 1_048_576)
        main(["analyze", str(workbook_path), "--no-html"])
        err = capsys.readouterr().err
        assert "200 MB of formulas" in err
        assert "about 2 minutes" in err

    def test_it_goes_to_stderr_so_a_piped_report_stays_clean(
        self, workbook_path, monkeypatch, capsys
    ):
        from linexcel import analyzer

        monkeypatch.setattr(analyzer, "sheet_bytes", lambda data: 200 * 1_048_576)
        main(["analyze", str(workbook_path), "--no-html", "--json", "-"])
        captured = capsys.readouterr()
        assert "should take" in captured.err
        assert "should take" not in captured.out

    def test_it_names_the_two_ways_out(self, workbook_path, monkeypatch, capsys):
        """Someone told a run will be long wants to know what else they can do."""
        from linexcel import analyzer

        monkeypatch.setattr(analyzer, "sheet_bytes", lambda data: 200 * 1_048_576)
        main(["analyze", str(workbook_path), "--no-html"])
        err = capsys.readouterr().err
        assert "--dry-run" in err and "-v" in err

    def test_the_dry_run_states_it_too(self, workbook_path, capsys):
        main(["analyze", str(workbook_path), "--dry-run"])
        assert "should take" in capsys.readouterr().out


class TestTheEstimateReadsAsAnOrderOfMagnitude:
    """Quoting seconds would claim a precision it does not have."""

    from linexcel.cli import _format_duration as _fmt

    def test_seconds_are_rounded_to_five(self):
        from linexcel.cli import _format_duration

        assert _format_duration(12) == "about 10 seconds"
        assert _format_duration(0.2) == "about 5 seconds"

    def test_past_a_minute_and_a_half_it_speaks_in_minutes(self):
        from linexcel.cli import _format_duration

        assert _format_duration(100) == "about 2 minutes"
        assert _format_duration(600) == "about 10 minutes"
