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


class TestSayingHowLongItWillTake:
    """A big workbook announces itself before the wait, not after.

    The estimate comes from the uncompressed size of the sheet parts, which
    the zip index carries — so asking costs about a twentieth of a
    millisecond, and every run can afford to ask.
    """

    def test_the_index_gives_the_weight_without_unpacking_anything(self, workbook_path):
        from linexcel.structure import sheet_bytes

        assert sheet_bytes(workbook_path.read_bytes()) > 0

    def test_something_that_is_not_a_package_weighs_nothing(self):
        from linexcel.structure import sheet_bytes

        assert sheet_bytes(b"not a zip") == 0

    def test_a_small_workbook_says_nothing(self, workbook_path, capsys):
        main(["analyze", str(workbook_path), "--no-html"])
        assert "should take" not in capsys.readouterr().err

    def test_a_large_one_says_how_long_before_it_starts(
        self, workbook_path, monkeypatch, capsys
    ):
        from linexcel import structure

        monkeypatch.setattr(structure, "sheet_bytes", lambda data: 200 * 1_048_576)
        main(["analyze", str(workbook_path), "--no-html"])
        err = capsys.readouterr().err
        assert "200 MB of formulas" in err
        assert "about 2 minutes" in err
        assert "an estimate, not a promise" in err

    def test_it_goes_to_stderr_so_a_piped_report_stays_clean(
        self, workbook_path, monkeypatch, capsys
    ):
        from linexcel import structure

        monkeypatch.setattr(structure, "sheet_bytes", lambda data: 200 * 1_048_576)
        main(["analyze", str(workbook_path), "--no-html", "--json", "-"])
        captured = capsys.readouterr()
        assert "should take" in captured.err
        assert "should take" not in captured.out

    def test_it_names_the_ways_out(self, workbook_path, monkeypatch, capsys):
        """Someone told a run will be long wants to know what else they can do."""
        from linexcel import structure

        monkeypatch.setattr(structure, "sheet_bytes", lambda data: 200 * 1_048_576)
        main(["analyze", str(workbook_path), "--no-html"])
        err = capsys.readouterr().err
        assert "--time-budget" in err and "-v" in err

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


class TestTheEstimateCountsFormulas:
    """Weight alone could not see a small file of expensive formulas."""

    def test_formulas_are_counted_in_the_sheet_xml(self, workbook_path):
        from linexcel.structure import count_formulas

        assert count_formulas(workbook_path.read_bytes()) > 0

    def test_something_that_is_not_a_package_has_no_formulas(self):
        from linexcel.structure import count_formulas

        assert count_formulas(b"not a zip") == 0

    def test_a_small_workbook_skips_the_count(self, workbook_path, monkeypatch):
        """Counting unpacks the sheets; a quick run is not worth that."""
        from linexcel import structure

        def _boom(data):
            raise AssertionError("count_formulas should not run under the floor")

        monkeypatch.setattr(structure, "count_formulas", _boom)
        assert structure.estimate_seconds(workbook_path.read_bytes()) >= 0

    def test_a_heavy_one_pays_the_count(self, workbook_path, monkeypatch):
        from linexcel import structure

        monkeypatch.setattr(structure, "sheet_bytes", lambda data: 200 * 1_048_576)
        monkeypatch.setattr(structure, "count_formulas", lambda data: 1_000_000)
        estimate = structure.estimate_seconds(workbook_path.read_bytes())
        # 100 s of reading + 30 s of evaluation
        assert estimate == pytest.approx(130.0)


class TestOverrunNotice:
    """A run that sails past its estimate says so while there is still time
    to act on it, not in a post-mortem."""

    def test_it_fires_when_the_run_overruns(self, monkeypatch, capsys):
        import time

        from linexcel import cli

        monkeypatch.setattr(cli, "OVERRUN_NOTICE_FLOOR_SECONDS", 0.05)
        with cli._overrun_notice(0.0):
            time.sleep(0.3)
        err = capsys.readouterr().err
        assert "still running" in err
        assert "--time-budget" in err

    def test_it_names_the_estimate_it_overran(self, monkeypatch, capsys):
        import time

        from linexcel import cli

        monkeypatch.setattr(cli, "OVERRUN_NOTICE_FLOOR_SECONDS", 0.0)
        monkeypatch.setattr(cli, "OVERRUN_FACTOR", 0.01)
        with cli._overrun_notice(10.0):
            time.sleep(0.3)
        assert "estimate" in capsys.readouterr().err

    def test_a_run_that_finishes_in_time_says_nothing(self, capsys):
        from linexcel import cli

        with cli._overrun_notice(0.0):
            pass
        assert capsys.readouterr().err == ""


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
