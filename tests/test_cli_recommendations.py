"""The run ends with plain-English advice about what a rerun could do better.

Hints go to stderr, after the warnings; the report itself never changes.
"""

import json

import pytest
from test_external import _link, linked_workbook

from linexcel.cli import main
from linexcel.insights import WorkbookRenderError


@pytest.fixture()
def workbook_path(lineage_excel: bytes, tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(lineage_excel)
    return path


@pytest.fixture()
def cached_link_workbook(tmp_path):
    """A workbook reading an external file that is cached but never read."""
    path = tmp_path / "linked.xlsx"
    path.write_bytes(
        linked_workbook(
            {"A1": 2, "B1": "='[1]Annual'!B4 * A1"},
            [_link("Budget.xlsx", "Annual", {"B4": 21})],
        )
    )
    return path


class TestScreenshotRecommendation:
    def test_a_multi_sheet_workbook_is_told_about_the_flag(
        self, workbook_path, capsys
    ):
        assert main(["analyze", str(workbook_path), "--no-html"]) == 0
        hint = capsys.readouterr().err
        assert "hint:" in hint
        assert "--screenshots" in hint

    def test_no_suggestion_once_screenshots_were_rendered(
        self, workbook_path, tmp_path, monkeypatch, capsys
    ):
        monkeypatch.setattr(
            "linexcel.result.LineageResult.save_screenshots",
            lambda self, output_dir, **kwargs: {"Sales": []},
        )
        argv = [
            "analyze",
            str(workbook_path),
            "--no-html",
            "--screenshots",
            str(tmp_path / "shots"),
        ]
        assert main(argv) == 0
        assert "hint:" not in capsys.readouterr().err

    def test_a_single_sheet_workbook_suggests_nothing(self, tmp_path, capsys):
        from test_external import workbook

        path = tmp_path / "one.xlsx"
        path.write_bytes(workbook({"A1": 1, "B1": "=A1*2"}))
        assert main(["analyze", str(path), "--no-html"]) == 0
        assert "hint:" not in capsys.readouterr().err


class TestScreenshotFailureIsExplained:
    """A failed render must not take the finished report down with it."""

    def test_the_report_is_still_written_and_the_reason_named(
        self, workbook_path, tmp_path, monkeypatch, capsys
    ):
        def _boom(self, output_dir, **kwargs):
            raise WorkbookRenderError(
                "Workbook screenshots require LibreOffice and pdftoppm; "
                "LibreOffice could not be found. Install with: ..."
            )

        monkeypatch.setattr(
            "linexcel.result.LineageResult.save_screenshots", _boom
        )
        out = tmp_path / "o.html"
        argv = [
            "analyze",
            str(workbook_path),
            "-o",
            str(out),
            "--screenshots",
            str(tmp_path / "shots"),
        ]
        assert main(argv) == 0
        assert "<html" in out.read_text(encoding="utf-8").lower()
        err = capsys.readouterr().err
        assert "warning:" in err
        assert "LibreOffice could not be found" in err
        assert "hint:" in err and "--screenshots" in err

    def test_nothing_is_dropped_silently(
        self, workbook_path, tmp_path, monkeypatch, capsys
    ):
        def _boom(self, output_dir, **kwargs):
            raise WorkbookRenderError("LibreOffice did not produce a PDF")

        monkeypatch.setattr(
            "linexcel.result.LineageResult.save_screenshots", _boom
        )
        argv = [
            "analyze",
            str(workbook_path),
            "--no-html",
            "--screenshots",
            str(tmp_path / "shots"),
        ]
        assert main(argv) == 0
        assert "did not produce a PDF" in capsys.readouterr().err


class TestRefsDirRecommendation:
    def test_an_unread_external_workbook_gets_a_pointer(
        self, cached_link_workbook, capsys
    ):
        assert main(["analyze", str(cached_link_workbook), "--no-html"]) == 0
        err = capsys.readouterr().err
        assert "hint:" in err
        assert "--refs-dir" in err

    def test_the_hint_is_not_repeated_when_the_warning_already_said_it(
        self, tmp_path, capsys
    ):
        """A link with no cache already triggers the warning's own pointer."""
        path = tmp_path / "missing.xlsx"
        path.write_bytes(
            linked_workbook(
                {"A1": 2, "B1": "='[1]Sheet1'!A1"}, [_link("Gone.xlsx", "Sheet1")]
            )
        )
        assert main(["analyze", str(path), "--no-html"]) == 0
        err = capsys.readouterr().err
        assert "warning:" in err and "--refs-dir" in err
        assert "hint:" not in err


class TestRichSummary:
    def test_the_default_summary_is_one_plain_line(self, workbook_path, capsys):
        assert main(["analyze", str(workbook_path), "--no-html"]) == 0
        err = capsys.readouterr().err
        assert "Nodes:" in err
        assert "Node kind" not in err

    def test_rich_draws_a_table_of_nodes_by_kind(self, workbook_path, capsys):
        pytest.importorskip("rich")
        assert main(["analyze", str(workbook_path), "--no-html", "--rich"]) == 0
        err = capsys.readouterr().err
        assert "Node kind" in err
        assert "Edges:" in err

    def test_rich_degrades_to_plain_text_without_rich(
        self, workbook_path, monkeypatch, capsys
    ):
        monkeypatch.setattr("linexcel.progress._rich_console", lambda: None)
        assert main(["analyze", str(workbook_path), "--no-html", "--rich"]) == 0
        err = capsys.readouterr().err
        assert "Nodes:" in err
        assert "linexcel[progress]" in err

    def test_the_report_is_identical_with_or_without_rich(
        self, workbook_path, tmp_path
    ):
        plain = tmp_path / "plain.json"
        fancy = tmp_path / "fancy.json"
        base = ["analyze", str(workbook_path), "--no-html"]
        assert main([*base, "--json", str(plain)]) == 0
        assert main([*base, "--rich", "--json", str(fancy)]) == 0
        a = json.loads(plain.read_text(encoding="utf-8"))
        b = json.loads(fancy.read_text(encoding="utf-8"))
        a["meta"].pop("analyzedAt")
        b["meta"].pop("analyzedAt")
        assert a == b
