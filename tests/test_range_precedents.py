"""Regression tests for inverted and anchor-crossing range precedents."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from linexcel import analyze
from linexcel.decompose import _ref_preview
from linexcel.refs import (
    MAX_COL,
    MAX_ROW,
    Rect,
    parse_ref,
    parse_ref_detailed,
    stretch_ref,
)
from linexcel.tables import _collect_defined_names


def _wired_cells(
    graph: dict, target_id: str, max_row: int = 20, max_col: int = 20
) -> set[tuple[int, int]]:
    wired: set[tuple[int, int]] = set()
    for edge in graph["edges"]:
        if edge["target"] != target_id:
            continue
        source = edge["source"]
        if ":" not in source:
            continue
        rect = parse_ref(source.split(":", 1)[1])
        if rect is None:
            continue
        for row in range(rect.r1, min(rect.r2, max_row) + 1):
            for col in range(rect.c1, min(rect.c2, max_col) + 1):
                wired.add((row, col))
    return wired


class TestParseRefDetailedNormalization:
    def test_reversed_columns_normalized_with_anchors_preserved(self):
        detail = parse_ref_detailed("B5:A7", "S")
        assert detail is not None
        assert detail.rect == parse_ref("B5:A7", "S") == Rect("S", 5, 1, 7, 2)
        assert detail.rect.ncells == 6
        assert detail.rect.clipped(10, 10) == Rect("S", 5, 1, 7, 2)

        anchored = parse_ref_detailed("$B$5:A7", "S")
        assert anchored is not None
        assert anchored.rect == Rect("S", 5, 1, 7, 2)
        assert (anchored.row1_abs, anchored.row2_abs) == (True, False)
        assert (anchored.col1_abs, anchored.col2_abs) == (False, True)

    def test_reversed_rows_and_both_axes_preserve_anchors(self):
        rows_rev = parse_ref_detailed("A7:$B$5", "S")
        assert rows_rev is not None
        assert rows_rev.rect == Rect("S", 5, 1, 7, 2)
        assert (rows_rev.row1_abs, rows_rev.row2_abs) == (True, False)
        assert (rows_rev.col1_abs, rows_rev.col2_abs) == (False, True)

        both_rev = parse_ref_detailed("$B$7:A5", "S")
        assert both_rev is not None
        assert both_rev.rect == Rect("S", 5, 1, 7, 2)
        assert (both_rev.row1_abs, both_rev.row2_abs) == (False, True)
        assert (both_rev.col1_abs, both_rev.col2_abs) == (False, True)

    def test_reversed_whole_columns_and_rows_preserve_anchors(self):
        cols = parse_ref_detailed("$B:A", "S")
        assert cols is not None
        assert cols.rect == Rect("S", 1, 1, MAX_ROW, 2)
        assert (cols.col1_abs, cols.col2_abs) == (False, True)

        rows = parse_ref_detailed("$7:5", "S")
        assert rows is not None
        assert rows.rect == Rect("S", 5, 1, 7, MAX_COL)
        assert (rows.row1_abs, rows.row2_abs) == (False, True)


class TestStretchRefShapes:
    def test_rows_one_end_anchored_crossed(self):
        detail = parse_ref_detailed("A1:$A$5", "S")
        assert detail is not None
        # B1:B10 filled down: B1 reads A1:A5, B10 reads A5:A10 -> union A1:A10.
        assert stretch_ref(detail, 1, 2, (1, 10), (2, 2)) == Rect("S", 1, 1, 10, 1)
        # B1:B3 stays above anchor: exact union is A1:A5, no extra rows.
        assert stretch_ref(detail, 1, 2, (1, 3), (2, 2)) == Rect("S", 1, 1, 5, 1)

    def test_rows_anchored_start_crossed_upward(self):
        detail = parse_ref_detailed("$A$5:A8", "S")
        assert detail is not None
        # Rep at row 5, group spans rows 1..10: row 1 reads A4:A5, row 10 reads A5:A13.
        assert stretch_ref(detail, 5, 2, (1, 10), (2, 2)) == Rect("S", 4, 1, 13, 1)

    def test_columns_one_end_anchored_crossed(self):
        detail = parse_ref_detailed("C$1:$C$1", "S")
        assert detail is not None
        # A1 holds =SUM(C$1:$C$1), B1 holds =SUM(D$1:$C$1) -> union C1:D1.
        assert stretch_ref(detail, 1, 1, (1, 1), (1, 2)) == Rect("S", 1, 3, 1, 4)

    def test_columns_both_ends_moving_high_to_low(self):
        detail = parse_ref_detailed("B$2:A$3", "S")
        assert detail is not None
        # A1 holds =SUM(B$2:A$3) [A2:B3], B1 holds =SUM(C$2:B$3) [B2:C3] -> union A2:C3.
        assert stretch_ref(detail, 1, 1, (1, 1), (1, 2)) == Rect("S", 2, 1, 3, 3)


class TestAnalyzeRangePrecedents:
    def test_ten_row_anchored_end_crossed_wires_all_precedents(self, tmp_path: Path):
        path = tmp_path / "anchored_end_crossed.xlsx"
        book = Workbook()
        sheet = book.active
        assert sheet is not None
        sheet.title = "S"
        for row in range(1, 11):
            sheet.cell(row=row, column=1, value=row)
            sheet.cell(row=row, column=2, value=f"=SUM(A{row}:$A$5)")
        book.save(path)

        graph = analyze(path).graph
        groups = [n for n in graph["nodes"] if n.get("kind") == "group"]
        assert len(groups) == 1
        group_id = groups[0]["id"]
        assert group_id == "g:S!B1#10"
        expected = {(row, 1) for row in range(1, 11)}
        assert _wired_cells(graph, group_id) == expected

    def test_lone_formula_reversed_columns_wires_precedents(self, tmp_path: Path):
        path = tmp_path / "lone_reversed_cols.xlsx"
        book = Workbook()
        sheet = book.active
        assert sheet is not None
        sheet.title = "S"
        value = 1
        for row in range(1, 8):
            for col in (1, 2):
                sheet.cell(row=row, column=col, value=value)
                value += 1
        sheet.cell(row=1, column=4, value="=SUM(B5:A7)")
        book.save(path)

        graph = analyze(path).graph
        expected = {(row, col) for row in range(5, 8) for col in (1, 2)}
        assert _wired_cells(graph, "c:S!D1") == expected

    def test_column_anchored_and_high_to_low_groups_wire_all_precedents(
        self, tmp_path: Path
    ):
        path = tmp_path / "column_shapes.xlsx"
        book = Workbook()
        sheet = book.active
        assert sheet is not None
        sheet.title = "S"
        for row in range(1, 4):
            for col in range(1, 5):
                sheet.cell(row=row, column=col, value=row * 10 + col)
        # Columns: one end anchored, one end moving
        sheet.cell(row=5, column=1, value="=SUM(C$1:$C$1)")
        sheet.cell(row=5, column=2, value="=SUM(D$1:$C$1)")
        # Columns: both ends moving, written high to low
        sheet.cell(row=6, column=1, value="=SUM(B$2:A$3)")
        sheet.cell(row=6, column=2, value="=SUM(C$2:B$3)")
        book.save(path)

        graph = analyze(path).graph
        assert _wired_cells(graph, "g:S!A5#2") == {(1, 3), (1, 4)}
        assert _wired_cells(graph, "g:S!A6#2") == {
            (row, col) for row in (2, 3) for col in (1, 2, 3)
        }

    def test_other_callers_see_normalized_rect_for_reversed_columns(self):
        class _DummyResolver:
            def external_value(self, _ref):
                return None, None

            def value(self, _sheet, _row, _col):
                return None, None, None

        preview, date_text = _ref_preview(_DummyResolver(), "B5:A7", "S")  # type: ignore[arg-type]
        assert date_text is None
        assert preview == {"range": "S!A5:B7", "n": 6}

        book = Workbook()
        sheet = book.active
        assert sheet is not None
        sheet.title = "S"
        book.defined_names.add(DefinedName("RevBox", attr_text="S!$B$5:$A$7"))
        names = _collect_defined_names(book)
        assert names["RevBox"] == [Rect("S", 5, 1, 7, 2)]
