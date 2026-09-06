"""Structured AI tables: the model describes, the code renders.

A small local model cannot be trusted with pipe-table syntax, so the prompts
have it emit ``{{Tn}}`` placeholders plus a ```json_tables payload, and
``_insert_tables`` does the rendering. These tests pin the contract: valid
payloads become well-formed tables, invalid ones cost the tables but never
the prose.
"""

import json

from linexcel.aidoc import _insert_tables, render_markdown_table


class TestRenderMarkdownTable:
    def test_basic_table_is_well_formed(self):
        md = render_markdown_table(
            ["Sheet", "Formulas"],
            [["Sales", 80], ["Summary", 68]],
        )
        lines = md.split("\n")
        assert lines[0] == "| Sheet | Formulas |"
        assert lines[1] == "|---|---:|"  # numeric column right-aligned
        assert lines[2] == "| Sales | 80 |"
        assert lines[3] == "| Summary | 68 |"

    def test_pipes_and_newlines_inside_cells_are_neutralised(self):
        md = render_markdown_table(["A"], [["x|y\nz"]])
        assert md == "| A |\n|---|\n| x\\|y z |"

    def test_ragged_rows_cannot_shift_columns(self):
        md = render_markdown_table(["A", "B"], [["only-one"], ["a", "b", "c"]])
        assert "| only-one |  |" in md
        assert "| a | b |" in md

    def test_caption_is_an_italic_line_above_the_table(self):
        md = render_markdown_table(["A"], [[1]], caption="Per sheet")
        assert md.startswith("*Per sheet*\n\n| A |")

    def test_non_numeric_column_stays_left_aligned(self):
        md = render_markdown_table(["Name"], [["Sales"]])
        assert "|---|" in md


class TestInsertTables:
    def test_placeholder_is_replaced_by_the_rendered_table(self):
        payload = [
            {
                "id": "T1",
                "caption": "Sheets",
                "columns": ["Sheet", "Formulas"],
                "rows": [["Sales", 80]],
            }
        ]
        text = (
            "Overview text.\n\n{{T1}}\n\nMore text.\n```json_tables\n"
            + json.dumps(payload)
            + "\n```"
        )
        out = _insert_tables(text)
        assert "```json_tables" not in out
        assert "{{T1}}" not in out
        assert "*Sheets*" in out
        assert "| Sheet | Formulas |" in out
        assert out.startswith("Overview text.")
        assert out.endswith("More text.")

    def test_malformed_json_drops_tables_but_keeps_prose(self):
        text = "Prose stays.\n\n{{T1}}\n```json_tables\n{not json\n```"
        out = _insert_tables(text)
        assert "Prose stays." in out
        assert "{{T1}}" not in out
        assert "```" not in out

    def test_missing_table_drops_its_placeholder_cleanly(self):
        text = "Before.\n\n{{T9}}\n\nAfter.\n```json_tables\n[]\n```"
        out = _insert_tables(text)
        assert "{{T9}}" not in out
        assert "Before." in out and "After." in out
        assert "\n\n\n" not in out

    def test_text_without_a_block_is_returned_stripped(self):
        assert _insert_tables("  plain markdown\n") == "plain markdown"

    def test_entries_with_wrong_shape_are_skipped(self):
        payload = [{"id": "T1", "columns": "oops", "rows": []}, "junk", 42]
        text = "Hi.\n```json_tables\n" + json.dumps(payload) + "\n```"
        assert _insert_tables(text) == "Hi."
