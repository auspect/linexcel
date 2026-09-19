"""The fast error scan preserves errors without confusing empty cells."""

import io
import zipfile

from linexcel.loader import _error_cached_values


def package(cells, sheet="S"):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            f'<workbook><sheet name="{sheet}" r:id="rId1"/></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships><Relationship Id="rId1" '
            'Target="worksheets/sheet1.xml"/></Relationships>',
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            f"<worksheet><sheetData>{cells}</sheetData></worksheet>",
        )
    return buffer.getvalue()


def test_self_closing_error_cell_cannot_steal_the_next_error_value():
    data = package('<c r="A1" t="e"/><c r="B1" t="e"><v>#DIV/0!</v></c>')
    assert _error_cached_values(data) == {("S", 1, 2): "#DIV/0!"}


def test_ordinary_values_do_not_appear_in_error_overlay():
    data = package('<c r="A1"><v>3</v></c><c r="B1" t="str"><v>text</v></c>')
    assert _error_cached_values(data) == {}


def test_sheet_identity_decodes_xml_entities_exactly_once():
    data = package(
        '<c r="A1" t="e"><v>#VALUE!</v></c>',
        sheet="R&amp;D &apos;2026&apos; &amp;copy;",
    )
    assert _error_cached_values(data) == {("R&D '2026' &copy;", 1, 1): "#VALUE!"}
