"""Versioned native corrections must survive linexcel's import adaptations.

Older supported engines are skipped rather than making their known incorrect
results an acceptance oracle. These cases do not certify row visibility or
locale-dependent date parsing.
"""

import io
from importlib.metadata import version

import formualizer as fz
import pytest
from openpyxl import Workbook
from packaging.version import Version

from linexcel.engine import _open_workbook

FORMUALIZER_VERSION = Version(version("formualizer"))


@pytest.fixture(params=["native", "linexcel"])
def mode(request):
    return request.param


def _evaluate_workbook(cells, mode):
    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    for address, value in cells.items():
        sheet[address] = value
    stream = io.BytesIO()
    book.save(stream)
    data = stream.getvalue()
    engine = (
        fz.Workbook.from_bytes(data)
        if mode == "native"
        else _open_workbook(data, parallel=False)
    )
    engine.evaluate_all()
    return engine


@pytest.mark.skipif(
    FORMUALIZER_VERSION < Version("0.11.0"),
    reason="Left-associative powers and AND/OR error propagation require 0.11.0.",
)
@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=2^3^2", 64),
        ("=AND(FALSE,1/0)", {"type": "Error", "kind": "Div"}),
        ("=OR(TRUE,1/0)", {"type": "Error", "kind": "Div"}),
    ],
)
def test_power_associativity_and_logical_errors(mode, formula, expected):
    engine = _evaluate_workbook({"D1": formula}, mode)
    assert engine.get_value("S", 1, 4) == expected


@pytest.mark.skipif(
    FORMUALIZER_VERSION < Version("0.11.1"),
    reason="Literal percent, underscore and backslash criteria require 0.11.1.",
)
@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ('=COUNTIF(A1:A7,"a%b")', 1),
        ('=COUNTIF(A1:A7,"a_b")', 1),
        ('=COUNTIF(A1:A7,"a\\b")', 1),
        ('=SUMIF(A1:A7,"a%b",B1:B7)', 1),
        ('=SUMIF(A1:A7,"a_b",B1:B7)', 3),
        ('=SUMIF(A1:A7,"a\\b",B1:B7)', 4),
        ('=AVERAGEIF(A1:A7,"a%b",B1:B7)', 1),
        ('=AVERAGEIF(A1:A7,"a_b",B1:B7)', 3),
        ('=AVERAGEIF(A1:A7,"a\\b",B1:B7)', 4),
        # Excel's own wildcard and its escape retain their meaning.
        ('=COUNTIF(A1:A7,"a*b")', 7),
        ('=COUNTIF(A1:A7,"a~*b")', 1),
    ],
)
def test_criteria_keep_sql_like_characters_literal(mode, formula, expected):
    labels = ["a%b", "axb", "a_b", "a\\b", "a*b", "a?b", "a~b"]
    cells = {f"A{row}": label for row, label in enumerate(labels, 1)}
    cells.update({f"B{row}": row for row in range(1, 8)})
    cells["D1"] = formula

    engine = _evaluate_workbook(cells, mode)

    # The three literal criteria each match exactly one input, with weights
    # 1, 3 and 4. The unescaped Excel wildcard matches all seven inputs.
    assert engine.get_value("S", 1, 4) == expected


@pytest.mark.skipif(
    FORMUALIZER_VERSION < Version("0.11.1"),
    reason="Literal percent and underscore criteria require 0.11.1.",
)
@pytest.mark.xfail(
    FORMUALIZER_VERSION == Version("0.14.0"),
    strict=True,
    raises=AssertionError,
    reason="0.14.0 copied COUNTIF families treat literal _/% as wildcards; "
    "upstream fix is unreleased (2026-11-13 upgrade-plan addendum).",
)
@pytest.mark.parametrize("criterion", ["a_c", "50%"])
def test_copied_countif_families_keep_sql_like_characters_literal(mode, criterion):
    labels = ["a_c", "abc", "50%", "50x"]
    cells = {}
    matching_rows = []
    for row in range(1, 41):
        cells[f"A{row}"] = labels[(row - 1) % len(labels)]
        cells[f"C{row}"] = "a_c" if row % 2 else "50%"
        cells[f"D{row}"] = f"=COUNTIF($A$1:$A$40,C{row})"
        if cells[f"C{row}"] == criterion:
            matching_rows.append(row)

    engine = _evaluate_workbook(cells, mode)

    # Each literal occurs ten times; its SQL wildcard lookalike must not count.
    assert [engine.get_value("S", row, 4) for row in matching_rows] == [10] * 20


@pytest.mark.skipif(
    FORMUALIZER_VERSION < Version("0.11.1"),
    reason="AGGREGATE options 4-7 require 0.11.1.",
)
@pytest.mark.parametrize(
    ("option", "expected"),
    [
        (4, {"type": "Error", "kind": "Div"}),
        (5, {"type": "Error", "kind": "Div"}),
        (6, 4),
        (7, 4),
    ],
)
def test_aggregate_options_propagate_or_ignore_a_real_error(mode, option, expected):
    engine = _evaluate_workbook(
        {"B1": 1, "B2": "=1/0", "B3": 3, "D1": f"=AGGREGATE(9,{option},B1:B3)"},
        mode,
    )

    assert engine.get_value("S", 2, 2) == {"type": "Error", "kind": "Div"}
    assert engine.get_value("S", 1, 4) == expected


_DYNAMIC_ARRAY_METADATA = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" '
    'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" '
    'splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" '
    'coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" '
    'count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
    '</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk>'
    "</cellMetadata></metadata>"
)


def _saved_dynamic_spill():
    """A file as Excel saves a spill: anchor formula plus plain cached members."""
    import re
    import zipfile

    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    for address in ("A1", "A2", "A3", "C1"):
        sheet[address] = 0
    source = io.BytesIO()
    book.save(source)
    cells = {
        "A1": '<c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(3)</f>'
        "<v>1</v></c>",
        "A2": '<c r="A2"><v>2</v></c>',
        "A3": '<c r="A3"><v>3</v></c>',
        "C1": '<c r="C1"><f>SUM(A1#)</f><v>6</v></c>',
    }
    target = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(source.getvalue())) as zin,
        zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout,
    ):
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode()
                for address, xml in cells.items():
                    text = re.sub(rf'<c r="{address}"[^>]*>.*?</c>', xml, text)
                data = text.encode()
            elif item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/metadata.xml" ContentType="application/'
                    b"vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata"
                    b'+xml"/></Types>',
                )
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.'
                    b'org/officeDocument/2006/relationships/sheetMetadata" '
                    b'Target="metadata.xml"/></Relationships>',
                )
            zout.writestr(item, data)
        zout.writestr("xl/metadata.xml", _DYNAMIC_ARRAY_METADATA)
    return target.getvalue()


@pytest.mark.skipif(
    FORMUALIZER_VERSION < Version("0.14.0"),
    reason="0.13.0 reads a saved spill back as #SPILL! because its cached members "
    "block the anchor; values.UNCOMPUTED_ERROR_KINDS keeps the file's value then.",
)
def test_saved_dynamic_spill_recalculates(mode):
    data = _saved_dynamic_spill()
    engine = (
        fz.Workbook.from_bytes(data)
        if mode == "native"
        else _open_workbook(data, parallel=False)
    )
    engine.evaluate_all()

    assert [engine.get_value("S", row, 1) for row in (1, 2, 3)] == [1, 2, 3]
    assert engine.get_value("S", 1, 3) == 6
