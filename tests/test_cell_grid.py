"""Per-sheet grids for literal input cells (index schema 5).

A grid keeps a literal cell only when the node read back from it equals the
node the XML reader produced, field by field, except ``rawXmlValue``. These
tests check that criterion on every literal kind, the cells that must stay
nodes, the schema check on open and the calculation's selection of grid cells.
"""

import copy
import datetime as dt
import io
import re
import sqlite3
import zipfile

import pytest
from openpyxl import Workbook

from linexcel import lazy_store
from linexcel.lazy_store import (
    IndexSchemaError,
    WorkbookIndex,
    _pack,
    _unpack,
    build_index,
)


def _recording(monkeypatch):
    """The literal input nodes the reader hands to the index, as produced."""
    produced = {}
    original = WorkbookIndex.put

    def put(self, node, *, only_if_absent=False):
        if node.get("kind") == "input" and node.get("formula") is None:
            produced[node["id"]] = copy.deepcopy(node)
        return original(self, node, only_if_absent=only_if_absent)

    monkeypatch.setattr(WorkbookIndex, "put", put)
    return produced


def _without_raw(node):
    node = _unpack(_pack(node))  # the stored form: key order and JSON types
    node["sourceValueProvenance"].pop("rawXmlValue")
    return node


def _grid_ids(index):
    """Cell ids read back from the index that are not in its nodes table."""
    with WorkbookIndex(index) as store:
        stored = {nid for (nid,) in store.db.execute("SELECT id FROM nodes")}
        return {nid for nid in store.cell_ids() if nid not in stored}


def _sheet_xml(book, sheet_data):
    """``book`` saved with its first sheet's cells replaced by raw XML."""
    source, output = io.BytesIO(), io.BytesIO()
    book.save(source)
    with zipfile.ZipFile(source) as archive, zipfile.ZipFile(output, "w") as result:
        for name in archive.namelist():
            data = archive.read(name)
            if name == "xl/worksheets/sheet1.xml":
                data = re.sub(
                    rb"<sheetData\s*/>|<sheetData>.*</sheetData>",
                    b"<sheetData>" + sheet_data.encode() + b"</sheetData>",
                    data,
                    flags=re.S,
                )
            result.writestr(name, data)
    return output.getvalue()


def test_every_literal_kind_reads_back_as_the_reader_produced_it(tmp_path, monkeypatch):
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["Name", 12, 3.25, dt.datetime(2024, 2, 29, 6, 30), True, "#N/A"])
    sheet.append(["Été", -7, 1e-05, dt.time(18, 15), False, "#DIV/0!"])
    sheet["G1"] = dt.timedelta(hours=27, minutes=30)
    sheet["G1"].number_format = "[h]:mm:ss"
    sheet["H2"] = 2**53 + 1
    sheet["I2"] = 0.1 + 0.2
    sheet.cell(3, 600, "far right")
    sheet.cell(4, 2, 44000).number_format = "0.00"
    sheet.row_dimensions[2].hidden = True
    sheet.column_dimensions["C"].hidden = True
    hidden = book.create_sheet("Hidden")
    hidden.sheet_state = "hidden"
    hidden["A1"] = 1
    path = tmp_path / "kinds.xlsx"
    book.save(path)
    produced = _recording(monkeypatch)
    build_index(path, tmp_path / "index.sqlite")
    in_grid = _grid_ids(tmp_path / "index.sqlite")
    # Every literal of the workbook is a grid cell, the wide one included.
    assert in_grid == set(produced)
    assert "Data!WB3" in in_grid
    with WorkbookIndex(tmp_path / "index.sqlite") as store:
        for nid, node in produced.items():
            sheet_name, row, col = node["sheet"], node["row"], node["column"]
            read = store.grid.get(sheet_name, row, col)
            expected = _without_raw(node)
            assert read == expected, nid
            assert list(read) == list(expected), nid
            assert list(read["sourceValueProvenance"]) == list(
                expected["sourceValueProvenance"]
            )
        assert store["Data!D1"]["value"] == "2024-02-29T06:30:00"
        assert store["Data!D1"]["cached_comparison_value"] == pytest.approx(45351.27083)
        assert store["Data!G1"]["value"] == "1 day, 3:30:00"
        assert store["Data!E1"]["value"] is True
        assert store["Data!F2"]["cachedValueKind"] == "error"
        # openpyxl saves this integer as the nearest double; the grid keeps
        # the type the reader parsed.
        big = store["Data!H2"]["value"]
        assert (big, type(big)) == (
            produced["Data!H2"]["value"],
            type(produced["Data!H2"]["value"]),
        )
        assert store["Data!B2"]["context"]["rowHidden"] is True
        assert store["Data!C1"]["context"]["columnHidden"] is True
        assert store["Hidden!A1"]["context"]["sheetState"] == "hidden"
        assert "rawXmlValue" not in store["Data!B1"]["sourceValueProvenance"]


def test_cells_a_grid_cannot_rebuild_exactly_stay_nodes(tmp_path):
    book = Workbook()
    book.active.title = "S"
    data = _sheet_xml(
        book,
        '<row r="1">'
        '<c r="A1" t="d"><v>2024-01-02T03:04:05</v></c>'
        '<c r="B1" t="s"><v>999</v></c>'
        '<c r="C1"><f>1+1</f><v>2</v></c>'
        '<c r="D1" t="inlineStr"><is><t>inline</t></is></c>'
        '<c r="E1" t="str"><v>plain</v></c>'
        "</row>"
        '<row r="2"><c r="B2"><v>5</v></c><c r="A2"><v>4</v></c>'
        '<c r="C2"><v>1</v></c><c r="C2"><v>2</v></c></row>'
        '<row r="3"><c><v>30</v></c><c><v>31</v></c></row>',
    )
    path = tmp_path / "odd.xlsx"
    path.write_bytes(data)
    build_index(path, tmp_path / "index.sqlite")
    in_grid = _grid_ids(tmp_path / "index.sqlite")
    with WorkbookIndex(tmp_path / "index.sqlite") as store:
        # ISO date, unreadable text, formula: nodes, raw XML kept.
        assert store["S!A1"]["sourceValueProvenance"]["rawXmlValue"] == (
            "2024-01-02T03:04:05"
        )
        assert store["S!B1"]["diagnostics"]
        assert store["S!C1"]["kind"] == "cell"
        # A cell listed after a later one of its row stays a node; a cell
        # listed twice keeps its last value, as a node replaced it before.
        assert (store["S!A2"]["value"], store["S!B2"]["value"]) == (4, 5)
        assert store["S!C2"]["value"] == 2
        # Inline and plain strings, and cells without an address, are exact.
        assert store["S!D1"]["value"] == "inline"
        assert store["S!E1"]["value"] == "plain"
        assert store["S!A3"]["sourceValueProvenance"]["cellAddressOrigin"] == (
            "inferred"
        )
        assert store["S!B3"]["value"] == 31
        assert len(store) == 10
    assert {"S!A1", "S!B1", "S!C1", "S!A2", "S!C2"}.isdisjoint(in_grid)
    assert {"S!D1", "S!E1", "S!B2", "S!A3", "S!B3"} <= in_grid


def test_an_index_of_another_schema_is_refused(tmp_path):
    path = tmp_path / "index.sqlite"
    with sqlite3.connect(path) as db:
        db.execute("CREATE TABLE metadata(key TEXT PRIMARY KEY,value TEXT NOT NULL)")
        db.execute("INSERT INTO metadata VALUES ('meta','{\"schemaVersion\":4}')")
    with pytest.raises(IndexSchemaError) as refused:
        WorkbookIndex(path)
    assert refused.value.found is None
    with sqlite3.connect(path) as db:
        db.execute("INSERT INTO metadata VALUES ('schemaVersion','4')")
    with pytest.raises(IndexSchemaError, match="reimport"):
        WorkbookIndex(path)
    assert lazy_store.INDEX_FILENAME == "index-v5.sqlite"


@pytest.fixture
def column_index(tmp_path):
    book = Workbook()
    book.active.title = "Big"
    for row in range(1, 301):
        book.active.cell(row, 1, row)
    book.active["B1"] = "=SUM(A1:A300)"
    path = tmp_path / "column.xlsx"
    book.save(path)
    build_index(path, tmp_path / "index.sqlite")
    return tmp_path / "index.sqlite"


def test_overlapping_selected_ranges_count_each_grid_cell_once(column_index):
    with WorkbookIndex(column_index) as store:
        store.prepare_input_selection()
        store.select_range_inputs("Big", (1, 1, 300, 1))
        store.select_range_inputs("Big", (100, 1, 400, 1))
        store.select_input("Big!A5")
        assert store.selected_input_count() == 300
        cells = list(store.selected_input_cells())
        assert [(c["row"], c["cached_comparison_value"]) for c in cells] == [
            (row, row) for row in range(1, 301)
        ]
        nodes = list(store.selected_inputs())
        assert [n["id"] for n in nodes[:3]] == ["Big!A1", "Big!A2", "Big!A3"]
        assert store.selected_input_sample(5) == [
            "Big!A1",
            "Big!A10",
            "Big!A100",
            "Big!A101",
            "Big!A102",
        ]
        assert list(store.selected_stored_inputs()) == []


def test_a_later_writer_turns_a_grid_cell_into_a_node(column_index):
    with WorkbookIndex(column_index, writable=True) as store:
        node = store["Big!A7"]
        node["diagnostics"] = ["Marked after the import."]
        store.put(node)
        store.db.commit()
    with WorkbookIndex(column_index) as store:
        assert store["Big!A7"]["diagnostics"] == ["Marked after the import."]
        assert len(store) == 302
        assert store.count_cells("Big", (1, 1, 300, 1)) == 300
        assert "Big!A7" not in _grid_ids(column_index)


def test_a_grid_date_displays_from_its_saved_serial(tmp_path):
    from linexcel.excel_presentation import sheet_presentation

    book = Workbook()
    book.active.title = "Dates"
    book.active["A1"] = dt.datetime(2024, 2, 29)
    book.active["A1"].number_format = "yyyy-mm-dd"
    book.active["B1"] = 45351.5
    book.active["B1"].number_format = "dd/mm/yyyy hh:mm"
    path = tmp_path / "dates.xlsx"
    book.save(path)
    build_index(path, tmp_path / "index.sqlite")
    assert {"Dates!A1", "Dates!B1"} <= _grid_ids(tmp_path / "index.sqlite")
    window = sheet_presentation(path, tmp_path / "index.sqlite", "Dates")
    cells = {cell["address"]: cell for cell in window["cells"]}
    assert (cells["A1"]["display"], cells["A1"]["formatStatus"]) == (
        "2024-02-29",
        "formatted",
    )
    assert cells["B1"]["display"] == "29/02/2024 12:00"
