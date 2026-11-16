"""What ``lazy_store``'s own reads of the index return for literal cells.

Index schema 5 moves literal input cells out of the ``nodes`` table into
per-sheet grids. ``lazy_store`` is the storage module, so its own queries stay
in it, but each query that can return, count or join an input cell must give
the same answer once inputs live elsewhere. Each test below pins one such
query through the public function that holds it, on the workbook of
``test_index_reads`` (341 inputs, 11 formulas, 4 ranges and a name), with
expected answers checked by hand.

The other reads of ``nodes`` in ``lazy_store`` never reach an input cell: they
filter on formulas, names, ranges or array results, or join an edge's target
only. The last test pins the two facts that last point relies on.
"""

import sqlite3

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from test_index_reads import reads_workbook

from linexcel.lazy_store import (
    WorkbookIndex,
    _unpack,
    build_index,
    evidence_graph,
    export_structure,
    graph_neighborhood,
    graph_page,
    structure_page,
)

_EDGES = {
    ("Arr!B1:B3", "Arr!A1"),
    ("Big!A1:A300", "Calc!B6"),
    ("Calc!A2", "Calc!B2"),
    ("Calc!B1", "Calc!B4"),
    ("Calc!B3", "name:*:Rate"),
    ("Data!A2:A6", "Calc!B2"),
    ("Data!B2", "Calc!B5"),
    ("Data!B2", "Data!D2"),
    ("Data!B2:B6", "Calc!B1"),
    ("Data!B2:B6", "Calc!B2"),
    ("Data!B3", "Data!D3"),
    ("Data!B4", "Data!D4"),
    ("Data!B5", "Data!D5"),
    ("Data!B6", "Data!D6"),
    ("name:*:Rate", "Calc!B4"),
}


@pytest.fixture(scope="module")
def index(tmp_path_factory):
    root = tmp_path_factory.mktemp("storage-reads")
    book = root / "reads.xlsx"
    reads_workbook().save(book)
    path = root / "index.sqlite"
    build_index(book, path)
    return path


def _ids(nodes):
    return [node["id"] for node in nodes]


# build_index: the node counts of the metadata count input cells.


def test_import_counts_every_node_kind(index):
    with WorkbookIndex(index) as store:
        meta = store.meta
    assert meta["nodeCounts"] == {"cell": 11, "input": 341, "name": 1, "range": 4}
    assert (meta["nodeCount"], meta["inputCount"], meta["formulaCount"]) == (
        357,
        341,
        11,
    )


def test_import_counts_every_node_kind_of_every_sheet(index):
    with WorkbookIndex(index) as store:
        meta = store.meta
    assert meta["perSheetCounts"] == {
        "defined names": {"name": 1},
        "Data": {"input": 19, "cell": 5, "range": 2},
        "Calc": {"input": 7, "cell": 5},
        "Big": {"input": 300, "range": 1},
        "Labels": {"input": 12},
        "Arr": {"cell": 1, "input": 3, "range": 1},
    }
    assert meta["sheetNodeCounts"] == {
        "defined names": 1,
        "Data": 26,
        "Calc": 12,
        "Big": 301,
        "Labels": 12,
        "Arr": 5,
    }


# _index_dependencies: a literal saved inside a declared array is an array
# member, owned by the array's anchor.


def test_import_makes_a_literal_inside_an_array_a_member_of_it(tmp_path):
    book = Workbook()
    book.active["A1"] = ArrayFormula(ref="A1:A3", text="=SEQUENCE(3,1)")
    book.active["A2"] = 999
    book.active["B1"] = 5
    path = tmp_path / "array.xlsx"
    book.save(path)
    build_index(path, tmp_path / "index.sqlite")
    with WorkbookIndex(tmp_path / "index.sqlite") as store:
        member, outside = store["Sheet!A2"], store["Sheet!B1"]
        anchors = [
            source
            for (source,) in store.db.execute(
                "SELECT source FROM edges WHERE target='Sheet!A2'"
            )
        ]
    assert member["kind"] == "array_result"
    assert member["arrayAnchor"] == "Sheet!A1"
    assert member["dependencies"] == ["Sheet!A1"]
    assert member["provenance"] == "saved_cache"
    assert member["value"] == 999
    assert anchors == ["Sheet!A1"]
    assert (outside["kind"], outside["value"]) == ("input", 5)


# graph_page: exact reference, count and listing.


def test_graph_page_finds_an_input_by_its_exact_reference(index):
    page = graph_page(index, query="Data!B2")
    assert _ids(page["nodes"]) == ["Data!B2"]
    assert page["nodes"][0]["value"] == 10
    assert page["pagination"]["total"] == 1
    sheet_page = graph_page(index, query="Data!B2", sheet="Data")
    assert _ids(sheet_page["nodes"]) == ["Data!B2"]


def test_graph_page_counts_inputs(index):
    assert graph_page(index)["pagination"]["total"] == 357
    assert graph_page(index, kind="input")["pagination"]["total"] == 341
    assert graph_page(index, sheet="Labels")["pagination"]["total"] == 12
    assert graph_page(index, sheet="Big", kind="input")["pagination"]["total"] == 300
    # "North" is a saved value of five text cells, none of them a formula.
    assert graph_page(index, query="North")["pagination"]["total"] == 5


def test_graph_page_lists_inputs_in_import_order(index):
    first = graph_page(index, limit=10)
    assert _ids(first["nodes"]) == [
        "name:*:Rate",
        "Data!A1",
        "Data!B1",
        "Data!C1",
        "Data!D1",
        "Data!A2",
        "Data!B2",
        "Data!C2",
        "Data!D2",
        "Data!A3",
    ]
    assert first["nodes"][1]["value"] == "Region"
    assert first["pagination"]["hasMore"] is True
    across = graph_page(index, offset=36, limit=4)
    assert _ids(across["nodes"]) == ["Calc!B6", "Big!A1", "Big!A2", "Big!A3"]
    found = graph_page(index, query="North")
    assert _ids(found["nodes"]) == [
        "Data!A2",
        "Data!A6",
        "Calc!A2",
        "Labels!A2",
        "Labels!A4",
    ]
    inputs = graph_page(index, kind="input", offset=338, limit=5)
    assert _ids(inputs["nodes"]) == ["Arr!B1", "Arr!B2", "Arr!B3"]


# structure_page: count, node page and edge page.


def test_structure_page_counts_inputs(index):
    assert structure_page(index)["pagination"]["total"] == 357


def test_structure_page_lists_inputs_in_import_order(index):
    first = structure_page(index, limit=5)["nodes"]
    assert first == [
        {"id": "name:*:Rate", "kind": "name", "sheet": None},
        {"id": "Data!A1", "kind": "input", "sheet": "Data"},
        {"id": "Data!B1", "kind": "input", "sheet": "Data"},
        {"id": "Data!C1", "kind": "input", "sheet": "Data"},
        {"id": "Data!D1", "kind": "input", "sheet": "Data"},
    ]
    last = structure_page(index, offset=350, limit=10)
    assert _ids(last["nodes"]) == [
        "Arr!B1",
        "Arr!B2",
        "Arr!B3",
        "Data!B2:B6",
        "Data!A2:A6",
        "Big!A1:A300",
        "Arr!B1:B3",
    ]
    assert last["pagination"]["hasMore"] is False


def test_structure_page_keeps_the_edges_an_input_feeds(index):
    page = structure_page(index, part="edges", limit=200)
    assert page["pagination"]["total"] == 15
    assert {(e["source"], e["target"]) for e in page["edges"]} == _EDGES
    assert {e["kind"] for e in page["edges"]} == {"dep"}
    assert page["edges"][2] == {"source": "Calc!A2", "target": "Calc!B2", "kind": "dep"}


# _edges_for_nodes: induced edges between selected nodes, input sources kept.


def test_evidence_of_a_formula_keeps_the_edge_from_its_input(index):
    graph = evidence_graph(index, "Data!D2")
    assert _ids(graph["nodes"]) == ["Data!D2", "Data!B2"]
    assert graph["edges"] == [{"source": "Data!B2", "target": "Data!D2", "kind": "dep"}]


# evidence_graph without a node: one node per sheet, then a workbook sample.


def test_evidence_sample_takes_an_input_from_a_sheet_without_formulas(index):
    ids = _ids(evidence_graph(index)["nodes"])
    # Formulas first in each sheet; Big and Labels hold only literals.
    assert ids[:5] == ["Data!D2", "Calc!B1", "Big!A1", "Labels!A1", "Arr!A1"]


def test_evidence_sample_completes_with_inputs_in_import_order(index):
    graph = evidence_graph(index)
    # The other formulas, then the name and the first inputs, up to 18.
    assert _ids(graph["nodes"])[5:] == [
        "Data!D3",
        "Data!D4",
        "Data!D5",
        "Data!D6",
        "Calc!B2",
        "Calc!B4",
        "Calc!B5",
        "Calc!B6",
        "name:*:Rate",
        "Data!A1",
        "Data!B1",
        "Data!C1",
        "Data!D1",
    ]
    assert graph["pagination"] == {"returned": 18, "total": 357, "truncated": True}


# export_structure: every node and every edge.


def test_export_lists_every_node_with_its_saved_value(index):
    nodes = export_structure(index)["nodes"]
    assert len(nodes) == 357
    assert _ids(nodes[:3]) == ["name:*:Rate", "Data!A1", "Data!B1"]
    by_id = {node["id"]: node for node in nodes}
    assert sum(node["kind"] == "input" for node in nodes) == 341
    assert by_id["Big!A300"]["value"] == 300
    assert by_id["Data!C3"]["value"] is False
    assert by_id["Data!C5"]["value"] == "#N/A"
    assert by_id["Data!C5"]["cachedValueKind"] == "error"


def test_export_keeps_the_edges_inputs_feed(index):
    edges = export_structure(index)["edges"]
    assert {(e["source"], e["target"]) for e in edges} == _EDGES


# Neighbourhood: range members, membership edges and an input's position.


def test_neighbourhood_of_a_range_reaches_its_input_members(index):
    graph = graph_neighborhood(index, "Data!B2:B6")
    assert _ids(graph["nodes"]) == [
        "Data!B2:B6",
        "Calc!B1",
        "Calc!B2",
        "Data!B2",
        "Data!B3",
        "Data!B4",
        "Data!B5",
        "Data!B6",
    ]


def test_neighbourhood_links_input_members_to_their_range(index):
    edges = graph_neighborhood(index, "Data!B2:B6")["edges"]
    members = [e["source"] for e in edges if e["kind"] == "range"]
    assert members == ["Data!B2", "Data!B3", "Data!B4", "Data!B5", "Data!B6"]
    assert {e["target"] for e in edges if e["kind"] == "range"} == {"Data!B2:B6"}


def test_neighbourhood_of_an_input_reaches_the_range_that_holds_it(index):
    graph = graph_neighborhood(index, "Data!B3")
    assert _ids(graph["nodes"]) == ["Data!B3", "Data!B2:B6", "Data!D3"]
    assert graph["edges"] == [
        {"source": "Data!B3", "target": "Data!B2:B6", "kind": "range"},
        {"source": "Data!B3", "target": "Data!D3", "kind": "dep"},
    ]


# The two reads a caller gets as stored records or as SQL.


def test_attached_pattern_sizes_count_inputs(index, tmp_path):
    with WorkbookIndex(index) as store, sqlite3.connect(tmp_path / "x.sqlite") as db:
        db.execute("ATTACH DATABASE ? AS original", (str(index),))
        sizes = {
            (sheet, pattern or None): size
            for sheet, pattern, size in db.execute(
                store.attached_pattern_sizes_sql("original")
            )
        }
    # Inputs, ranges and names have no formula pattern.
    unpatterned = {s: size for (s, pattern), size in sizes.items() if not pattern}
    assert unpatterned == {
        "": 1,
        "Data": 21,
        "Calc": 7,
        "Big": 301,
        "Labels": 12,
        "Arr": 4,
    }
    assert sum(sizes.values()) == 357


def test_window_records_hold_the_inputs_of_a_window(index):
    with WorkbookIndex(index) as store:
        nodes = [
            _unpack(record)
            for record in store.window_records("Data", (1, 1, 3, 3), 1 << 20)
        ]
    assert sorted((n["row"], n["column"], n["kind"], n["value"]) for n in nodes) == [
        (1, 1, "input", "Region"),
        (1, 2, "input", "Amount"),
        (1, 3, "input", "Flag"),
        (2, 1, "input", "North"),
        (2, 2, "input", 10),
        (2, 3, "input", True),
        (3, 1, "input", "South"),
        (3, 2, "input", 20),
        (3, 3, "input", False),
    ]


# What the other reads rely on.


def test_inputs_are_never_edge_targets_nor_formula_pattern_members(index):
    export = export_structure(index)
    kinds = {node["id"]: node["kind"] for node in export["nodes"]}
    assert not [e for e in export["edges"] if kinds[e["target"]] == "input"]
    assert sum(kinds[e["source"]] == "input" for e in export["edges"]) == 8
    assert not [
        node["id"]
        for node in export["nodes"]
        if node["kind"] == "input" and node.get("patternId")
    ]


def test_graph_page_search_matches_cell_references_and_sheet_names(index):
    # "a30": the ids Big!A30 and Big!A300, the range Big!A1:A300 and the
    # formula of Calc!B6, =SUM(Big!A1:A300).
    found = graph_page(index, query="a30")
    assert sorted(_ids(found["nodes"])) == [
        "Big!A1:A300",
        "Big!A30",
        "Big!A300",
        "Calc!B6",
    ]
    assert found["pagination"]["total"] == 4
    # A sheet name matches every node whose id holds it.
    assert graph_page(index, query="Labels!")["pagination"]["total"] == 12


# graph_page: a search reads each grid cell once, and builds no cell address
# a query cannot occur in (batch 2 ter: a search with no match read the
# 2.5 million grid cells of fpna twice, 10.4 s against 3.6 s with schema 4).


def _counting_scan(monkeypatch):
    from linexcel import cell_grid

    seen = {"cells": 0}
    scan = cell_grid.GridReader.scan

    def counted(self, sheet, bounds=None):
        for cell in scan(self, sheet, bounds):
            seen["cells"] += 1
            yield cell

    monkeypatch.setattr(cell_grid.GridReader, "scan", counted)
    return seen


def test_a_search_with_no_match_reads_each_grid_cell_once(index, monkeypatch):
    with WorkbookIndex(index) as store:
        grid_cells = sum(store.grid.count(name) for name in store.grid.sheets)
    assert grid_cells > 300
    seen = _counting_scan(monkeypatch)
    found = graph_page(index, query="zz no such text")
    assert found["pagination"]["total"] == 0 and found["nodes"] == []
    assert seen["cells"] == grid_cells


def test_a_search_builds_no_address_a_query_cannot_occur_in(tmp_path, monkeypatch):
    import re

    from linexcel import lazy_store

    book = Workbook()
    sheet = book.active
    sheet.title = "My Data"
    for row in range(1, 41):
        sheet.append([f"name {row}", row * 1.5])
    other = book.create_sheet("Plain")
    for row in range(1, 41):
        other.append([row, f"item {row}"])
    path = tmp_path / "book.xlsx"
    book.save(path)
    index = tmp_path / "index.sqlite"
    build_index(path, index)
    compared = []
    like_search = lazy_store._like_search

    def recorded(query):
        like = like_search(query)

        def test(text):
            compared.append(text)
            return like(text)

        return test

    monkeypatch.setattr(lazy_store, "_like_search", recorded)
    address = re.compile(r"!\$?[A-Z]+\$?[0-9]+$")
    # A space can be in a cell address only through a sheet name: every cell
    # of 'My Data' matches by its address, no address of Plain can match.
    assert graph_page(index, query="y da")["pagination"]["total"] == 80
    assert graph_page(index, query="item 3")["pagination"]["total"] == 11
    assert graph_page(index, query="zz no such text")["pagination"]["total"] == 0
    assert not [text for text in compared if address.search(text or "")]
    # Letters and digits can be an address: those queries still test it.
    assert graph_page(index, query="plain!b1")["pagination"]["total"] == 11
    assert [text for text in compared if address.search(text or "")]


def test_a_search_counts_every_match_past_a_full_page(index):
    # Big!A1:A300 holds 300 literal cells; the range node and the formula
    # =SUM(Big!A1:A300) of Calc!B6 also hold "big!". A page of 5 is full
    # long before the grid of Big ends: the total still counts all of them.
    with WorkbookIndex(index) as store:
        grid = store.grid.count("Big")
        others = store.db.execute(
            "SELECT count(*) FROM nodes WHERE id LIKE '%big!%' "
            "OR formula LIKE '%big!%' OR search_value LIKE '%big!%'"
        ).fetchone()[0]
    assert (grid, others) == (300, 2)
    for limit in (5, 200):
        found = graph_page(index, query="big!", limit=limit)
        assert found["pagination"]["total"] == grid + others
        assert len(found["nodes"]) == limit
