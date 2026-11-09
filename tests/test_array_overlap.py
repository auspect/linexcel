"""Malformed overlapping array declarations retain values, never a chosen owner."""

import io
import zipfile
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from test_total_lineage import fixture_source, handle

from linexcel.catalog import Catalog
from linexcel.excel_presentation import sheet_presentation
from linexcel.graph_store import build_graph
from linexcel.lazy_store import (
    WorkbookIndex,
    _pack,
    export_structure,
    graph_neighborhood,
    graph_page,
    structure_page,
)
from linexcel.web_data import build_source_database


def overlapping_source(tmp_path, second="B1"):
    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    sheet["E1"], sheet["F1"] = 11, 22
    sheet["A1"] = ArrayFormula(ref="A1:B2", text="=E1")
    sheet[second] = ArrayFormula(
        ref={"B1": "B1:C3", "B2": "B2:C3", "C1": "C1:D3"}[second], text="=F1"
    )
    if second != "B2":
        sheet["B2"] = 7
    source, output = io.BytesIO(), io.BytesIO()
    book.save(source)
    with zipfile.ZipFile(source) as archive, zipfile.ZipFile(output, "w") as result:
        for name in archive.namelist():
            data = archive.read(name)
            if name == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(data)
                for cell in root.findall("{*}sheetData/{*}row/{*}c"):
                    if cell.get("r") in {"A1", second, "B2"}:
                        value = cell.find("{*}v")
                        if value is None:
                            value = ET.SubElement(
                                cell,
                                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v",
                            )
                        value.text = (
                            "7"
                            if cell.get("r") == "B2"
                            else "11"
                            if cell.get("r") == "A1"
                            else "22"
                        )
                data = ET.tostring(root)
            result.writestr(name, data)
    return fixture_source(tmp_path, 1, "overlap.xlsx", {}, content=output.getvalue())


@pytest.mark.parametrize("second", ["B1", "B2"], ids=["stored-member", "stored-anchor"])
def test_stored_overlapping_member_is_unresolved_but_value_is_visible(tmp_path, second):
    source = overlapping_source(tmp_path, second)
    with WorkbookIndex(source.path) as index:
        node = index.node("S!B2")
        assert node["value"] == node["cachedValue"] == 7
        assert node["kind"] == "unresolved"
        assert not node.get("arrayAnchor") and not node["dependencies"]
        assert not node["sourceValueProvenance"].get("arrayAnchor")
    result = Catalog(tmp_path / "catalog").lineage(
        [source], handle(source, "B2"), direction="upstream", depth=5
    )
    assert not result["edges"] and len(result["nodes"]) == 1
    assert result["nodes"][0]["value"] == 7
    assert result["coverage"]["boundaries"] == ["unsupported_reference"]
    assert not result["coverage"]["complete"]
    view = sheet_presentation(
        source.path.parent / "workbook.xlsx", source.path, "S", rows=3, columns=6
    )
    cell = next(c for c in view["cells"] if c["address"] == "B2")
    assert (cell["value"], cell["display"], cell["valueStatus"]) == (7, "7", "present")
    assert not cell["arrayAnchor"]


@pytest.mark.parametrize("second", ["B1", "B2"])
@pytest.mark.parametrize("owner", ["first", "second"])
def test_downstream_of_either_overlap_anchor_never_chooses_an_owner(
    tmp_path, second, owner
):
    source = overlapping_source(tmp_path, second)
    root = "A1" if owner == "first" else second
    result = Catalog(tmp_path / "catalog").lineage(
        [source], handle(source, root), direction="downstream", depth=5
    )
    ids = {n["id"]: n["address"] for n in result["nodes"]}
    assert not any(ids[e["target"]] == "B2" for e in result["edges"])
    assert "unsupported_reference" in result["coverage"]["boundaries"]
    assert not result["coverage"]["complete"]


def test_nonoverlapping_array_ownership_and_downstream_are_unchanged(tmp_path):
    source = overlapping_source(tmp_path, "C1")
    with WorkbookIndex(source.path) as index:
        node = index.node("S!B2")
        assert node["kind"] == "array_result" and node["arrayAnchor"] == "S!A1"
        assert node["value"] == 7
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.lineage(
        [source], handle(source, "B2"), direction="upstream", depth=5
    )
    assert {n["address"] for n in result["nodes"]} == {"B2", "A1", "E1"}
    assert result["coverage"]["complete"]
    for anchor in ["A1", "C1"]:
        result = catalog.lineage(
            [source], handle(source, anchor), direction="downstream", depth=5
        )
        assert result["coverage"]["complete"]


@pytest.mark.parametrize(
    "graph_view",
    [
        "page",
        "owner_neighborhood",
        "member_neighborhood",
        "structure",
        "export",
        "hierarchy",
    ],
)
def test_old_stored_first_owner_is_not_reused_or_exported(tmp_path, graph_view):
    source = overlapping_source(tmp_path)
    with WorkbookIndex(source.path, writable=True) as index:
        node = index["S!B2"]
        node.update(kind="array_result", arrayAnchor="S!A1", dependencies=["S!A1"])
        node.pop("arrayOwnershipUnknown", None)
        node["sourceValueProvenance"]["arrayAnchor"] = "S!A1"
        index.db.execute(
            "UPDATE nodes SET kind='array_result',data=? WHERE id=?",
            (_pack(node), node["id"]),
        )
        index.db.execute("INSERT OR IGNORE INTO edges VALUES ('S!A1','S!B2')")
        index.db.commit()
        raw_edges = index.db.execute("SELECT count(*) FROM edges").fetchone()[0]
    original = source.path.read_bytes()
    if graph_view == "hierarchy":
        meta = build_graph(source.path, tmp_path / "hierarchy.sqlite")
        assert meta["edgeCount"] == raw_edges - 1
    else:
        if graph_view == "page":
            graph = graph_page(source.path)
        elif graph_view.endswith("neighborhood"):
            root = "S!A1" if graph_view.startswith("owner") else "S!B2"
            graph = graph_neighborhood(source.path, root, depth=3, limit=3)
            assert not graph["pagination"]["truncated"]
            assert ("S!B2" if root == "S!A1" else "S!A1") not in {
                node["id"] for node in graph["nodes"]
            }
        elif graph_view == "structure":
            edges = []
            while True:
                graph = structure_page(
                    source.path, part="edges", offset=len(edges), limit=1
                )
                assert graph["pagination"]["total"] == raw_edges - 1
                edges.extend(graph["edges"])
                if not graph["pagination"]["hasMore"]:
                    break
                assert graph["edges"]
            graph["edges"] = edges
        else:
            graph = export_structure(source.path)
        assert not any(
            edge["source"] == "S!A1" and edge["target"] == "S!B2"
            for edge in graph["edges"]
        )
    assert source.path.read_bytes() == original
    catalog = Catalog(tmp_path / "catalog")
    for direction, address in [("upstream", "B2"), ("downstream", "A1")]:
        result = catalog.lineage(
            [source], handle(source, address), direction=direction, depth=5
        )
        assert not result["coverage"]["complete"]
        assert "unsupported_reference" in result["coverage"]["boundaries"]
        assert not any(
            e["source"] == handle(source, "A1") and e["target"] == handle(source, "B2")
            for e in result["edges"]
        )
    view = sheet_presentation(
        source.path.parent / "workbook.xlsx", source.path, "S", rows=3, columns=6
    )
    cell = next(c for c in view["cells"] if c["address"] == "B2")
    assert cell["value"] == 7 and not cell["arrayAnchor"]
    output = tmp_path / "source-data.sqlite"
    build_source_database(source.path, output, max_bytes=4 * 1024 * 1024)
    import json
    import sqlite3
    from contextlib import closing

    with closing(sqlite3.connect(output)) as db:
        value, provenance = db.execute(
            "SELECT value,value_provenance FROM cells WHERE cell='B2'"
        ).fetchone()
        assert value == 7 and not json.loads(provenance).get("arrayAnchor")


def test_large_fully_stored_array_reports_overlap_beyond_enumeration_ceiling(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    sheet["A1"] = ArrayFormula(ref="A1:A201", text="=1")
    sheet["A101"] = ArrayFormula(ref="A101:B201", text="=2")
    for row in range(2, 202):
        if row != 101:
            sheet.cell(row, 1, row)
    content = io.BytesIO()
    book.save(content)
    source = fixture_source(
        tmp_path, 1, "large-overlap.xlsx", {}, content=content.getvalue()
    )
    result = Catalog(tmp_path / "catalog").lineage(
        [source], handle(source, "A1"), direction="downstream", depth=5, limit=200
    )
    assert "unsupported_reference" in result["coverage"]["boundaries"]
    assert not result["coverage"]["complete"]


@pytest.mark.parametrize("size", [10, 2000])
def test_normal_array_graph_reads_stay_bounded_without_decoding_structure(
    tmp_path, monkeypatch, size
):
    from linexcel import lazy_store

    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    sheet["A1"] = ArrayFormula(ref=f"A1:A{size}", text="=1")
    for row in range(2, size + 1):
        sheet.cell(row, 1, row)
    content = io.BytesIO()
    book.save(content)
    source = fixture_source(tmp_path, 1, "normal.xlsx", {}, content=content.getvalue())
    decoded = []
    unpack = lazy_store._unpack

    def track(data):
        decoded.append(None)
        return unpack(data)

    monkeypatch.setattr(lazy_store, "_unpack", track)
    page = structure_page(source.path, part="edges", limit=2)
    assert len(page["edges"]) == 2
    assert page["pagination"] == {"offset": 0, "total": size - 1, "hasMore": True}
    meta = build_graph(source.path, tmp_path / "normal-hierarchy.sqlite")
    assert meta["edgeCount"] == size - 1
    assert not decoded  # Complete counts and structure never decode cell blobs.
    around = graph_neighborhood(source.path, "S!A1", depth=3, limit=3)
    assert len(around["nodes"]) == 3
    assert len(decoded) <= 30  # Includes the bounded nearby-label context reads.
    assert around["pagination"]["truncated"]
    assert len(around["edges"]) == 2
    assert all(not n.get("arrayOwnershipUnknown") for n in around["nodes"])
