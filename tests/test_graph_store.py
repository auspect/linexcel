"""Whole-workbook coverage remains exact in every bounded hierarchy view."""

import json
import os
import sqlite3
from contextlib import closing
from pathlib import Path

import pytest
from openpyxl import Workbook

from linexcel.graph_store import build_graph, generation, read_view
from linexcel.lazy_store import build_index


def test_hierarchy_conserves_topology_and_can_locate_every_cell(tmp_path):
    book = Workbook()
    book.active.title = "Data"
    for i in range(1, 1100):
        book.active.cell(i, 1, i)
        book.active.cell(i, 2, f"=A{i}+A1")
    for i in range(34):
        sheet = book.create_sheet(f"Sheet{i}")
        sheet["A1"] = "=Data!B1099"
    book.active["C1"] = "=C2"
    book.active["C2"] = "=C1"
    book.active["C3"] = "=SUM(A1:A1099)"
    source, index, artifact = [
        tmp_path / x for x in ("book.xlsx", "index.sqlite", "hierarchy.sqlite")
    ]
    book.save(source)
    build_index(source, index)
    original = index.read_bytes()
    meta = build_graph(index, artifact)
    assert index.read_bytes() == original
    with closing(sqlite3.connect(index)) as db:
        assert (
            meta["edgeCount"] == db.execute("SELECT count(*) FROM edges").fetchone()[0]
        )
        all_ids = {row[0] for row in db.execute("SELECT id FROM nodes")}
    seen, pending, visited = set(), [0], set()
    while pending:
        view_id = pending.pop()
        assert view_id not in visited
        visited.add(view_id)
        view = read_view(artifact, generation(index), view=view_id)
        assert len(view["nodes"]) <= 33
        assert len(view["edges"]) <= 33 * 32
        assert sum(n["nodeCount"] for n in view["nodes"]) == len(all_ids)
        assert (
            sum(n["internalEdges"] for n in view["nodes"])
            + sum(e["weight"] for e in view["edges"])
        ) == meta["edgeCount"]
        for node in view["nodes"]:
            if node["kind"] == "group":
                assert node["label"].count(" … ") <= 1
                pending.append(node["viewId"])
            elif node["kind"] != "outside":
                seen.add(node["nodeId"])
    assert seen == all_ids
    view = read_view(artifact, generation(index), node_id="Data!B1099")
    assert any(n.get("nodeId") == "Data!B1099" for n in view["nodes"])
    with pytest.raises(ValueError, match="stale"):
        read_view(artifact, "different index")
    with pytest.raises(KeyError):
        read_view(artifact, generation(index), node_id="missing")


def test_hierarchy_long_paths_close_handles_and_empty_workbook(tmp_path):
    root = tmp_path / ("long-directory-" * 8) / ("deep-directory-" * 8)
    if os.name == "nt":
        root = Path("\\\\?\\" + str(root))
    root.mkdir(parents=True)
    book, index, artifact = [
        root / name for name in ("book.xlsx", "index.sqlite", "hierarchy.sqlite")
    ]
    Workbook().save(book)
    build_index(book, index)
    build_graph(index, artifact)
    assert read_view(artifact, generation(index))["nodes"] == []
    # Windows refuses these replacements if any SQLite connection is left open.
    artifact.rename(root / "old.sqlite")
    build_graph(index, artifact)
    assert read_view(artifact, generation(index))["meta"]["nodeCount"] == 0


def test_sheet_overview_weights_and_r1c1_drilldown(tmp_path):
    book = Workbook()
    source_sheet = book.active
    source_sheet.title = "Source"
    target = book.create_sheet("Result")
    for row in range(1, 81):
        source_sheet.cell(row, 1, row)
        target.cell(row, 1, f"=Source!A{row}*2")
    source, index, artifact = [
        tmp_path / name for name in ("book.xlsx", "index.sqlite", "hierarchy.sqlite")
    ]
    book.save(source)
    build_index(source, index)
    build_graph(index, artifact)
    stamp = generation(index)
    overview = read_view(artifact, stamp)
    sheets = {n["label"]: n for n in overview["nodes"]}
    assert all(n["level"] == "sheet" for n in sheets.values())
    assert overview["edges"] == [
        {
            "source": sheets["Source"]["id"],
            "target": sheets["Result"]["id"],
            "weight": 80,
        }
    ]
    result = read_view(artifact, stamp, view=sheets["Result"]["viewId"])
    assert result == read_view(artifact, stamp, sheet="Result")
    patterns = [n for n in result["nodes"] if n["kind"] == "group"]
    assert len(patterns) == 1
    assert patterns[0]["level"] == "pattern"
    assert patterns[0]["nodeCount"] == 80
    assert patterns[0]["signature"]
    leaf = read_view(artifact, stamp, node_id="Result!A80")
    assert any(n.get("nodeId") == "Result!A80" for n in leaf["nodes"])
    assert len(leaf["nodes"]) <= 33
    assert len(json.dumps(overview)) < 20_000


def test_small_pattern_opens_cells_without_a_single_partition(tmp_path):
    book = Workbook()
    for row in range(1, 21):
        book.active.cell(row, 1, row)
        book.active.cell(row, 2, f"=A{row}*2")
    source, index, artifact = [
        tmp_path / name for name in ("book.xlsx", "index.sqlite", "hierarchy.sqlite")
    ]
    book.save(source)
    build_index(source, index)
    build_graph(index, artifact)
    stamp = generation(index)
    sheet = read_view(artifact, stamp)["nodes"][0]
    groups = read_view(artifact, stamp, view=sheet["viewId"])["nodes"]
    pattern = next(n for n in groups if n.get("level") == "pattern")
    leaf = read_view(artifact, stamp, view=pattern["viewId"])
    assert len([n for n in leaf["nodes"] if n.get("nodeId")]) == 20
    assert not any(n["kind"] == "group" for n in leaf["nodes"])


def test_mixed_pattern_partitions_have_true_bounds_and_bounded_cell_previews(tmp_path):
    from openpyxl.utils.cell import coordinate_to_tuple

    book = Workbook()
    sheet = book.active
    for row in range(1, 75):
        sheet.cell(row, 1, row)
        sheet.cell(row, 4, "=SUM($A$1:$A$74)")
        sheet.cell(row, 16, f"={row}")
    sheet["B1"] = "x" * 1000
    source, index, artifact = [
        tmp_path / name for name in ("book.xlsx", "index.sqlite", "hierarchy.sqlite")
    ]
    book.save(source)
    build_index(source, index)
    build_graph(index, artifact)
    view = read_view(artifact, generation(index), node_id="Sheet!D50", index_path=index)
    cell = next(n for n in view["nodes"] if n.get("nodeId") == "Sheet!D50")
    assert cell["formula"] == "=SUM($A$1:$A$74)"
    for crumb in view["breadcrumbs"]:
        if crumb["level"] == "partition":
            start, end = map(coordinate_to_tuple, crumb["label"].split(" … "))
            assert start[0] <= 50 <= end[0]
            assert start[1] <= 4 <= end[1]
    text = read_view(artifact, generation(index), node_id="Sheet!B1", index_path=index)
    cell = next(n for n in text["nodes"] if n.get("nodeId") == "Sheet!B1")
    assert cell["value"] == "x" * 256
    assert cell["contentTruncated"]
    with closing(sqlite3.connect(artifact)) as db:
        assert not db.execute(
            "SELECT g.id FROM groups g JOIN groups c ON c.parent=g.id "
            "WHERE g.level='partition' GROUP BY g.id HAVING count(*)=1"
        ).fetchall()
