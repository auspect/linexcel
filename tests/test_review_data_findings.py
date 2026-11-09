"""Regressions from the independent data/lineage review, with explicit graphs."""

import io
import json
import socket
import zipfile
from copy import deepcopy
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from test_array_source_views import MAIN, array_bytes
from test_external import _link, alternate_urls_workbook
from test_total_lineage import fixture_source, graph_signature, handle, node_signature

from linexcel.catalog import Catalog
from linexcel.catalog_links import list_links
from linexcel.excel_presentation import sheet_presentation
from linexcel.external import read_external_links
from linexcel.lazy_store import WorkbookIndex


def rewrite(content, part, edit):
    output = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(content)) as source,
        zipfile.ZipFile(output, "w") as target,
    ):
        for name in source.namelist():
            data = source.read(name)
            if name == part:
                root = ET.fromstring(data)
                edit(root)
                data = ET.tostring(root)
            target.writestr(name, data)
    return output.getvalue()


@pytest.mark.parametrize("stored_blank", [True, False], ids=["empty-c", "absent-c"])
@pytest.mark.parametrize("route", ["root", "external", "local", "range"])
def test_empty_array_members_keep_their_anchor(tmp_path, stored_blank, route):
    content, _ = array_bytes(blank=True)

    def edit(root):
        row = root.find("{*}sheetData/{*}row[@r='3']")
        if not stored_blank:
            row.remove(row.find("{*}c[@r='K3']"))
        if route == "local":
            cell = ET.SubElement(row, f"{{{MAIN}}}c", r="M3")
            ET.SubElement(cell, f"{{{MAIN}}}f").text = "K3"

    content = rewrite(content, "xl/worksheets/sheet1.xml", edit)
    target = fixture_source(
        tmp_path, 1, "arrays.xlsx", {}, sheet="Feuil1", content=content
    )
    sources = [target]
    if route == "external":
        reader = fixture_source(
            tmp_path,
            2,
            "reader.xlsx",
            {"A1": "='[1]Feuil1'!K3"},
            links=[_link("arrays.xlsx", "Feuil1")],
        )
        sources.append(reader)
        selected = handle(reader)
    else:
        selected = handle(
            target,
            "M3" if route == "local" else "K3:L3" if route == "range" else "K3",
            "Feuil1",
        )
    result = Catalog(tmp_path / "catalog").lineage(
        sources, selected, direction="upstream", depth=10
    )
    inputs = [
        node_signature(target, c, "Feuil1")
        for c in ["A2", "B2", "A3", "B3", "A4", "B4"]
    ]
    span, anchor, blank = [
        node_signature(target, c, "Feuil1")
        for c in ["A2:B4", "J2", "K3:L3" if route == "range" else "K3"]
    ]
    nodes = [*inputs, span, anchor, blank]
    edges = [(i, span, "range") for i in inputs] + [
        (span, anchor, "formula"),
        (anchor, blank, "formula"),
    ]
    if route in {"external", "local"}:
        consumer = (
            node_signature(reader, "A1")
            if route == "external"
            else node_signature(target, "M3", "Feuil1")
        )
        nodes.append(consumer)
        edges.append(
            (blank, consumer, "external_link" if route == "external" else "formula")
        )
    assert graph_signature(result) == {
        "nodes": sorted(nodes),
        "edges": sorted(edges),
        "boundaries": sorted((i, "input_constant") for i in inputs),
    }
    assert result["coverage"]["complete"]
    if route != "range":
        member = next(n for n in result["nodes"] if n["address"] == "K3")
        assert member["kind"] == "array_result" and member["value"] is None
    view = sheet_presentation(
        target.path.parent / "workbook.xlsx",
        target.path,
        "Feuil1",
        row=3,
        column=11,
        rows=1,
        columns=1,
    )
    if view["cells"]:
        assert view["cells"][0]["value"] is None and view["cells"][0]["display"] == ""
        assert view["cells"][0]["valueStatus"] == "blank"


@pytest.mark.parametrize("depth", [0, 1])
def test_implicit_member_of_huge_array_is_bounded(tmp_path, depth):
    book = Workbook()
    book.active["A1"] = ArrayFormula(ref="A1:XFD1048576", text="=1")
    content = io.BytesIO()
    book.save(content)
    target = fixture_source(
        tmp_path, 1, "huge.xlsx", {}, sheet="Sheet", content=content.getvalue()
    )
    with WorkbookIndex(target.path) as index:
        assert len(index) == 1
    result = Catalog(tmp_path / "catalog").lineage(
        [target],
        handle(target, "XFD1048576", "Sheet"),
        direction="upstream",
        depth=depth,
        limit=2,
    )
    assert {n["address"] for n in result["nodes"]} == (
        {"XFD1048576", "A1"} if depth else {"XFD1048576"}
    )
    assert result["coverage"]["complete"] is bool(depth)
    assert result["coverage"]["boundaries"] == ([] if depth else ["depth"])
    with WorkbookIndex(target.path) as index:
        assert len(index) == 1  # Read-time resolution never expands the rectangle.


def test_array_anchor_downstream_reaches_external_reader_of_empty_member(tmp_path):
    content, _ = array_bytes(blank=True)
    content = rewrite(
        content,
        "xl/workbook.xml",
        lambda root: root.remove(root.find("{*}definedNames")),
    )
    target = fixture_source(tmp_path, 1, "arrays.xlsx", {}, content=content)
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": "='[1]Feuil1'!K3"},
        links=[_link("arrays.xlsx", "Feuil1")],
    )
    result = Catalog(tmp_path / "catalog").lineage(
        [target, reader],
        handle(target, "J2", "Feuil1"),
        direction="downstream",
        depth=5,
    )
    anchor = node_signature(target, "J2", "Feuil1")
    members = [
        node_signature(target, c, "Feuil1") for c in ["K2", "J3", "K3", "J4", "K4"]
    ]
    blank = node_signature(target, "K3", "Feuil1")
    consumer = node_signature(reader, "A1")
    assert graph_signature(result) == {
        "nodes": sorted([anchor, *members, consumer]),
        "edges": sorted(
            [
                *((anchor, m, "formula") for m in members),
                (blank, consumer, "external_link"),
            ]
        ),
        "boundaries": [],
    }
    assert result["coverage"]["complete"]


def test_huge_implicit_array_downstream_never_claims_complete(tmp_path):
    book = Workbook()
    book.active["A1"] = ArrayFormula(ref="A1:XFD1048576", text="=1")
    content = io.BytesIO()
    book.save(content)
    target = fixture_source(
        tmp_path, 1, "huge.xlsx", {}, sheet="Sheet", content=content.getvalue()
    )
    result = Catalog(tmp_path / "catalog").lineage(
        [target],
        handle(target, "A1", "Sheet"),
        direction="downstream",
        depth=2,
        limit=2,
    )
    assert not result["coverage"]["complete"]
    assert result["coverage"]["capsHit"] == ["node_budget"]
    assert len(result["nodes"]) <= 2


@pytest.mark.parametrize(
    "target_path,filename",
    [
        ("https://tenant.sharepoint.com/sites/x/budget.xlsx?d=abc", "budget.xlsx"),
        ("https://tenant.sharepoint.com/sites/x/budget.xlsx#section", "budget.xlsx"),
        (
            "https://d.docs.live.net/abc/Documents/budget.xlsx?download=1#top",
            "budget.xlsx",
        ),
        (
            "HTTP://example.test/Ventes%20Région%20été.xlsx?x=1",
            "Ventes Région été.xlsx",
        ),
        ("https://example.test/budget%3Ffinal.xlsx?d=abc#top", "budget?final.xlsx"),
        ("https://example.test/budget%23final.xlsx?d=abc#top", "budget#final.xlsx"),
        ("/local/budget.xlsx?literal", "budget.xlsx?literal"),
        (r"C:\local\budget.xlsx?literal", "budget.xlsx?literal"),
        ("file:///C:/local/budget.xlsx?literal", "budget.xlsx?literal"),
    ],
)
def test_cloud_queries_use_only_url_path_without_changing_local_names(
    tmp_path, monkeypatch, target_path, filename
):
    def forbidden(*args, **kwargs):
        pytest.fail("Relationship targets must never be fetched")

    monkeypatch.setattr(socket, "create_connection", forbidden)
    monkeypatch.setattr(socket.socket, "connect", forbidden)
    content = alternate_urls_workbook(target_path, filename)
    entry = read_external_links(content)["1"]
    assert entry.name == filename and entry.relationship_path == target_path
    target = fixture_source(tmp_path, 1, filename, {"B2": 10})
    reader = fixture_source(tmp_path, 2, "reader.xlsx", {}, content=content)
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.lineage(
        [target, reader], handle(reader), direction="upstream", depth=5
    )
    fact = list_links(catalog, [target, reader])["links"][0]
    assert (
        fact["status"] == "resolved_by_name" and fact["relationshipPath"] == target_path
    )
    assert result["coverage"]["complete"] and len(result["edges"]) == 1
    assert result["edges"][0]["relationshipPath"] == target_path


@pytest.mark.parametrize(
    "variation", ["target", "type", "mode", "restored", "identical"]
)
def test_conflicting_relationship_ids_have_unknown_identity(tmp_path, variation):
    content = alternate_urls_workbook("/dir/budget.xlsx", "budget.xlsx")

    def edit(root):
        original = next(r for r in root if r.get("Id") == "rId1")
        duplicate = deepcopy(original)
        if variation in {"target", "restored"}:
            duplicate.set("Target", "/dir/other.xlsx")
        elif variation == "type":
            duplicate.set("Type", "http://example.test/other")
        elif variation == "mode":
            duplicate.set("TargetMode", "Internal")
        root.append(duplicate)
        if variation == "restored":
            root.append(deepcopy(original))

    content = rewrite(content, "xl/externalLinks/_rels/externalLink1.xml.rels", edit)
    entry = read_external_links(content)["1"]
    valid = variation == "identical"
    assert entry.name == ("budget.xlsx" if valid else "")
    assert entry.relationship_path == ("/dir/budget.xlsx" if valid else "")
    targets = [
        fixture_source(tmp_path, n, name, {"B2": 10})
        for n, name in [(1, "budget.xlsx"), (2, "other.xlsx")]
    ]
    reader = fixture_source(tmp_path, 3, "reader.xlsx", {}, content=content)
    sources = [*targets, reader]
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.lineage(sources, handle(reader), direction="upstream", depth=5)
    fact = list_links(catalog, sources)["links"][0]
    assert fact["status"] == ("resolved_by_name" if valid else "unresolved")
    if valid:
        assert len(result["edges"]) == 1 and result["coverage"]["complete"]
    else:
        assert fact["candidates"] == [] and fact["targetRevision"] is None
        assert not result["edges"] and not result["coverage"]["complete"]
        assert set(result["coverage"]["boundaries"]) == {
            "unresolved_external",
            "cache_only",
        }


def test_old_external_identity_facts_require_rebuild_instead_of_a_wrong_edge(tmp_path):
    target = fixture_source(tmp_path, 1, "budget.xlsx", {"B2": 10})
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": "='[1]S'!B2"},
        links=[_link("budget.xlsx", "S")],
    )
    with WorkbookIndex(reader.path, writable=True) as index:
        metadata = index.meta
        metadata["externalFactsVersion"] = 1
        index.db.execute(
            "UPDATE metadata SET value=? WHERE key='meta'", (json.dumps(metadata),)
        )
        index.db.commit()
    catalog = Catalog(tmp_path / "catalog")
    result = catalog.lineage(
        [target, reader], handle(reader), direction="upstream", depth=5
    )
    assert "index_rebuild_required" in result["coverage"]["capsHit"]
    assert not result["coverage"]["complete"]
    assert not any(e["kind"] == "external_link" for e in result["edges"])
