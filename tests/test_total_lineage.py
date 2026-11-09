"""OOXML-generated (not Excel-authored) external lineage contract fixtures."""

import hashlib
import io
import json
import sqlite3
import threading
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import replace
from xml.etree import ElementTree as ET

import pytest
from test_external import EXCEL_TARGETS, _link, alternate_urls_workbook, linked_workbook

from linexcel.catalog import Catalog, CatalogLimits, Source, cell_handle, revision
from linexcel.catalog_links import list_links, set_pin
from linexcel.lazy_store import WorkbookIndex, build_index


def fixture_source(
    tmp_path, number, filename, cells, *, links=(), sheet="S", caches=None, content=None
):
    directory = tmp_path / str(number)
    directory.mkdir(exist_ok=True)
    content = content if content is not None else linked_workbook(cells, links, sheet)
    if caches:
        output = io.BytesIO()
        with (
            zipfile.ZipFile(io.BytesIO(content)) as archive,
            zipfile.ZipFile(output, "w") as target,
        ):
            for name in archive.namelist():
                data = archive.read(name)
                if name == "xl/worksheets/sheet1.xml":
                    root = ET.fromstring(data)
                    for cell in root.findall("{*}sheetData/{*}row/{*}c"):
                        if cell.get("r") in caches:
                            value = cell.find("{*}v")
                            if value is None:
                                value = ET.SubElement(
                                    cell,
                                    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v",
                                )
                            value.text = str(caches[cell.get("r")])
                    data = ET.tostring(root)
                target.writestr(name, data)
        content = output.getvalue()
    path = directory / "workbook.xlsx"
    path.write_bytes(content)
    index = directory / "index.sqlite"
    build_index(path, index, filename=filename)
    return Source(f"{number:032x}", filename, index)


def handle(source, address="A1", sheet="S"):
    from linexcel.refs import quote_sheet

    with WorkbookIndex(source.path) as index:
        return cell_handle(
            revision(source.project_id, index.meta["sha256"]),
            f"{quote_sheet(sheet)}!{address}",
        )


def graph_signature(result):
    nodes = {
        n["id"]: (
            n["filename"],
            n["sheet"],
            n["address"],
            n["revision"]["projectId"],
            n["revision"]["sha256"],
        )
        for n in result["nodes"]
    }
    return {
        "nodes": sorted(nodes.values()),
        "edges": sorted(
            (nodes[e["source"]], nodes[e["target"]], e["kind"]) for e in result["edges"]
        ),
        "boundaries": sorted(
            (nodes[b["node"]], b["reason"]) for b in result["boundaries"]
        ),
    }


def node_signature(source, address, sheet="S"):
    with WorkbookIndex(source.path) as index:
        return source.filename, sheet, address, source.project_id, index.meta["sha256"]


def chain(tmp_path):
    inputs = fixture_source(tmp_path, 1, "inputs.xlsx", {"A1": 10, "A2": 20})
    regional = fixture_source(
        tmp_path,
        2,
        "regional.xlsx",
        {"B1": "=SUM('[1]S'!A1:A2)"},
        links=[_link("../inputs.xlsx", "S", {"A1": 10, "A2": 20})],
        caches={"B1": 30},
    )
    consolidated = fixture_source(
        tmp_path,
        3,
        "consolidated.xlsx",
        {"C1": "='[1]S'!B1*2"},
        links=[_link("refs/regional.xlsx", "S", {"B1": 30})],
        caches={"C1": 60},
    )
    return [inputs, regional, consolidated]


@pytest.mark.parametrize("target_path,filename", EXCEL_TARGETS)
def test_alternate_url_targets_resolve_by_basename_without_fetching(
    tmp_path, monkeypatch, target_path, filename
):
    import socket

    def no_network(*args, **kwargs):
        pytest.fail("External relationship targets must never be fetched")

    monkeypatch.setattr(socket, "create_connection", no_network)
    monkeypatch.setattr(socket.socket, "connect", no_network)
    target = fixture_source(tmp_path, 1, filename, {"B2": 10})
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {},
        content=alternate_urls_workbook(target_path, filename),
    )
    sources = [target, reader]
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(sources, handle(reader), direction="upstream", depth=10)
    b2, a1 = node_signature(target, "B2"), node_signature(reader, "A1")
    assert graph_signature(result) == {
        "nodes": sorted([b2, a1]),
        "edges": [(b2, a1, "external_link")],
        "boundaries": [(b2, "input_constant")],
    }
    assert result["coverage"]["complete"]
    assert result["edges"][0]["relationshipPath"] == target_path
    fact = list_links(catalog, sources)["links"][0]
    assert fact["status"] == "resolved_by_name"
    assert fact["filename"] == filename
    assert fact["relationshipPath"] == target_path


def test_exact_three_workbook_graph_and_value(tmp_path):
    sources = chain(tmp_path)
    i1 = node_signature(sources[0], "A1")
    i2 = node_signature(sources[0], "A2")
    ir = node_signature(sources[0], "A1:A2")
    r1 = node_signature(sources[1], "B1")
    c1 = node_signature(sources[2], "C1")
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        sources, handle(sources[2], "C1"), direction="upstream", depth=10
    )
    assert graph_signature(result) == {
        "nodes": sorted([c1, i1, ir, i2, r1]),
        "edges": [
            (i1, ir, "range"),
            (ir, r1, "external_link"),
            (i2, ir, "range"),
            (r1, c1, "external_link"),
        ],
        "boundaries": [
            (i1, "input_constant"),
            (i2, "input_constant"),
        ],
    }
    assert result["coverage"]["complete"] and result["coverage"]["exhaustive"]
    assert len(result["revisions"]) == 3
    values = {(n["filename"], n["address"]): n["value"] for n in result["nodes"]}
    assert (
        (values[("inputs.xlsx", "A1")] + values[("inputs.xlsx", "A2")]) * 2
        == values[("consolidated.xlsx", "C1")]
        == 60
    )
    assert result["nodes"][0]["valueProvenance"] == "saved_cache"
    downstream = catalog.lineage(
        sources, handle(sources[0], "A2"), direction="downstream", depth=10
    )
    assert graph_signature(downstream) == {
        "nodes": sorted([c1, i2, r1]),
        "edges": [
            (i2, r1, "external_link"),
            (r1, c1, "external_link"),
        ],
        "boundaries": [],
    }
    assert downstream["coverage"]["complete"]
    assert catalog.status()["links"]["counts"]["resolved_by_name"] == 2


@pytest.mark.parametrize(
    "target",
    [
        "../refs/INPUTS.xlsx",
        "file:///C:/Secret%20Folder/inputs.xlsx",
        r"\\server\share\inputs.xlsx",
    ],
)
def test_paths_names_and_diagnostic_only(tmp_path, target):
    inputs = fixture_source(tmp_path, 1, "inputs.xlsx", {"A1": 600})
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": "='[1]S'!A1"},
        links=[_link(target, "S", {"A1": 500})],
    )
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        [inputs, reader], handle(reader), direction="upstream", depth=4
    )
    edge = result["edges"][0]
    assert edge["relationshipPath"] == target
    assert edge["resolutionStatus"] == "resolved_by_name"
    assert edge["cacheDiagnostic"] == [
        {"address": "A1", "cached": 500, "current": 600, "agreement": False}
    ]
    assert edge["versionCaveat"]


def test_ambiguity_pins_stale_deleted_and_durable(tmp_path):
    a = fixture_source(tmp_path, 1, "same.xlsx", {"A1": 500})
    b = fixture_source(tmp_path, 2, "SAME.xlsx", {"A1": 600})
    reader = fixture_source(
        tmp_path,
        3,
        "reader.xlsx",
        {"A1": "='[1]S'!A1"},
        links=[_link("same.xlsx", "S", {"A1": 600})],
    )
    sources = [a, b, reader]
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(sources, handle(reader), direction="upstream", depth=5)
    assert {b["reason"] for b in result["boundaries"]} == {
        "ambiguous_external",
        "cache_only",
    }
    assert len(result["nodes"]) == 1 and not result["edges"]
    assert result["coverage"]["exhaustive"] and not result["coverage"]["complete"]
    fact = list_links(catalog, sources)["links"][0]
    target = next(r for r in fact["candidates"] if r["projectId"] == a.project_id)
    set_pin(catalog, sources, fact["id"], target["projectId"], target["sha256"])
    pinned = catalog.lineage(sources, handle(reader), direction="upstream", depth=5)
    edge = pinned["edges"][0]
    endpoint = next(n for n in pinned["nodes"] if n["id"] == edge["source"])
    assert edge["resolutionStatus"] == "pinned"
    assert edge["source"] == handle(a) and edge["target"] == handle(reader)
    assert endpoint["revision"] == target
    assert node_signature(a, "A1")[3:] == (target["projectId"], target["sha256"])
    assert endpoint["value"] == 500
    for path in catalog.directory.iterdir():
        path.unlink()
    catalog = Catalog(catalog.directory)
    catalog.ensure(sources)
    assert list_links(catalog, sources)["links"][0]["pin"]["projectId"] == a.project_id
    replaced = fixture_source(tmp_path, 1, "same.xlsx", {"A1": 700})
    sources = [replaced, b, reader]
    catalog.ensure(sources)
    stale = list_links(catalog, sources)["links"][0]
    assert (
        stale["status"] == "pinned"
        and stale["stalePin"]
        and stale["targetRevision"] is None
    )
    downstream = catalog.lineage(
        sources, handle(replaced), direction="downstream", depth=5
    )
    assert not downstream["edges"] and not downstream["coverage"]["complete"]
    assert (
        next(
            b
            for b in downstream["boundaries"]
            if b["reason"] == "possible_unresolved_reader"
        )["count"]
        == 1
    )
    catalog.ensure([b, reader])
    assert list_links(catalog, [b, reader])["links"][0]["stalePin"]
    set_pin(catalog, [b, reader], fact["id"], remove=True)
    assert list_links(catalog, [b, reader])["links"][0]["status"] == "resolved_by_name"


@pytest.mark.parametrize(
    "operand,subtype",
    [
        ("'[1]S'!Name", "defined_name_or_invalid"),
        ("[1]!Table1[Amount]", "structured"),
        ("'[1]S'!A:A", "whole_row_or_column"),
        ("'[1]S'!1:3", "whole_row_or_column"),
        ("'[1]S:T'!A1", "three_dimensional"),
        ("INDIRECT(\"'[1]S'!A1\")", "dynamic"),
    ],
)
def test_unsupported_external_grammar(tmp_path, operand, subtype):
    reader = fixture_source(
        tmp_path,
        1,
        "reader.xlsx",
        {"A1": "=" + operand},
        links=[_link("missing.xlsx", "S")],
    )
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage([reader], handle(reader), direction="upstream", depth=5)
    assert result["coverage"]["boundaries"] == ["unsupported_reference"]
    assert not result["edges"]
    fact = list_links(catalog, [reader])["links"][0]
    assert not fact["supported"] and fact["referenceType"] == subtype


def test_cross_workbook_cycle(tmp_path):
    a = fixture_source(
        tmp_path, 1, "a.xlsx", {"A1": "='[1]S'!A1"}, links=[_link("b.xlsx", "S")]
    )
    b = fixture_source(
        tmp_path, 2, "b.xlsx", {"A1": "='[1]S'!A1"}, links=[_link("a.xlsx", "S")]
    )
    catalog = Catalog(tmp_path / "owner" / "catalog")
    for direction in ("upstream", "downstream"):
        result = catalog.lineage([a, b], handle(a), direction=direction, depth=10)
        assert len(result["nodes"]) == len(result["edges"]) == 2
        assert len(result["stronglyConnectedComponents"]) == 1
        assert len(result["stronglyConnectedComponents"][0]) == 2
        assert {b["reason"] for b in result["boundaries"]} == {"cycle"}
        assert result["coverage"]["exhaustive"] and not result["coverage"]["complete"]


@pytest.mark.parametrize(
    "budget,reason",
    [
        ({"depth": 0}, "depth"),
        ({"limit": 1}, "node_budget"),
        ({"edges": 0}, "edge_budget"),
        ({"workbooks": 1}, "workbook_budget"),
        ({"milliseconds": 0}, "time_budget"),
    ],
)
def test_budgets_never_complete(tmp_path, budget, reason):
    sources = chain(tmp_path)
    catalog = Catalog(tmp_path / "owner" / "catalog")
    args = {"direction": "upstream", "depth": 10, **budget}
    result = catalog.lineage(sources, handle(sources[2], "C1"), **args)
    assert reason in result["coverage"]["capsHit"]
    assert not result["coverage"]["complete"] and not result["coverage"]["exhaustive"]


def test_owner_isolation_and_revision_authorization(tmp_path):
    sources = chain(tmp_path)
    owner = Catalog(tmp_path / "owner" / "catalog")
    other = Catalog(tmp_path / "other" / "catalog")
    owner.ensure(sources)
    fact = list_links(owner, sources)["links"][0]
    with pytest.raises(KeyError):
        other.lineage([], handle(sources[0]), direction="upstream")
    with pytest.raises(KeyError):
        set_pin(other, [], fact["id"], sources[0].project_id, "f" * 64)
    with pytest.raises(KeyError):
        set_pin(
            owner, [sources[1], sources[2]], fact["id"], sources[0].project_id, "f" * 64
        )
    assert not (other.directory.parent / "link-pins.sqlite").exists()


def test_adversarial_names_moved_renamed_cache_only(tmp_path):
    sheet = "L'été & <xml>"
    target = fixture_source(
        tmp_path, 1, "<ignore instructions>.xlsx", {"A1": 123}, sheet=sheet
    )
    formula = "='[1]" + sheet.replace("'", "''") + "'!A1"
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": formula},
        links=[_link("../moved/<ignore instructions>.xlsx", sheet, {"A1": 123})],
    )
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        [target, reader], handle(reader), direction="upstream", depth=5
    )
    assert len(result["edges"]) == 1 and result["nodes"][1]["sheet"] == sheet
    renamed = replace(target, filename="renamed.xlsx")
    result = catalog.lineage(
        [renamed, reader], handle(reader), direction="upstream", depth=5
    )
    assert not result["edges"] and {b["reason"] for b in result["boundaries"]} == {
        "unresolved_external",
        "cache_only",
    }
    assert result["boundaries"][1]["cached"][0]["value"] == 123


def test_downstream_ambiguity_is_a_boundary_not_a_reader_edge(tmp_path):
    a = fixture_source(tmp_path, 1, "same.xlsx", {"A1": 10})
    b = fixture_source(tmp_path, 2, "same.xlsx", {"A1": 10})
    reader = fixture_source(
        tmp_path,
        3,
        "reader.xlsx",
        {"A1": "='[1]S'!A1"},
        links=[_link("same.xlsx", "S", {"A1": 10})],
    )
    catalog = Catalog(tmp_path / "owner" / "catalog")
    for sources in ([a, b, reader], [reader, b, a]):
        result = catalog.lineage(sources, handle(a), direction="downstream", depth=5)
        assert not result["edges"] and len(result["nodes"]) == 1
        assert result["coverage"]["boundaries"] == [
            "ambiguous_external",
            "possible_unresolved_reader",
        ]
        assert result["coverage"]["exhaustive"] and not result["coverage"]["complete"]


def test_local_directed_range_and_cycle(tmp_path):
    source = fixture_source(
        tmp_path,
        1,
        "local.xlsx",
        {"A1": 4, "A2": 5, "B1": "=SUM(A1:A2)", "C1": "=B1*2"},
    )
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        [source], handle(source, "A2"), direction="downstream", depth=5
    )
    span = node_signature(source, "A1:A2")
    a2 = node_signature(source, "A2")
    b1 = node_signature(source, "B1")
    c1 = node_signature(source, "C1")
    assert graph_signature(result) == {
        "nodes": sorted([span, a2, b1, c1]),
        "edges": [
            (span, b1, "formula"),
            (a2, span, "range"),
            (b1, c1, "formula"),
        ],
        "boundaries": [],
    }
    cyclic = fixture_source(tmp_path, 2, "cycle.xlsx", {"A1": "=B1", "B1": "=A1"})
    result = catalog.lineage([cyclic], handle(cyclic), direction="upstream", depth=10)
    assert len(result["stronglyConnectedComponents"][0]) == 2


def test_response_link_and_legacy_index_budgets(tmp_path):
    sources = chain(tmp_path)
    catalog = Catalog(
        tmp_path / "owner" / "catalog", CatalogLimits(response_bytes=4096)
    )
    result = catalog.lineage(
        sources, handle(sources[2], "C1"), direction="upstream", depth=10
    )
    assert (
        len(json.dumps(result, ensure_ascii=False, separators=(",", ":")).encode())
        <= 4096
    )
    assert "response_bytes" in result["coverage"]["capsHit"]
    assert {r["projectId"] for r in result["revisions"]} == {
        s.project_id for s in sources
    }
    assert result["trace"]["revisions"] == result["revisions"]
    assert result["trace"]["counts"]["workbooks"] == 3
    ids = {n["id"] for n in result["nodes"]}
    assert all(e["source"] in ids and e["target"] in ids for e in result["edges"])
    assert all(b["node"] in ids for b in result["boundaries"])
    catalog = Catalog(tmp_path / "capped" / "catalog", CatalogLimits(lineage_links=0))
    result = catalog.lineage(
        sources, handle(sources[2], "C1"), direction="upstream", depth=10
    )
    assert "link_budget" in result["coverage"]["capsHit"]
    assert not result["edges"]
    with sqlite3.connect(sources[0].path) as db:
        metadata = json.loads(
            db.execute("SELECT value FROM metadata WHERE key='meta'").fetchone()[0]
        )
        metadata.pop("externalFactsVersion")
        db.execute(
            "UPDATE metadata SET value=? WHERE key='meta'", (json.dumps(metadata),)
        )
    catalog = Catalog(tmp_path / "legacy" / "catalog")
    result = catalog.lineage(
        sources, handle(sources[0]), direction="upstream", depth=10
    )
    assert "index_rebuild_required" in result["coverage"]["capsHit"]


def install_projects(tmp_path, sources, owner):
    import shutil

    for source in sources:
        directory = tmp_path / owner / "projects" / source.project_id
        directory.mkdir(parents=True)
        (directory / "references").mkdir()
        shutil.copyfile(source.path, directory / "index.sqlite")
        shutil.copyfile(
            source.path.parent / "workbook.xlsx", directory / "workbook.xlsx"
        )
        (directory / "project.json").write_text(
            json.dumps(
                {
                    "id": source.project_id,
                    "filename": source.filename,
                    "name": source.filename,
                    "createdAt": 9999999999,
                    "revision": "legacy",
                }
            )
        )


def test_api_cli_pin_and_owner_isolation(tmp_path, capsys):
    from test_web_index import call

    from linexcel.cli import main
    from linexcel.web import create_app

    sources = chain(tmp_path)
    owner = hashlib.sha256(b"alice").hexdigest()
    root = tmp_path / "data"
    install_projects(root, sources, owner)
    app = create_app(root, retention_days=None, resolve_user=lambda e: e.get("user"))
    try:
        code, links = call(app, "GET", "/api/library/links")
        assert code == 200
        fact = links["links"][0]
        target = fact["targetRevision"]
        body = {
            "linkId": fact["id"],
            "targetProject": target["projectId"],
            "targetSha256": target["sha256"],
        }
        assert call(app, "POST", "/api/library/pins", body, user="bob")[0] == 404
        assert (
            call(app, "POST", "/api/library/pins", {**body, "targetSha256": "f" * 64})[
                0
            ]
            == 404
        )
        assert (
            call(app, "POST", "/api/library/pins", {**body, "owner": owner})[0] == 400
        )
        assert call(app, "POST", "/api/library/pins", body)[0] == 200
        assert (
            call(app, "GET", "/api/library/links", query="status=pinned")[1]["links"][
                0
            ]["status"]
            == "pinned"
        )
        code, result = call(
            app,
            "GET",
            "/api/library/lineage",
            query="direction=upstream&depth=10&handle=" + handle(sources[2], "C1"),
        )
        assert code == 200 and len(result["revisions"]) == 3
        assert (
            call(
                app,
                "GET",
                "/api/library/lineage",
                query="direction=upstream&handle=" + handle(sources[0]),
                user="bob",
            )[0]
            == 404
        )
        assert (
            call(
                app,
                "GET",
                "/api/library/lineage",
                query="direction=sideways&handle=" + handle(sources[0]),
            )[0]
            == 400
        )
        assert (
            call(app, "DELETE", "/api/library/pins/" + fact["id"], user="bob")[0] == 404
        )
        assert call(app, "DELETE", "/api/library/pins/" + fact["id"])[0] == 200
    finally:
        app.close()
    base = ["library", "--data-dir", str(root), "--owner", owner]
    assert main([*base, "links", "--status", "resolved_by_name"]) == 0
    assert len(json.loads(capsys.readouterr().out)["links"]) == 2
    assert main([*base, "pin", fact["id"], target["projectId"], target["sha256"]]) == 0
    assert json.loads(capsys.readouterr().out)["pin"]["sha256"] == target["sha256"]
    assert (
        main(
            [
                *base,
                "lineage",
                handle(sources[2], "C1"),
                "--direction",
                "upstream",
                "--depth",
                "10",
            ]
        )
        == 0
    )
    assert len(json.loads(capsys.readouterr().out)["revisions"]) == 3
    assert main([*base, "unpin", fact["id"]]) == 0
    assert json.loads(capsys.readouterr().out)["pin"] is None


def test_pin_serializes_with_rebuild_and_pinned_generation_survives(
    tmp_path, monkeypatch
):
    from linexcel.catalog_build import WorkbookBuilder

    sources = chain(tmp_path)
    catalog = Catalog(tmp_path / "owner" / "catalog", CatalogLimits(connections=1))
    catalog.ensure(sources)
    fact = list_links(catalog, sources)["links"][0]
    entered, release = threading.Event(), threading.Event()
    original = WorkbookBuilder.body

    def blocked(builder):
        entered.set()
        assert release.wait(20)
        return original(builder)

    monkeypatch.setattr(WorkbookBuilder, "body", blocked)
    with catalog.pin() as (db, current), ThreadPoolExecutor(max_workers=2) as executor:
        build = executor.submit(catalog.ensure, sources, force=True)
        assert entered.wait(10)
        target = fact["targetRevision"]
        pin = executor.submit(
            set_pin, catalog, sources, fact["id"], target["projectId"], target["sha256"]
        )
        assert db.execute("SELECT count(*) FROM links").fetchone()[0] == 2
        assert (catalog.directory / f"{current['generation']}.sqlite").exists()
        release.set()
        build.result(timeout=30)
        pin.result(timeout=30)
    assert not (catalog.directory / f"{current['generation']}.sqlite").exists()
    result = catalog.lineage(
        sources, handle(sources[2], "C1"), direction="upstream", depth=10
    )
    assert result["trace"]["counts"]["indexConnectionsPeak"] == 1
    assert len(result["revisions"]) == 3


def test_range_and_blank_handles_can_be_traversed_again(tmp_path):
    target = fixture_source(tmp_path, 1, "target.xlsx", {"A1": 1, "A2": 2})
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": "=SUM('[1]S'!A1:A2)", "A2": "='[1]S'!A3"},
        links=[_link("target.xlsx", "S")],
    )
    catalog = Catalog(tmp_path / "owner" / "catalog")
    sources = [target, reader]
    graph = catalog.lineage(sources, handle(reader), direction="upstream", depth=10)
    selected = next(n for n in graph["nodes"] if n["kind"] == "range")
    result = catalog.lineage(
        sources, selected["handle"], direction="upstream", depth=10
    )
    assert len(result["nodes"]) == 3 and result["coverage"]["complete"]
    graph = catalog.lineage(
        sources, handle(reader, "A2"), direction="upstream", depth=10
    )
    blank = next(n for n in graph["nodes"] if n["valueProvenance"] == "saved_blank")
    result = catalog.lineage(sources, blank["handle"], direction="downstream", depth=10)
    assert len(result["edges"]) == 1


@pytest.mark.parametrize("blank", [False, True])
@pytest.mark.parametrize(
    "root,readers",
    [
        ("B2", ["A1", "A3", "A4"]),
        ("B3", ["A2", "A3"]),
        ("B2:B3", ["A1", "A2", "A3", "A4"]),
    ],
)
def test_downstream_member_does_not_inherit_other_range_readers(
    tmp_path, blank, root, readers
):
    cells = {"B2": 10, "B4": "=SUM(B2:B3)", "C4": "=B4*2"}
    if not blank:
        cells["B3"] = 20
    target = fixture_source(tmp_path, 1, "budget.xlsx", cells)
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {
            "A1": "='[1]S'!$B$2",
            "A2": "='[1]S'!$B$3",
            "A3": "=SUM('[1]S'!B2:B3)",
            "A4": "=SUM('[1]S'!B1:B2)",
            "A5": "='[1]S'!C4",
            "D5": "=A5+1",
        },
        links=[_link("budget.xlsx", "S")],
    )
    sources = [target, reader]
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        sources, handle(target, root), direction="downstream", depth=10
    )
    selected, span, b4, c4 = (
        node_signature(target, address) for address in (root, "B2:B3", "B4", "C4")
    )
    a5, d5 = (node_signature(reader, address) for address in ("A5", "D5"))
    direct = [node_signature(reader, address) for address in readers]
    expected_edges = [
        (span, b4, "formula"),
        (b4, c4, "formula"),
        (c4, a5, "external_link"),
        (a5, d5, "formula"),
        *((selected, node, "external_link") for node in direct),
    ]
    if root != "B2:B3":
        expected_edges.append((selected, span, "range"))
    assert graph_signature(result) == {
        "nodes": sorted({selected, span, b4, c4, a5, d5, *direct}),
        "edges": sorted(expected_edges),
        "boundaries": [],
    }
    assert result["coverage"]["complete"] and result["coverage"]["exhaustive"]
    if blank and root == "B3":
        assert result["nodes"][0]["valueProvenance"] == "saved_blank"


@pytest.mark.parametrize("identity", ["unresolved", "ambiguous"])
@pytest.mark.parametrize(
    "root,readers",
    [
        ("B2", ["A1", "A3", "A4"]),
        ("B3", ["A2", "A3"]),
        ("B2:B3", ["A1", "A2", "A3", "A4"]),
    ],
)
def test_downstream_range_member_keeps_exact_uncertain_reader_scope(
    tmp_path, identity, root, readers
):
    target = fixture_source(
        tmp_path,
        1,
        "budget.xlsx",
        {"B2": 10, "B3": 20, "B4": "=SUM(B2:B3)"},
    )
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {
            "A1": "='[1]S'!B2",
            "A2": "='[1]S'!B3",
            "A3": "=SUM('[1]S'!B2:B3)",
            "A4": "=SUM('[1]S'!B1:B2)",
        },
        links=[
            _link("oldname.xlsx" if identity == "unresolved" else "budget.xlsx", "S")
        ],
    )
    sources = [target, reader]
    if identity == "ambiguous":
        sources.append(fixture_source(tmp_path, 3, "budget.xlsx", {"B2": 99}))
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        sources, handle(target, root), direction="downstream", depth=10
    )
    facts = list_links(catalog, sources)["links"]
    expected_ids = sorted(
        fact["id"] for fact in facts if fact["sourceNode"].rsplit("!", 1)[-1] in readers
    )
    selected, span, b4 = (
        node_signature(target, address) for address in (root, "B2:B3", "B4")
    )
    expected_edges = [(span, b4, "formula")]
    if root != "B2:B3":
        expected_edges.append((selected, span, "range"))
    expected_boundaries = [(selected, "possible_unresolved_reader")]
    if identity == "ambiguous":
        expected_boundaries.extend(
            (selected, "ambiguous_external") for _ in expected_ids
        )
    assert graph_signature(result) == {
        "nodes": sorted({selected, span, b4}),
        "edges": sorted(expected_edges),
        "boundaries": sorted(expected_boundaries),
    }
    possible = [
        boundary
        for boundary in result["boundaries"]
        if boundary["reason"] == "possible_unresolved_reader"
    ]
    assert possible == [
        {
            "node": handle(target, root),
            "reason": "possible_unresolved_reader",
            "count": len(readers),
            "linkIds": expected_ids,
            "sampleOmitted": 0,
        }
    ]
    assert not result["coverage"]["complete"] and result["coverage"]["exhaustive"]


@pytest.mark.parametrize("identity", ["resolved", "unresolved", "ambiguous"])
@pytest.mark.parametrize("root", ["B2", "B3", "B2:B3"])
def test_downstream_unsupported_reader_stays_possible_without_widening(
    tmp_path, root, identity
):
    target = fixture_source(
        tmp_path,
        1,
        "budget.xlsx",
        {"B2": 10, "B3": 20, "B4": "=SUM(B2:B3)"},
    )
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": "='[1]S'!ExternalName"},
        links=[
            _link("oldname.xlsx" if identity == "unresolved" else "budget.xlsx", "S")
        ],
    )
    sources = [target, reader]
    if identity == "ambiguous":
        sources.append(fixture_source(tmp_path, 3, "budget.xlsx", {"B2": 99}))
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        sources, handle(target, root), direction="downstream", depth=10
    )
    selected, span, b4 = (
        node_signature(target, address) for address in (root, "B2:B3", "B4")
    )
    expected_edges = [(span, b4, "formula")]
    if root != "B2:B3":
        expected_edges.append((selected, span, "range"))
    # No verified bounds: the selected cell and its genuine local consumer may
    # both be read. A range reached only as a membership step adds no new scope.
    reasons = []
    if identity != "resolved":
        reasons.append("possible_unresolved_reader")
    if identity != "unresolved":
        reasons.append("unsupported_reference")
    assert graph_signature(result) == {
        "nodes": sorted({selected, span, b4}),
        "edges": sorted(expected_edges),
        "boundaries": sorted(
            (node, reason) for node in (selected, b4) for reason in reasons
        ),
    }
    if identity != "resolved":
        possible = [
            b
            for b in result["boundaries"]
            if b["reason"] == "possible_unresolved_reader"
        ]
        assert len(possible) == 2 and all(b["count"] == 1 for b in possible)
    assert not result["coverage"]["complete"] and result["coverage"]["exhaustive"]


def test_power_query_and_macro_are_typed_boundaries(tmp_path):
    from pathlib import Path

    source = tmp_path / "power_query.xlsx"
    source.write_bytes(Path("tests/fixtures/power_query.xlsx").read_bytes())
    index = tmp_path / "pq.sqlite"
    build_index(source, index, filename=source.name)
    item = Source("a" * 32, source.name, index)
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        [item], handle(item, "A2", "Loaded"), direction="upstream", depth=10
    )
    assert {b["reason"] for b in result["boundaries"]} == {"power_query"}
    assert not result["coverage"]["complete"] and result["coverage"]["exhaustive"]
    source = Path("tests/fixtures/macros.xlsm")
    index = tmp_path / "macro.sqlite"
    build_index(source, index, filename=source.name)
    item = Source("b" * 32, source.name, index)
    with WorkbookIndex(index) as workbook:
        nid = workbook.db.execute(
            "SELECT id FROM nodes WHERE kind='vba' LIMIT 1"
        ).fetchone()[0]
        selected = cell_handle(revision(item.project_id, workbook.meta["sha256"]), nid)
    result = catalog.lineage([item], selected, direction="upstream", depth=10)
    assert {b["reason"] for b in result["boundaries"]} == {"macro"}


def test_imported_reference_copy_is_replaced_by_authoritative_revision(tmp_path):
    import shutil

    target = fixture_source(tmp_path, 1, "inputs.xlsx", {"A1": 123})
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": "='[1]S'!A1"},
        links=[_link("inputs.xlsx", "S", {"A1": 100})],
    )
    refs = tmp_path / "refs"
    refs.mkdir()
    shutil.copyfile(target.path.parent / "workbook.xlsx", refs / "inputs.xlsx")
    build_index(
        reader.path.parent / "workbook.xlsx",
        reader.path,
        refs_dir=refs,
        filename=reader.filename,
    )
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        [target, reader], handle(reader), direction="upstream", depth=5
    )
    assert len(result["nodes"]) == 2 and len(result["edges"]) == 1
    assert result["edges"][0]["kind"] == "external_link"
    assert result["nodes"][1]["value"] == 123
    assert result["nodes"][1]["revision"]["projectId"] == target.project_id


@pytest.mark.parametrize("column", ["2020", "inputs.xlsx"])
@pytest.mark.parametrize("dynamic", [False, True])
def test_local_structured_column_is_never_an_external_book(tmp_path, column, dynamic):
    from openpyxl import load_workbook
    from openpyxl.worksheet.table import Table

    expression = f"Sales[{column}]"
    formula = f"=SUM(OFFSET({expression},0,0))" if dynamic else f"=SUM({expression})"
    source = fixture_source(
        tmp_path,
        1,
        "local.xlsx",
        {
            "A1": column,
            "B1": "Value",
            "A2": 1,
            "A3": 2,
            "B2": 3,
            "B3": 4,
            "C2": formula,
        },
    )
    path = source.path.parent / "workbook.xlsx"
    book = load_workbook(path)
    book.active.add_table(Table(displayName="Sales", ref="A1:B3"))
    book.save(path)
    book.close()
    build_index(path, source.path, filename=source.filename)
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        [source], handle(source, "C2"), direction="upstream", depth=10
    )
    assert list_links(catalog, [source])["links"] == []
    assert "unresolved_external" not in result["coverage"]["boundaries"]
    if not dynamic:
        assert result["coverage"]["complete"]
        assert len(result["nodes"]) == 4 and len(result["edges"]) == 3


def test_incomplete_source_sheet_has_its_own_boundary(tmp_path):
    source = fixture_source(tmp_path, 1, "partial.xlsx", {"A1": 1})
    with sqlite3.connect(source.path) as db:
        meta = json.loads(
            db.execute("SELECT value FROM metadata WHERE key='meta'").fetchone()[0]
        )
        meta["sheetDetails"][0]["status"] = "partial"
        db.execute("UPDATE metadata SET value=? WHERE key='meta'", (json.dumps(meta),))
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage([source], handle(source), direction="upstream", depth=10)
    assert result["coverage"]["boundaries"] == [
        "input_constant",
        "source_index_incomplete",
    ]
    assert not result["coverage"]["exhaustive"] and not result["coverage"]["complete"]


@pytest.mark.parametrize(
    "sheet,address,possible",
    [
        ("S", "A1", True),
        ("S", "A1:B2", True),
        ("S", "B1:B3", False),
        ("Missing Sheet", "A1", False),
        ("Missing Sheet", "ExternalName", True),
    ],
)
def test_downstream_renamed_target_never_guesses_or_overstates_scope(
    tmp_path, sheet, address, possible
):
    target = fixture_source(tmp_path, 1, "renamed.xlsx", {"A1": 10})
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": f"='[1]{sheet}'!{address}"},
        links=[_link("oldname.xlsx", sheet, {"A1": 10})],
    )
    sources = [target, reader]
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(sources, handle(target), direction="downstream", depth=10)
    assert not result["edges"]
    assert graph_signature(result)["nodes"] == [node_signature(target, "A1")]
    assert result["coverage"]["complete"] is (not possible)
    assert result["trace"]["complete"] is (not possible)
    assert result["coverage"]["exhaustive"]
    assert result["coverage"]["scope"] == (
        "loaded owner revisions only; never-loaded workbooks are not visible"
    )
    assert result["trace"]["scope"] == result["coverage"]["scope"]
    fact = list_links(catalog, sources)["links"][0]
    assert fact["status"] == "unresolved" and fact["targetRevision"] is None
    assert result["boundaries"] == (
        [
            {
                "node": handle(target),
                "reason": "possible_unresolved_reader",
                "count": 1,
                "linkIds": [fact["id"]],
                "sampleOmitted": 0,
            }
        ]
        if possible
        else []
    )


@pytest.mark.parametrize("samples", [0, 3])
def test_possible_unresolved_readers_count_all_links_with_bounded_samples(
    tmp_path, samples
):
    target = fixture_source(tmp_path, 1, "renamed.xlsx", {"A1": 10})
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {f"B{r}": "='[1]S'!A1" for r in range(1, 42)},
        links=[_link("oldname.xlsx", "S")],
    )
    sources = [target, reader]
    catalog = Catalog(
        tmp_path / "owner" / "catalog",
        CatalogLimits(lineage_reader_samples=samples, response_bytes=4096),
    )
    result = catalog.lineage(sources, handle(target), direction="downstream", depth=10)
    with catalog.pin() as (db, _):
        expected_ids = [
            row[0] for row in db.execute("SELECT id FROM links ORDER BY id")
        ]
    assert len(expected_ids) == 41
    assert result["boundaries"] == [
        {
            "node": handle(target),
            "reason": "possible_unresolved_reader",
            "count": 41,
            "linkIds": expected_ids[:samples],
            "sampleOmitted": 41 - samples,
        }
    ]
    assert not result["edges"] and not result["coverage"]["complete"]
    assert result["coverage"]["exhaustive"] and result["coverage"]["capsHit"] == []
    assert result["trace"]["limits"]["unresolvedReaderSamples"] == samples
    assert (
        len(json.dumps(result, ensure_ascii=False, separators=(",", ":")).encode())
        <= 4096
    )
    # Other owners' link facts do not enter the selected owner's generation.
    other = Catalog(tmp_path / "other" / "catalog")
    own_scope = other.lineage(
        [target], handle(target), direction="downstream", depth=10
    )
    assert own_scope["coverage"]["complete"] and not own_scope["boundaries"]


def test_graph_signatures_distinguish_same_filename_and_address_revisions(tmp_path):
    a = fixture_source(tmp_path, 1, "same.xlsx", {"A1": 10})
    b = fixture_source(tmp_path, 2, "same.xlsx", {"A1": 20})
    catalog = Catalog(tmp_path / "owner" / "catalog")
    baseline = graph_signature(catalog.lineage([a], handle(a), direction="upstream"))
    for other in (
        replace(a, project_id=b.project_id),
        replace(b, project_id=a.project_id),
    ):
        observed = graph_signature(
            catalog.lineage([other], handle(other), direction="upstream")
        )
        assert observed != baseline
        assert observed["nodes"][0][:3] == baseline["nodes"][0][:3]
