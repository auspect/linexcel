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
from test_external import _link, linked_workbook

from linexcel.catalog import Catalog, CatalogLimits, Source, cell_handle, revision
from linexcel.catalog_links import list_links, set_pin
from linexcel.lazy_store import WorkbookIndex, build_index


def fixture_source(
    tmp_path, number, filename, cells, *, links=(), sheet="S", caches=None
):
    directory = tmp_path / str(number)
    directory.mkdir(exist_ok=True)
    content = linked_workbook(cells, links, sheet)
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
    nodes = {n["id"]: (n["filename"], n["address"]) for n in result["nodes"]}
    return {
        "nodes": sorted(nodes.values()),
        "edges": sorted(
            (nodes[e["source"]], nodes[e["target"]], e["kind"]) for e in result["edges"]
        ),
        "boundaries": sorted(
            (nodes[b["node"]], b["reason"]) for b in result["boundaries"]
        ),
    }


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


def test_exact_three_workbook_graph_and_value(tmp_path):
    sources = chain(tmp_path)
    catalog = Catalog(tmp_path / "owner" / "catalog")
    result = catalog.lineage(
        sources, handle(sources[2], "C1"), direction="upstream", depth=10
    )
    assert graph_signature(result) == {
        "nodes": [
            ("consolidated.xlsx", "C1"),
            ("inputs.xlsx", "A1"),
            ("inputs.xlsx", "A1:A2"),
            ("inputs.xlsx", "A2"),
            ("regional.xlsx", "B1"),
        ],
        "edges": [
            (("inputs.xlsx", "A1"), ("inputs.xlsx", "A1:A2"), "range"),
            (("inputs.xlsx", "A1:A2"), ("regional.xlsx", "B1"), "external_link"),
            (("inputs.xlsx", "A2"), ("inputs.xlsx", "A1:A2"), "range"),
            (("regional.xlsx", "B1"), ("consolidated.xlsx", "C1"), "external_link"),
        ],
        "boundaries": [
            (("inputs.xlsx", "A1"), "input_constant"),
            (("inputs.xlsx", "A2"), "input_constant"),
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
        "nodes": [
            ("consolidated.xlsx", "C1"),
            ("inputs.xlsx", "A2"),
            ("regional.xlsx", "B1"),
        ],
        "edges": [
            (("inputs.xlsx", "A2"), ("regional.xlsx", "B1"), "external_link"),
            (("regional.xlsx", "B1"), ("consolidated.xlsx", "C1"), "external_link"),
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
    assert (
        catalog.lineage(sources, handle(reader), direction="upstream", depth=5)[
            "edges"
        ][0]["resolutionStatus"]
        == "pinned"
    )
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
        assert result["coverage"]["boundaries"] == ["ambiguous_external"]
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
    assert graph_signature(result) == {
        "nodes": [
            ("local.xlsx", "A1:A2"),
            ("local.xlsx", "A2"),
            ("local.xlsx", "B1"),
            ("local.xlsx", "C1"),
        ],
        "edges": [
            (("local.xlsx", "A1:A2"), ("local.xlsx", "B1"), "formula"),
            (("local.xlsx", "A2"), ("local.xlsx", "A1:A2"), "range"),
            (("local.xlsx", "B1"), ("local.xlsx", "C1"), "formula"),
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
