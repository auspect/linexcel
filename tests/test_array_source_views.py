"""Hand-computed OOXML array caches survive every source-value reader."""

import csv
import hashlib
import io
import json
import shutil
import sqlite3
import time
import zipfile
from contextlib import closing
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula
from test_web import request

from linexcel.ask_tools import QuestionTools
from linexcel.catalog import Catalog, Source, revision
from linexcel.excel_presentation import sheet_presentation
from linexcel.header_bands import literal
from linexcel.insights import extract_workbook_context
from linexcel.lazy_store import WorkbookIndex, _pack, build_index
from linexcel.web import _write, create_app
from linexcel.web_data import build_source_database
from linexcel.web_report import prepare_report

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
METADATA = f'''<metadata xmlns="{MAIN}"
 xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
 <metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000"
 copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1"
 clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>
 <futureMetadata name="XLDAPR" count="1"><bk><extLst>
 <ext uri="{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}">
 <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>
 </ext></extLst></bk></futureMetadata>
 <cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'''


def array_bytes(*, blank=False, cached_heading=False, large=False):
    book = Workbook()
    sheet = book.active
    sheet.title = "Feuil1"
    if large:
        sheet["A1"] = ArrayFormula(ref="A1:A205", text="=SEQUENCE(205)")
        for row in range(2, 206):
            sheet.cell(row, 1, row)
        sheet["B1"] = "=SUM(A1:A205)"
        caches = {f"A{row}": row for row in range(1, 206)}
    else:
        sheet.append(["Base amount", "Base units"])
        for values in ((100, 2), (200, 3), (300, 5)):
            sheet.append(values)
        sheet["G2"] = ArrayFormula(ref="G2:H3", text="=A2:B3")
        sheet["J2"] = ArrayFormula(ref="J2:K4", text="=A2:B4")
        caches = {
            "G2": 100,
            "H2": 2,
            "G3": 200,
            "H3": 3,
            "J2": 100,
            "K2": 2,
            "J3": 200,
            "K3": 3,
            "J4": 300,
            "K4": 5,
        }
        if blank:
            caches["K3"] = None
        if cached_heading:
            caches["H2"] = "Cached heading"
        for address, value in caches.items():
            if address not in {"G2", "J2"}:
                sheet[address] = value
        for name, address in (("LegacyUnits", "H3"), ("SpillUnits", "K4")):
            book.defined_names.add(
                DefinedName(name, attr_text=f"Feuil1!${address[0]}${address[1:]}")
            )
    source = io.BytesIO()
    book.save(source)
    output = io.BytesIO()
    with zipfile.ZipFile(source) as archive, zipfile.ZipFile(output, "w") as result:
        for name in archive.namelist():
            data = archive.read(name)
            if name == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(data)
                for row in root.findall("{*}sheetData/{*}row"):
                    for cell in row.findall("{*}c"):
                        address = cell.get("r")
                        if address in caches:
                            value = cell.find("{*}v")
                            if value is None:
                                value = ET.SubElement(cell, f"{{{MAIN}}}v")
                            cached = caches[address]
                            value.text = str(cached) if cached is not None else None
                            if isinstance(cached, str):
                                cell.set("t", "str")
                                inline = cell.find("{*}is")
                                if inline is not None:
                                    cell.remove(inline)
                        if address == "J2":
                            cell.set("cm", "1")
                if blank:
                    row = root.find("{*}sheetData/{*}row[@r='3']")
                    ET.SubElement(row, f"{{{MAIN}}}c", r="K3")
                data = ET.tostring(root)
            elif name == "xl/_rels/workbook.xml.rels":
                root = ET.fromstring(data)
                ET.SubElement(
                    root,
                    f"{{{REL}}}Relationship",
                    Id="rIdMetadata",
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
                    Target="metadata.xml",
                )
                data = ET.tostring(root)
            elif name == "[Content_Types].xml":
                root = ET.fromstring(data)
                ET.SubElement(
                    root,
                    "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
                    PartName="/xl/metadata.xml",
                    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
                )
                data = ET.tostring(root)
            result.writestr(name, data)
        result.writestr("xl/metadata.xml", METADATA)
    return output.getvalue(), caches


@pytest.fixture(params=[False, True], ids=["saved-members", "blank-member"])
def arrays(tmp_path, request):
    content, expected = array_bytes(blank=request.param)
    source, index = tmp_path / "arrays.xlsx", tmp_path / "index.sqlite"
    source.write_bytes(content)
    build_index(source, index)
    return source, index, expected


@pytest.fixture
def array_project(tmp_path, arrays):
    source, index, expected = arrays
    app = create_app(tmp_path / "app", resolve_user=lambda env: env.get("user"))
    owner, pid = hashlib.sha256(b"alice").hexdigest(), "a" * 32
    path = app.store.root / owner / "projects" / pid
    (path / "references").mkdir(parents=True)
    shutil.copyfile(source, path / "workbook.xlsx")
    shutil.copyfile(index, path / "index.sqlite")
    _write(
        path / "project.json",
        {
            "id": pid,
            "name": "Array source views",
            "filename": "arrays.xlsx",
            "createdAt": time.time(),
            "revision": app.store._revision(path),
        },
    )
    try:
        yield app, owner, pid, path, expected
    finally:
        app.close()


def test_array_presentation_keeps_members_anchors_and_blank(arrays, monkeypatch):
    source, index, expected = arrays
    before = hashlib.sha256(source.read_bytes()).hexdigest()

    def forbidden(*args, **kwargs):
        pytest.fail("Source-value readers must never recalculate")

    monkeypatch.setattr("linexcel.lazy.evaluate_node", forbidden)
    monkeypatch.setattr("linexcel.analyzer.analyze_workbook", forbidden)
    view = sheet_presentation(source, index, "Feuil1", rows=6, columns=11)
    cells = {c["address"]: c for c in view["cells"]}
    assert {address: cells[address]["value"] for address in expected} == expected
    for address, value in expected.items():
        cell = cells[address]
        assert cell["value"] == value
        assert cell["display"] == (str(value) if value is not None else "")
        assert cell["valueStatus"] == ("present" if value is not None else "blank")
        if value is not None:
            provenance = cell["sourceValueProvenance"]
            assert provenance["role"] == "formula_cache"
            assert provenance["workbookSha256"] == before
            assert provenance["cell"] == address
            assert provenance["rawXmlValue"] == str(value)
            if address not in {"G2", "J2"}:
                assert cell["arrayAnchor"] == (
                    "Feuil1!G2" if address[0] in "GH" else "Feuil1!J2"
                )
                assert cell["sourceFormula"] is None
    assert cells["G2"]["sourceFormula"] == "=A2:B3"
    assert cells["J2"]["sourceFormula"] == "=A2:B4"
    assert hashlib.sha256(source.read_bytes()).hexdigest() == before


def test_source_database_counts_and_exports_array_caches(arrays, tmp_path):
    _, index, expected = arrays
    output = tmp_path / "data.sqlite"
    metadata = build_source_database(index, output, max_bytes=4 * 1024 * 1024)
    assert metadata["sheets"][0]["cellCount"] == 8 + sum(
        v is not None for v in expected.values()
    )
    with closing(sqlite3.connect(output)) as db:
        cells = {
            row[0]: row[1:]
            for row in db.execute(
                "SELECT cell,value,value_role,value_status,value_provenance FROM cells"
            )
        }
        for address, value in expected.items():
            if value is None:
                assert (
                    address not in cells
                )  # The declared source-cell policy excludes true blanks.
                continue
            stored, role, status, provenance = cells[address]
            assert stored == value and role == "formula_cache" and status == "present"
            assert json.loads(provenance)["cell"] == address
        assert db.execute("SELECT G,H,J,K FROM sheet_001 WHERE row=3").fetchone() == (
            200,
            3,
            200,
            expected["K3"],
        )


def test_array_data_routes_exports_and_cached_snapshot_upgrade(
    array_project, monkeypatch
):
    import linexcel.web_data as data

    app, _, pid, path, expected = array_project
    prefix = f"/api/projects/{pid}"
    # Simulate the previous cached source-data generation, then switch to the
    # current reader version. Existing projects must rebuild without reimport.
    version = data.DATA_VERSION
    with monkeypatch.context() as patch:
        patch.setattr(data, "DATA_VERSION", version - 1)
        assert request(app, "GET", prefix + "/data")[0] == 200
    stamp = json.loads((path / "source-data.json").read_text())
    with closing(sqlite3.connect(path / "source-data.sqlite")) as db, db:
        db.execute("DELETE FROM cells WHERE cell='H3'")
        db.execute("UPDATE sheet_001 SET H=NULL WHERE row=3")
    code, raw, _ = request(app, "GET", prefix + "/data")
    assert code == 200
    assert (
        json.loads((path / "source-data.json").read_text())["signature"]
        != stamp["signature"]
    )
    code, raw, _ = request(
        app,
        "GET",
        prefix + "/data/presentation",
        QUERY_STRING="sheet=Feuil1&row=1&column=1&rows=6&columns=11",
    )
    assert code == 200
    presented = {c["address"]: c for c in json.loads(raw)["cells"]}
    for address, value in expected.items():
        assert presented[address]["value"] == value
    code, raw, _ = request(
        app,
        "GET",
        prefix + "/datasette/source/cells.json",
        QUERY_STRING="_shape=objects",
    )
    assert code == 200
    exported = {c["cell"]: c for c in json.loads(raw)["rows"]}
    code, raw, _ = request(app, "GET", prefix + "/datasette/source/cells.csv")
    assert code == 200
    csv_cells = {c["cell"]: c for c in csv.DictReader(io.StringIO(raw.decode()))}
    code, raw, _ = request(
        app,
        "GET",
        prefix + "/datasette/source/sheet_001.json",
        QUERY_STRING="_shape=objects",
    )
    assert code == 200
    grid = {r["row"]: r for r in json.loads(raw)["rows"]}
    for address, value in expected.items():
        assert grid[int(address[1:])][address[0]] == value
        if value is not None:
            assert exported[address]["value"] == value
            assert exported[address]["value_role"] == "formula_cache"
            assert json.loads(csv_cells[address]["value_json"]) == value
    code, raw, _ = request(app, "GET", prefix + "/export")
    assert code == 200
    nodes = {n["cell"]: n for n in json.loads(raw)["graph"]["nodes"] if n.get("cell")}
    for address, value in expected.items():
        if value is not None:
            assert nodes[address]["value"] == value
            assert nodes[address]["sourceValueProvenance"]["role"] == "formula_cache"
    assert not app.store.tasks  # No import, evaluation or AI tasks were started.


def test_catalog_numeric_lookup_and_ask_keep_array_values(array_project, tmp_path):
    app, owner, pid, path, expected = array_project
    item = Source(pid, "arrays.xlsx", path / "index.sqlite")
    catalog = Catalog(tmp_path / "catalog")
    hit = next(
        h for h in catalog.search([item], "5")["results"] if h["address"] == "K4"
    )
    assert hit["value"] == 5 and hit["valueProvenance"] == "saved_cache"
    with WorkbookIndex(item.path) as index:
        rev = revision(pid, index.meta["sha256"])
    tools = QuestionTools(
        app.store,
        owner,
        {"attachment": None, "revisions": [rev]},
        tmp_path / "ask",
        time.monotonic() + 30,
    )
    read = tools.execute(
        "read_range", {"project": pid, "sheet": "Feuil1", "range": "G2:K4"}
    )["data"]
    cells = {c["address"]: c for c in read["cells"]}
    for address, value in expected.items():
        if value is not None:
            assert cells[address]["value"] == value
            assert cells[address]["valueProvenance"] == "saved_cache"
    found = tools.execute("find_value", {"number": 5, "tolerance": 0})["data"]
    assert any(c["address"] == "K4" and c["value"] == 5 for c in found["cells"])


def test_workbook_context_array_anchors_are_formulas(arrays):
    source, _, _ = arrays
    sheet = extract_workbook_context(source.read_bytes(), preview_columns=11)["sheets"][
        0
    ]
    row2 = sheet["preview"][1]["values"]
    assert row2[6] == "=A2:B3" and row2[9] == "=A2:B4"
    assert sheet["preview"][2]["values"][7] == 3


def test_array_cached_text_is_searchable_but_never_a_literal_heading(tmp_path):
    content, _ = array_bytes(cached_heading=True)
    source, index_path = tmp_path / "array.xlsx", tmp_path / "index.sqlite"
    source.write_bytes(content)
    build_index(source, index_path)
    item = Source("a" * 32, "array.xlsx", index_path)
    hit = Catalog(tmp_path / "catalog").search([item], "Cached heading")["results"][0]
    assert hit["address"] == "H2" and hit["value"] == "Cached heading"
    with WorkbookIndex(index_path) as index:
        node = index.node("Feuil1!H2")
        assert literal(node) is None
        assert not any(
            label["text"] == "Cached heading"
            for label in index.cell_labels("Feuil1", 3, 8)
        )


def test_existing_array_member_input_provenance_is_corrected_on_read(arrays):
    _, index_path, _ = arrays
    with WorkbookIndex(index_path, writable=True) as index:
        node = index["Feuil1!H3"]
        node["provenance"] = "saved_input"
        node["sourceValueProvenance"]["role"] = "input"
        index.db.execute(
            "UPDATE nodes SET data=? WHERE id=?", (_pack(node), node["id"])
        )
        index.db.commit()
    with WorkbookIndex(index_path) as index:
        node = index["Feuil1!H3"]
        assert node["value"] == 3 and node["provenance"] == "saved_cache"
        assert node["sourceValueProvenance"]["role"] == "formula_cache"


def test_frozen_report_range_counts_include_every_array_result(tmp_path):
    content, _ = array_bytes(large=True)
    app = create_app(tmp_path / "app", resolve_user=lambda env: env.get("user"))
    owner, pid = hashlib.sha256(b"alice").hexdigest(), "a" * 32
    path = app.store.root / owner / "projects" / pid
    (path / "references").mkdir(parents=True)
    (path / "workbook.xlsx").write_bytes(content)
    build_index(path / "workbook.xlsx", path / "index.sqlite")
    _write(
        path / "project.json",
        {
            "id": pid,
            "name": "Array source views",
            "filename": "arrays.xlsx",
            "revision": app.store._revision(path),
        },
    )
    try:
        record = prepare_report(app.store, owner, pid, {"nodeIds": ["Feuil1!B1"]})[
            "records"
        ][0]["ranges"][0]
        assert record["representedCellCount"] == 205
        assert len(record["members"]) == 200 and record["omitted"] == 5
        assert record["members"][0]["cachedValue"] == 1
        assert record["members"][-1]["cachedValue"] == 200
        assert not app.store.tasks
    finally:
        app.close()
