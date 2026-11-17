"""Every read of the index's ``nodes`` table goes through ``WorkbookIndex``.

Batch 2 ter moves literal values out of one record per cell into per-sheet
grids (index schema 5). Before that change, every module that read the
``nodes`` table with its own SQL now asks ``WorkbookIndex`` instead, so the
storage can change in one place. Each test below pins what one of those read
sites returns today, through the public function that holds it, on a small
workbook whose expected answers can be checked by hand. The last test fails if
a direct read comes back.
"""

import ast
import base64
import io
import json
import time
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table

from linexcel.lazy_store import INDEX_FILENAME
from linexcel.web import create_app

SOURCE = Path(__file__).resolve().parent.parent / "src" / "linexcel"


def reads_workbook() -> Workbook:
    """Literals of every type, formulas, ranges, a name, a table and an array.

    Data!A1:D6 is the table ``tblData``: text, numbers, booleans and one
    error literal, with a copied formula column D. Calc reads Data through
    two ranges, a single cell and a defined name. Big!A1:A300 is a range of
    300 literals, wider than the 200 members a step or a report lists. Labels
    has a header row and repeated row labels for header bands. Arr holds a
    three-cell array formula.
    """
    book = Workbook()
    data = book.active
    data.title = "Data"
    data.append(["Region", "Amount", "Flag", "Double"])
    for row, (region, flag) in enumerate(
        [("North", True), ("South", False), ("East", True), ("West", "#N/A")]
        + [("North", True)],
        start=2,
    ):
        data.append([region, 10 * (row - 1), flag, f"=B{row}*2"])
    data.add_table(Table(displayName="tblData", ref="A1:D6"))
    calc = book.create_sheet("Calc")
    for label, content in [
        ("Total", "=SUM(Data!B2:B6)"),
        ("North", "=SUMIFS(Data!B2:B6,Data!A2:A6,A2)"),
        ("Rate", 0.5),
        ("Scaled", "=B1*Rate"),
        ("First", "=Data!B2"),
        ("Big", "=SUM(Big!A1:A300)"),
    ]:
        calc.append([label, content])
    book.defined_names.add(DefinedName("Rate", attr_text="Calc!$B$3"))
    big = book.create_sheet("Big")
    for value in range(1, 301):
        big.append([value])
    labels = book.create_sheet("Labels")
    labels.append(["Region", "Q1", "Q2"])
    for region, q1, q2 in [("North", 5, 7), ("South", 6, 8), ("North", 9, 4)]:
        labels.append([region, q1, q2])
    arr = book.create_sheet("Arr")
    for row in (1, 2, 3):
        arr.cell(row, 2, row)
    arr["A1"] = ArrayFormula("A1:A3", "=B1:B3*2")
    return book


def _request(app, method, path, body=None):
    data = json.dumps(body).encode() if body is not None else b""
    status = []
    payload = b"".join(
        app(
            {
                "REQUEST_METHOD": method,
                "PATH_INFO": path,
                "SCRIPT_NAME": "",
                "CONTENT_TYPE": "application/json",
                "CONTENT_LENGTH": str(len(data)),
                "wsgi.input": io.BytesIO(data),
                "wsgi.url_scheme": "http",
                "HTTP_HOST": "127.0.0.1:8765",
                "user": "alice",
            },
            lambda code, headers: status.append(code),
        )
    )
    return int(status[0].split()[0]), json.loads(payload)


@pytest.fixture(scope="module")
def project(tmp_path_factory):
    """The workbook imported through serve, as a user would."""
    root = tmp_path_factory.mktemp("index-reads")
    app = create_app(root, resolve_user=lambda env: env.get("user"))
    stream = io.BytesIO()
    reads_workbook().save(stream)
    code, created = _request(
        app,
        "POST",
        "/api/projects",
        {
            "workbook": {
                "name": "reads.xlsx",
                "data": base64.b64encode(stream.getvalue()).decode(),
            }
        },
    )
    assert code in {200, 201, 202}, created
    task_id = created["task"]["id"]
    deadline = time.monotonic() + 60
    while app.store.tasks[task_id]["status"] not in {"succeeded", "failed"}:
        assert time.monotonic() < deadline, "import did not finish"
        time.sleep(0.02)
    assert app.store.tasks[task_id]["status"] == "succeeded"
    owner = app.store.tasks[task_id]["owner"]
    pid = created["project"]["id"]
    path = app.store.project_path(owner, pid)
    try:
        yield app, owner, pid, path
    finally:
        app.close()


def _addresses(cells):
    return [f"{c['sheet']}!{c['address']}" for c in cells]


@pytest.fixture
def question_tools(project, tmp_path):
    from linexcel.ask_tools import QuestionTools
    from linexcel.lazy_store import WorkbookIndex

    app, owner, pid, path = project
    with WorkbookIndex(path / INDEX_FILENAME) as index:
        sha, sheets = index.meta["sha256"], index.meta["sheets"]
    scope = {
        "attachment": None,
        "revisions": [
            {
                "projectId": pid,
                "sha256": sha,
                "id": f"{pid}:{sha}",
                "filename": "reads.xlsx",
                "sheets": sheets,
            }
        ],
    }
    return QuestionTools(
        app.store, owner, scope, tmp_path / "run", time.monotonic() + 120
    ), pid


def test_ask_describe_sheet_samples_the_first_twelve_cells_in_sheet_order(
    question_tools,
):
    tools, pid = question_tools
    data = tools.execute("describe_sheet", {"project": pid, "sheet": "Data"})["data"]
    assert _addresses(data["cells"]) == [
        f"Data!{column}{row}" for row in (1, 2, 3) for column in "ABCD"
    ]
    assert [c["value"] for c in data["cells"][:4]] == [
        "Region",
        "Amount",
        "Flag",
        "Double",
    ]


def test_ask_find_value_scans_cells_in_import_order(question_tools):
    tools, _ = question_tools
    data = tools.execute("find_value", {"number": 40, "tolerance": 0})["data"]
    assert _addresses(data["cells"]) == ["Data!B5", "Big!A40"]
    # 341 literals and 11 formula cells; names and ranges are not cells.
    assert data["coverage"] == {"complete": True, "scanned": 352}


@pytest.fixture
def catalog(project, tmp_path):
    from linexcel.catalog import Catalog, CatalogLimits, Source

    _, _, pid, path = project
    source = Source(pid, "reads.xlsx", path / INDEX_FILENAME)
    built = Catalog(tmp_path / "catalog", CatalogLimits())
    built.ensure([source])
    return built, source


def _sheet_coverage(catalog):
    return {s["sheet"]: s for s in catalog.status()["coverage"]["sheets"]}


def test_catalog_counts_every_node_kind_of_every_sheet(catalog):
    sheets = _sheet_coverage(catalog[0])
    assert {sheet: s["nodeKinds"] for sheet, s in sheets.items()} == {
        "Data": {"cell": 5, "input": 19, "range": 2},
        "Calc": {"cell": 5, "input": 7},
        "Big": {"input": 300, "range": 1},
        "Labels": {"input": 12},
        "Arr": {"cell": 1, "input": 3, "range": 1},
        None: {"name": 1},
    }
    assert {sheet: s["nodesTotal"] for sheet, s in sheets.items()} == {
        "Data": 26,
        "Calc": 12,
        "Big": 301,
        "Labels": 12,
        "Arr": 5,
        None: 1,
    }


def test_catalog_lists_defined_names(catalog):
    with catalog[0].pin() as (db, _):
        names = db.execute(
            "SELECT sheet,address,text,node FROM entries WHERE kind='name'"
        ).fetchall()
    assert names == [(None, None, "rate", "name:*:Rate")]
    assert _sheet_coverage(catalog[0])[None]["nodesScanned"] == 1


def test_catalog_scans_every_cell_of_every_sheet(catalog):
    sheets = _sheet_coverage(catalog[0])
    assert all(s["nodesScanned"] == s["nodesTotal"] for s in sheets.values())
    # Text values, the error literal #N/A included; numbers are not text.
    assert {sheet: s["textCellsObserved"] for sheet, s in sheets.items()} == {
        "Data": 10,
        "Calc": 6,
        "Big": 0,
        "Labels": 6,
        "Arr": 0,
        None: 0,
    }
    with catalog[0].pin() as (db, _):
        cells = db.execute(
            "SELECT address,text FROM entries WHERE kind='cell' AND sheet='Data' "
            "AND address LIKE 'A%' ORDER BY address"
        ).fetchall()
    assert cells == [
        ("A1", "region"),
        ("A2", "north"),
        ("A3", "south"),
        ("A4", "east"),
        ("A5", "west"),
    ]


def _handle(source, node_id):
    from linexcel.catalog import cell_handle, revision
    from linexcel.lazy_store import WorkbookIndex

    with WorkbookIndex(source.path) as index:
        return cell_handle(revision(source.project_id, index.meta["sha256"]), node_id)


def _node_id(handle):
    from linexcel.catalog import parse_handle

    return parse_handle(handle)[2]


def test_lineage_compares_external_caches_with_the_cells_they_name(tmp_path):
    from test_external import _link
    from test_total_lineage import fixture_source, handle

    from linexcel.catalog import Catalog

    inputs = fixture_source(tmp_path, 1, "inputs.xlsx", {"A1": 600, "A3": "x"})
    reader = fixture_source(
        tmp_path,
        2,
        "reader.xlsx",
        {"A1": "=SUM('[1]S'!A1:A3)"},
        links=[_link("inputs.xlsx", "S", {"A1": 500, "A2": 7, "A3": "x"})],
    )
    result = Catalog(tmp_path / "catalog").lineage(
        [inputs, reader], handle(reader), direction="upstream", depth=4
    )
    # A1 changed since the cache, A2 is no longer stored, A3 agrees.
    assert result["edges"][0]["cacheDiagnostic"] == [
        {"address": "A1", "cached": 500, "current": 600, "agreement": False},
        {"address": "A2", "cached": 7, "current": None, "agreement": None},
        {"address": "A3", "cached": "x", "current": "x", "agreement": True},
    ]


def test_lineage_reaches_the_cells_a_range_holds(catalog):
    built, source = catalog
    result = built.lineage(
        [source], _handle(source, "Calc!B1"), direction="upstream", depth=3
    )
    edges = sorted(
        (_node_id(e["source"]), _node_id(e["target"]), e["kind"])
        for e in result["edges"]
    )
    assert edges == [
        ("Data!B2", "Data!B2:B6", "range"),
        ("Data!B2:B6", "Calc!B1", "formula"),
        ("Data!B3", "Data!B2:B6", "range"),
        ("Data!B4", "Data!B2:B6", "range"),
        ("Data!B5", "Data!B2:B6", "range"),
        ("Data!B6", "Data!B2:B6", "range"),
    ]


def test_lineage_stops_at_an_array_wider_than_its_node_budget(catalog):
    built, source = catalog
    anchor = _handle(source, "Arr!A1")
    # Arr!A1:A3 has three cells and only its anchor is stored: with a budget
    # of two nodes the members cannot be certified, so the walk says so.
    narrow = built.lineage([source], anchor, direction="downstream", limit=2)
    assert [(_node_id(b["node"]), b["reason"]) for b in narrow["boundaries"]] == [
        ("Arr!A1", "node_budget")
    ]
    wide = built.lineage([source], anchor, direction="downstream", limit=200)
    assert wide["boundaries"] == []


def _occurrence_pages(built, source, query, sheet):
    """Every occurrence of the first hit for ``query``, one page at a time."""
    hit = next(
        r for r in built.search([source], query)["results"] if r["sheet"] == sheet
    )
    group, after, pages = hit["occurrenceGroup"], 0, []
    while after is not None:
        page = built.occurrences(
            [source],
            generation=group["generation"],
            group=group["id"],
            after=after,
            limit=1,
        )
        pages.append([r["address"] for r in page["results"]])
        after = page["pagination"]["nextAfter"]
    return pages


def test_occurrences_page_through_a_repeated_text(catalog):
    built, source = catalog
    # "North" is typed in Data!A2 and Data!A6: one group, read page by page.
    pages = _occurrence_pages(built, source, "north", "Data")
    assert [address for page in pages for address in page] == ["A2", "A6"]


def test_occurrences_of_a_labelled_number_use_its_row_label(catalog):
    built, source = catalog
    # Labels!B2 and Labels!B4 are numbers under "Q1" on rows labelled
    # "North"; their text is the label, read from the row's label column.
    pages = _occurrence_pages(built, source, "q1 north", "Labels")
    assert [address for page in pages for address in page] == ["B2", "B4"]


def test_search_finds_saved_numbers_without_a_label(catalog):
    built, source = catalog
    # Big!A1:A300 has no label: only the scan of saved numbers finds them.
    found = built.search([source], "300")["results"]
    assert [(r["sheet"], r["address"]) for r in found] == [("Big", "A300")]
    found = built.search([source], "40")["results"]
    assert [(r["sheet"], r["address"]) for r in found] == [
        ("Data", "B5"),
        ("Big", "A40"),
    ]


def test_presentation_window_shows_saved_values_within_its_byte_budget(
    project, monkeypatch
):
    from linexcel.excel_presentation import PresentationLimitError, sheet_presentation

    _, _, _, path = project
    window = sheet_presentation(
        path / "workbook.xlsx", path / INDEX_FILENAME, "Data", rows=6, columns=4
    )
    cells = {
        c["address"]: (c["display"], c["valueStatus"], c["formula"])
        for c in window["cells"]
    }
    assert len(cells) == 24
    assert cells["A1"] == ("Region", "present", None)
    assert cells["B5"] == ("40", "present", None)
    assert cells["C3"] == ("FALSE", "present", None)
    assert cells["C5"] == ("#N/A", "present", None)
    # openpyxl saves formulas without a cached value.
    assert cells["D6"] == ("", "missing", "=B6*2")
    monkeypatch.setattr("linexcel.excel_presentation.MAX_PRESENTATION_BYTES", 10)
    with pytest.raises(PresentationLimitError, match="Zone de présentation"):
        sheet_presentation(path / "workbook.xlsx", path / INDEX_FILENAME, "Data")


def test_header_bands_scan_cells_outside_described_tables(project):
    from linexcel.catalog import CatalogLimits
    from linexcel.header_bands import HeaderBands
    from linexcel.lazy_store import WorkbookIndex

    _, _, _, path = project
    table = {
        "name": "tblData",
        "headerRows": 1,
        "headers": ["Region", "Amount", "Flag", "Double"],
        "bounds": [1, 1, 6, 4],
    }
    with WorkbookIndex(path / INDEX_FILENAME) as index:
        labels = HeaderBands(index, "Labels", [], CatalogLimits(), lambda: False)
        alone = HeaderBands(index, "Data", [], CatalogLimits(), lambda: False)
        covered = HeaderBands(index, "Data", [table], CatalogLimits(), lambda: False)
    assert [(b.header_row, b.first_row, b.last_row) for b in labels.bands] == [
        (1, 2, 4)
    ]
    assert labels.bands[0].headers == {1: "Region", 2: "Q1", 3: "Q2"}
    assert labels.bands[0].row_columns == [1]
    assert labels.coverage["cellsScanned"] == 12
    assert alone.coverage["cellsScanned"] == 24 and len(alone.bands) == 1
    # Cells of a table whose headers already describe them are not scanned.
    assert covered.coverage["cellsScanned"] == 0 and covered.bands == []


def test_calculation_step_of_a_wide_range_counts_the_members_it_omits(project):
    from linexcel.lazy import evaluate_node

    _, _, _, path = project
    result = evaluate_node(
        path / "workbook.xlsx", "Calc!B6", index_path=path / INDEX_FILENAME
    )
    assert (result["status"], result["value"]) == ("completed", 45150.0)
    step = next(s for s in result["steps"] if s["nodeId"] == "Big!A1:A300")
    # A step lists 200 members of the 300-cell range and counts the rest.
    assert step["dependencies"] == [f"Big!A{row}" for row in range(1, 201)]
    assert step["dependenciesOmitted"] == 100


@pytest.fixture
def source_data(project, tmp_path):
    import sqlite3

    from linexcel.web_data import build_source_database

    _, _, _, path = project
    output = tmp_path / "source-data.sqlite"
    summary = build_source_database(path / INDEX_FILENAME, output, max_bytes=2**40)
    db = sqlite3.connect(output)
    try:
        yield summary, db
    finally:
        db.close()


def test_data_copy_sizes_each_sheet_grid_from_its_cells(source_data):
    summary, _ = source_data
    assert [
        (s["sheet"], s["table"], s["rowCount"], s["cellCount"], s["columns"])
        for s in summary["sheets"]
    ] == [
        ("Data", "sheet_001", 6, 24, ["A", "B", "C", "D"]),
        ("Calc", "sheet_002", 6, 12, ["A", "B"]),
        ("Big", "sheet_003", 300, 300, ["A"]),
        ("Labels", "sheet_004", 4, 12, ["A", "B", "C"]),
        # The array formula and its three inputs; its members have no cache.
        ("Arr", "sheet_005", 3, 4, ["A", "B"]),
    ]


def test_data_copy_writes_each_cell_with_its_provenance(source_data):
    _, db = source_data
    assert db.execute("SELECT * FROM sheet_001 ORDER BY row").fetchall() == [
        (1, "Region", "Amount", "Flag", "Double"),
        (2, "North", 10, 1, None),
        (3, "South", 20, 0, None),
        (4, "East", 30, 1, None),
        (5, "West", 40, "#N/A", None),
        (6, "North", 50, 1, None),
    ]
    assert db.execute(
        "SELECT cell,value,value_kind,source_type,value_status FROM cells "
        "WHERE sheet='Data' AND row=5 ORDER BY column"
    ).fetchall() == [
        ("A5", "West", "string", "inlineStr", "present"),
        ("B5", 40, "integer", "n", "present"),
        ("C5", "#N/A", "string", "e", "present"),
        ("D5", None, "null", "f", "missing"),
    ]
    assert dict(db.execute("SELECT sheet,count(*) FROM cells GROUP BY sheet")) == {
        "Data": 24,
        "Calc": 12,
        "Big": 300,
        "Labels": 12,
        "Arr": 4,
    }


def test_export_writes_every_node_in_import_order_with_its_context(project):
    from linexcel.web_export import ProjectExport

    app, owner, pid, _ = project
    graph = json.loads(b"".join(ProjectExport(app.store, owner, pid)))["graph"]
    ids = [n["id"] for n in graph["nodes"]]
    # The name, then each sheet's cells as read, then the ranges formulas use.
    assert ids[:9] == ["name:*:Rate"] + [
        f"Data!{c}{r}" for r, cs in ((1, "ABCD"), (2, "ABCD")) for c in cs
    ]
    assert ids[-4:] == ["Data!B2:B6", "Data!A2:A6", "Big!A1:A300", "Arr!B1:B3"]
    assert len(ids) == 357
    nodes = {n["id"]: n for n in graph["nodes"]}
    assert nodes["Data!B5"]["value"] == 40
    assert nodes["Data!D2"]["patternId"] == nodes["Data!D6"]["patternId"]
    # Exported with context: the row label on its left, then the one above.
    assert [label["text"] for label in nodes["Labels!B2"]["cellLabels"]] == [
        "North",
        "Q1",
    ]


def test_import_records_each_external_reference_of_a_formula(tmp_path):
    from test_external import _link
    from test_total_lineage import fixture_source

    from linexcel.lazy_store import WorkbookIndex

    reader = fixture_source(
        tmp_path,
        1,
        "reader.xlsx",
        {"A1": "=SUM('[1]S'!A1:A3)", "A2": "='[1]S'!B1*2", "A3": "=1"},
        links=[_link("inputs.xlsx", "S", {"A1": 500})],
    )
    with WorkbookIndex(reader.path) as index:
        facts = [
            (node, ordinal, json.loads(data))
            for node, ordinal, data in index.db.execute(
                "SELECT node,ordinal,data FROM external_facts ORDER BY node,ordinal"
            )
        ]
        count = index.meta["externalReferenceCount"]
    assert [
        (node, ordinal, f["reference"], f["bounds"], f["placeholder"])
        for node, ordinal, f in facts
    ] == [
        (
            "S!A1",
            0,
            "'[1]S'!A1:A3",
            [1, 1, 3, 1],
            "external-unresolved:inputs.xlsx:S!A1:A3",
        ),
        ("S!A2", 0, "'[1]S'!B1", [1, 2, 1, 2], "external-unresolved:inputs.xlsx:S!B1"),
    ]
    assert count == 2


def test_import_records_the_bounds_of_every_local_range(project):
    from linexcel.lazy_store import WorkbookIndex

    _, _, _, path = project
    with WorkbookIndex(path / INDEX_FILENAME) as index:
        ranges = index.db.execute(
            "SELECT * FROM lineage_ranges ORDER BY rowid"
        ).fetchall()
    assert ranges == [
        ("Data!B2:B6", "data", 2, 2, 6, 2),
        ("Data!A2:A6", "data", 2, 1, 6, 1),
        ("Big!A1:A300", "big", 1, 1, 300, 1),
        ("Arr!B1:B3", "arr", 1, 2, 3, 2),
    ]


@pytest.fixture
def hierarchy(project, tmp_path):
    from linexcel.graph_store import build_graph, read_view

    _, _, _, path = project
    output = tmp_path / "hierarchy.sqlite"
    built = build_graph(path / INDEX_FILENAME, output)

    def view(**where):
        return read_view(output, built["generation"], **where)

    return built, view


def _groups(view):
    return [(n["label"], n["nodeCount"], n.get("level")) for n in view["nodes"]]


def test_graph_sizes_each_formula_group_as_the_index_counts_it(hierarchy):
    built, view = hierarchy
    assert (built["nodeCount"], built["edgeCount"]) == (357, 15)
    # The five copies of =B2*2 form one pattern; the rest of Data are sources.
    assert _groups(view(sheet="Data"))[:2] == [
        ("◇ Data", 21, "sources"),
        ("=RC[-2]*2", 5, "pattern"),
    ]
    assert [g for g in _groups(view(sheet="Calc")) if g[2] == "pattern"] == [
        ("=R[-3]C*Rate", 1, "pattern"),
        ("=SUMIFS(Data!RC:R[4]C,Data!RC[-1]:R[4]C[-1],RC[-1])", 1, "pattern"),
        ("=Data!R[-3]C", 1, "pattern"),
        ("=SUM(Big!R[-5]C[-1]:R[294]C[-1])", 1, "pattern"),
        ("=SUM(Data!R[1]C:R[5]C)", 1, "pattern"),
    ]


def test_graph_places_every_node_under_its_sheet_and_pattern(hierarchy):
    _, view = hierarchy
    root = {n["label"]: n["nodeCount"] for n in view()["nodes"]}
    assert root == {
        "∅": 1,
        "Arr": 5,
        "Big": 301,
        "Calc": 12,
        "Data": 26,
        "Labels": 12,
    }
    pattern = view(node_id="Data!D3")
    assert [n.get("nodeId") for n in pattern["nodes"] if n.get("nodeId")] == [
        f"Data!D{row}" for row in range(2, 7)
    ]


def test_graph_keeps_every_index_edge_between_its_groups(hierarchy):
    _, view = hierarchy
    root = view()
    labels = {n["id"]: n["label"] for n in root["nodes"]}
    assert sorted(
        (labels[e["source"]], labels[e["target"]], e["weight"]) for e in root["edges"]
    ) == [
        ("Big", "Calc", 1),
        ("Calc", "∅", 1),
        ("Data", "Calc", 4),
        ("∅", "Calc", 1),
    ]
    # B2:B6 to D2:D6 in Data, B1 to B4 and A2 to B2 in Calc (B3 reaches B4
    # through the name Rate), the range B1:B3 to A1 in Arr.
    assert {n["label"]: n["internalEdges"] for n in root["nodes"]} == {
        "∅": 0,
        "Arr": 1,
        "Big": 0,
        "Calc": 2,
        "Data": 5,
        "Labels": 0,
    }


def test_report_counts_the_cells_each_selected_range_holds(project):
    from linexcel.web_report import prepare_report

    app, owner, pid, _ = project
    snapshot = prepare_report(
        app.store, owner, pid, {"nodeIds": ["Calc!B6", "Calc!B2"]}
    )
    ranges = {
        record["node"]["id"]: [
            (
                r["id"],
                r["representedCellCount"],
                r["positionCount"],
                len(r["members"]),
                r["omitted"],
            )
            for r in record["ranges"]
        ]
        for record in snapshot["records"]
    }
    assert ranges == {
        "Calc!B6": [("Big!A1:A300", 300, 300, 200, 100)],
        "Calc!B2": [("Data!A2:A6", 5, 5, 5, 0), ("Data!B2:B6", 5, 5, 5, 0)],
    }


def _direct_node_reads(path: Path) -> int:
    """SQL statements in ``path`` that name the index's ``nodes`` table.

    ``graph_store`` keeps a derived table also called ``nodes`` in its own
    hierarchy database, reached through its connection ``db``; there only
    the attached index (``original.nodes``) or the index connection counts.
    """
    import re

    table = re.compile(r"\b(?:FROM|JOIN)\s+(original\.)?nodes\b", re.IGNORECASE)
    count = 0
    for call in ast.walk(ast.parse(path.read_text("utf-8"))):
        if not (
            isinstance(call, ast.Call)
            and isinstance(call.func, ast.Attribute)
            and call.func.attr in {"execute", "executemany", "executescript"}
            and call.args
        ):
            continue
        text = " ".join(
            node.value
            for node in ast.walk(call.args[0])
            if isinstance(node, ast.Constant) and isinstance(node.value, str)
        )
        matches = table.findall(text)
        hierarchy = (
            path.name == "graph_store.py" and ast.unparse(call.func.value) == "db"
        )
        if any(attached or not hierarchy for attached in matches):
            count += 1
    return count


def test_only_workbook_index_reads_the_nodes_table():
    found = {
        path.name: count
        for path in sorted(SOURCE.glob("*.py"))
        if path.name != "lazy_store.py" and (count := _direct_node_reads(path))
    }
    assert found == {}
