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
    with WorkbookIndex(path / "index.sqlite") as index:
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


# Direct reads of the index's ``nodes`` table still to route, by module. The
# list only shrinks; it is empty once every read goes through WorkbookIndex.
REMAINING_DIRECT_READS = {
    "catalog_build.py": 3,
    "catalog_lineage.py": 3,
    "catalog_occurrences.py": 2,
    "catalog_query.py": 1,
    "excel_presentation.py": 1,
    "graph_store.py": 3,
    "header_bands.py": 1,
    "lazy.py": 1,
    "link_facts.py": 2,
    "web_data.py": 2,
    "web_export.py": 1,
    "web_report.py": 1,
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
    assert found == REMAINING_DIRECT_READS
