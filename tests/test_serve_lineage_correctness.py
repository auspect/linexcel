"""Lineage correctness on the ``linexcel serve`` path (critical review, batch 1 bis).

Each test reproduces a verified finding of ``docs/critical_review.md`` through
the HTTP coordinator, the same path the application uses.
"""

import base64
import io
import json
import time

import pytest
from openpyxl import Workbook

from linexcel.web import create_app


def call(app, method, path, body=None, query="", user="alice"):
    data = json.dumps(body or {}).encode()
    response = []
    result = b"".join(
        app(
            {
                "REQUEST_METHOD": method,
                "PATH_INFO": path,
                "QUERY_STRING": query,
                "CONTENT_TYPE": "application/json",
                "CONTENT_LENGTH": str(len(data)),
                "wsgi.input": io.BytesIO(data),
                "user": user,
            },
            lambda status, headers: response.append(status),
        )
    )
    return int(response[0].split()[0]), json.loads(result)


def finish(app, task_id):
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        owner = app.store.tasks[task_id]["owner"]
        task = app.store.get_task(owner, task_id)
        if task["status"] in {"failed", "succeeded", "cancelled"}:
            return task
        time.sleep(0.02)
    pytest.fail("worker did not finish")


@pytest.fixture
def serve(tmp_path):
    app = create_app(tmp_path, resolve_user=lambda env: env.get("user"))

    def upload(data, name="book.xlsx"):
        if isinstance(data, Workbook):
            stream = io.BytesIO()
            data.save(stream)
            data = stream.getvalue()
        code, imported = call(
            app,
            "POST",
            "/api/projects",
            {"workbook": {"name": name, "data": base64.b64encode(data).decode()}},
        )
        assert code == 200, imported
        task = finish(app, imported["task"]["id"])
        return imported["project"]["id"], task

    try:
        yield app, upload
    finally:
        app.close()


def neighborhood(app, project, node, depth=1, limit=200):
    code, around = call(
        app,
        "GET",
        f"/api/projects/{project}/neighborhood",
        query=f"nodeId={node.replace('!', '%21')}&depth={depth}&limit={limit}",
    )
    assert code == 200, around
    return around


def evaluate(app, project, node):
    code, submitted = call(
        app,
        "POST",
        f"/api/projects/{project}/tasks",
        {"operation": "evaluate", "nodeId": node},
    )
    assert code == 200, submitted
    task = finish(app, submitted["task"]["id"])
    assert task["status"] == "succeeded", task
    code, result = call(app, "GET", f"/api/tasks/{submitted['task']['id']}/export")
    assert code == 200
    return result


# F02: an input read only through a range must appear as used.


def range_book():
    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    for row in range(2, 12):
        sheet.cell(row, 3, row)
    sheet["I1"] = "=SUM(C2:C11)"
    sheet["J1"] = "=C5*2"
    return book


def test_cell_read_through_a_range_has_the_range_and_its_consumer_as_usages(serve):
    app, upload = serve
    project, task = upload(range_book())
    assert task["status"] == "succeeded"

    direct = neighborhood(app, project, "S!C6")
    assert {n["id"] for n in direct["nodes"]} == {"S!C6", "S!C2:C11"}
    assert {"source": "S!C6", "target": "S!C2:C11", "kind": "range"} in direct["edges"]
    assert not direct["pagination"]["truncated"]

    two = neighborhood(app, project, "S!C5", depth=2)
    ids = {n["id"] for n in two["nodes"]}
    assert {"S!C5", "S!J1", "S!C2:C11", "S!I1"} <= ids
    assert {"source": "S!C2:C11", "target": "S!I1", "kind": "dep"} in two["edges"]
    assert {"source": "S!C5", "target": "S!C2:C11", "kind": "range"} in two["edges"]
    # Every edge stays inside the returned nodes.
    assert all(e["source"] in ids and e["target"] in ids for e in two["edges"])


def test_range_neighborhood_lists_its_members_and_reports_truncation(serve):
    app, upload = serve
    project, _ = upload(range_book())

    whole = neighborhood(app, project, "S!C2:C11")
    members = {f"S!C{row}" for row in range(2, 12)}
    assert {n["id"] for n in whole["nodes"]} == {"S!C2:C11", "S!I1"} | members
    assert not whole["pagination"]["truncated"]

    cut = neighborhood(app, project, "S!C2:C11", limit=4)
    assert cut["pagination"]["returned"] == len(cut["nodes"]) == 4
    assert cut["pagination"]["truncated"]


def test_index_without_range_table_still_reports_membership(tmp_path):
    import sqlite3

    from linexcel.lazy_store import build_index, graph_neighborhood

    source = tmp_path / "ranges.xlsx"
    range_book().save(source)
    index = tmp_path / "ranges.sqlite"
    build_index(source, index)
    with sqlite3.connect(index) as db:  # An index written before lineage_ranges.
        db.execute("DROP TABLE lineage_ranges")
    around = graph_neighborhood(index, "S!C6", depth=2)
    assert {n["id"] for n in around["nodes"]} >= {"S!C6", "S!C2:C11", "S!I1"}
    assert {"source": "S!C6", "target": "S!C2:C11", "kind": "range"} in around["edges"]


# F03: one failing reference must not erase the references after it.


def node_detail(app, project, node):
    code, detail = call(
        app,
        "GET",
        f"/api/projects/{project}/node",
        query=f"nodeId={node.replace('!', '%21')}",
    )
    assert code == 200, detail
    return detail["node"]


def failing_reference_book():
    book = Workbook()
    first = book.active
    first.title = "A"
    for title in ("B", "C"):
        book.create_sheet(title)
    for sheet in book.worksheets:
        for row in range(1, 4):
            sheet.cell(row, 1, row)
    first["C2"] = "=B!A1+A:B!A1+C!A2"  # 3D reference outside an aggregate.
    first["C3"] = "=A1#+C!A3"  # Spill reference without a declared array.
    return book


def test_failing_reference_keeps_later_references_and_names_its_cause(serve):
    app, upload = serve
    project, _ = upload(failing_reference_book())

    span = node_detail(app, project, "A!C2")
    assert span["dependencies"] == ["B!A1", "C!A2"]
    assert any(
        "A:B!A1" in d and "3D references require" in d for d in span["diagnostics"]
    )
    assert not any("tokenization failed" in d for d in span["diagnostics"])
    assert {(c["feature"], c["reference"]) for c in span["capabilityDiagnostics"]} == {
        ("three_dimensional_references", "A:B!A1")
    }

    spill = node_detail(app, project, "A!C3")
    assert spill["dependencies"] == ["C!A3"]
    assert any("A1#" in d and "declared array" in d for d in spill["diagnostics"])
    assert {(c["feature"], c["reference"]) for c in spill["capabilityDiagnostics"]} == {
        ("arrays_and_spills", "A1#")
    }

    # The later source keeps its usage edge.
    around = neighborhood(app, project, "C!A2")
    assert {"source": "C!A2", "target": "A!C2", "kind": "dep"} in around["edges"]


def test_partially_extracted_formula_is_still_refused(serve):
    app, upload = serve
    project, _ = upload(failing_reference_book())
    result = evaluate(app, project, "A!C2")
    assert result["status"] == "unsupported"
    assert result.get("value") is None
    assert any("A:B!A1" in d for d in result["diagnostics"])


# F04: LET and LAMBDA names are local to their formula, not references.


def let_book():
    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    sheet["A1"] = 2
    sheet["A2"] = 5
    sheet["B1"] = "=LET(x,A1,x*3)"
    # The spelling Excel writes to the file.
    sheet["B2"] = "=_xlfn.LET(_xlpm.x,A1,_xlpm.x*3)"
    sheet["B3"] = "=_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.v,_xlpm.v+1),_xlpm.f(A1))"
    sheet["B4"] = "=LET(total,SUM(A1:A2),n,COUNT(A1:A2),total/n)"
    sheet["B5"] = "=LET(x,A1,x+missing)"  # A genuinely unknown name.
    return book


@pytest.mark.parametrize(
    ("cell", "dependencies", "value"),
    [
        ("S!B1", ["S!A1"], 6),
        ("S!B2", ["S!A1"], 6),
        ("S!B3", ["S!A1"], 3),
        ("S!B4", ["S!A1:A2"], 3.5),
    ],
)
def test_let_names_are_not_references_and_the_cell_is_calculated(
    serve, cell, dependencies, value
):
    app, upload = serve
    project, _ = upload(let_book())
    node = node_detail(app, project, cell)
    assert node["dependencies"] == dependencies
    assert node["diagnostics"] == []
    result = evaluate(app, project, cell)
    assert result["status"] == "completed", result["diagnostics"]
    assert result["value"] == value


def test_unknown_name_inside_let_is_still_refused(serve):
    app, upload = serve
    project, _ = upload(let_book())
    node = node_detail(app, project, "S!B5")
    assert "unresolved:S:missing" in node["dependencies"]
    assert not any(dep.endswith(":x") for dep in node["dependencies"])
    assert evaluate(app, project, "S!B5")["status"] == "unsupported"


# F09: a damaged workbook-level part degrades locally and visibly.


def damaged(book, edit):
    import zipfile

    stream = io.BytesIO()
    book.save(stream)
    out = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(stream.getvalue())) as source,
        zipfile.ZipFile(out, "w") as target,
    ):
        for item in source.infolist():
            target.writestr(item.filename, edit(item.filename, source.read(item)))
    return out.getvalue()


def resilient_book():
    import datetime

    from openpyxl.workbook.defined_name import DefinedName

    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    sheet["A1"] = 2
    sheet["A2"] = datetime.date(2024, 1, 31)
    sheet["B1"] = "=A1*10"
    sheet["B2"] = "=Rate*A1"
    other = book.create_sheet("T")
    other["A1"] = 4
    book.defined_names["Rate"] = DefinedName("Rate", attr_text="T!$A$1")
    return book


def project_meta(app, project):
    code, snapshot = call(app, "GET", f"/api/projects/{project}")
    assert code == 200, snapshot
    return snapshot["graph"]["meta"]


def test_truncated_styles_part_keeps_the_workbook_usable(serve):
    app, upload = serve
    data = damaged(
        resilient_book(),
        lambda name, body: body[: len(body) // 2] if name == "xl/styles.xml" else body,
    )
    project, task = upload(data)
    assert task["status"] == "succeeded", task
    meta = project_meta(app, project)
    assert meta["status"] == "partial"
    assert [s["status"] for s in meta["sheetDetails"]] == ["complete", "complete"]
    styles = [d for d in meta["diagnostics"] if d.get("part") == "xl/styles.xml"]
    assert len(styles) == 1 and "number formats are unknown" in styles[0]["message"]
    # A styled cell keeps an unknown value, never a guessed number or date.
    dated = node_detail(app, project, "S!A2")
    assert dated["value"] is None and dated["valueSource"] == "unknown"
    assert dated["diagnostics"]
    assert node_detail(app, project, "S!A1")["value"] == 2
    result = evaluate(app, project, "S!B1")
    assert result["status"] == "completed" and result["value"] == 20


def test_malformed_defined_name_is_skipped_locally(serve):
    app, upload = serve
    data = damaged(
        resilient_book(),
        lambda name, body: (
            body.replace(
                b'<definedName name="Rate"',
                b'<definedName name="Rate" localSheetId="x"',
            )
            if name == "xl/workbook.xml"
            else body
        ),
    )
    project, task = upload(data)
    assert task["status"] == "succeeded", task
    meta = project_meta(app, project)
    assert meta["status"] == "partial"
    names = [d for d in meta["diagnostics"] if d.get("part") == "xl/workbook.xml"]
    assert len(names) == 1 and "Rate" in names[0]["message"]
    assert "localSheetId" in names[0]["message"]
    assert evaluate(app, project, "S!B1")["value"] == 20
    # A formula reading the skipped name is refused, never computed from a guess.
    refused = evaluate(app, project, "S!B2")
    assert refused["status"] == "unsupported"
    assert any("Rate" in d for d in refused["diagnostics"])
