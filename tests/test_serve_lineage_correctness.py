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
