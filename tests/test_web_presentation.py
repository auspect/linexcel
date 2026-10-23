"""The worksheet presentation route shares Datasette's owner/source boundary."""

import json
import sqlite3

import pytest
from test_web_data import data_app as data_app
from test_web_data import request

from linexcel.web import _write


def test_presentation_endpoint_is_owner_scoped_and_source_bound(data_app):
    app, project, path = data_app
    route = f"/tools/api/projects/{project}/data/presentation"
    query = "sheet=Data+%26+Source&row=1&column=1&rows=2&columns=4"
    assert request(app, route, query=query, user="bob")[0] == 404
    status, headers, body = request(app, route, query=query)
    assert status == 200 and headers["Cache-Control"] == "no-store"
    view = json.loads(body)
    assert view["available"]
    assert view["window"] == {"row": 1, "column": 1, "rows": 2, "columns": 4}
    cells = {cell["address"]: cell for cell in view["cells"]}
    assert cells["B1"]["value"] == 6 and cells["B1"]["formula"] == "=2*3"
    assert cells["C1"]["value"] is True
    assert cells["D1"]["display"] == "literal <script>"
    (path / "workbook.xlsx").write_bytes(b"changed")
    assert request(app, route, query=query)[0] == 409


@pytest.mark.parametrize(
    "query",
    [
        "sheet=Data+%26+Source&row=zero",
        "sheet=Data+%26+Source&row=0",
        "sheet=Data+%26+Source&column=16385",
    ],
)
def test_presentation_invalid_window_returns_a_local_error(data_app, query):
    app, project, _ = data_app
    route = f"/tools/api/projects/{project}/data/presentation"
    status, _, body = request(app, route, query=query)
    assert status == 400 and json.loads(body)["error"]


def test_presentation_window_limits_and_unknown_sheet(data_app):
    app, project, _ = data_app
    route = f"/tools/api/projects/{project}/data/presentation"
    status, _, body = request(
        app, route, query="sheet=Empty&rows=999999&columns=999999"
    )
    view = json.loads(body)
    assert status == 200 and view["window"]["rows"] == 100
    assert view["window"]["columns"] == 40 and not view["cells"]
    assert request(app, route, query="sheet=Unknown")[0] == 404


def test_presentation_does_not_build_or_require_datasette_snapshot(
    data_app, monkeypatch
):
    app, project, path = data_app
    route = f"/tools/api/projects/{project}/data/presentation"

    def forbid_snapshot(*args, **kwargs):
        raise AssertionError("A bounded presentation must not build the full snapshot")

    monkeypatch.setattr(app.data_browser, "_snapshot", forbid_snapshot)
    status, _, _ = request(app, route, query="sheet=Data+%26+Source")
    assert status == 200
    assert not (path / "source-data.sqlite").exists()
    assert not (path / "source-data.json").exists()


def test_presentation_rejects_stale_index_hash_even_with_current_project_revision(
    data_app,
):
    app, project, path = data_app
    with sqlite3.connect(path / "index.sqlite") as database:
        raw = database.execute(
            "SELECT value FROM metadata WHERE key='meta'"
        ).fetchone()[0]
        metadata = json.loads(raw)
        metadata["sha256"] = "0" * 64
        database.execute(
            "UPDATE metadata SET value=? WHERE key='meta'", (json.dumps(metadata),)
        )
    route = f"/tools/api/projects/{project}/data/presentation"
    assert request(app, route, query="sheet=Data+%26+Source")[0] == 409


def test_presentation_macro_and_missing_index_remain_local_errors(data_app):
    app, project, path = data_app
    route = f"/tools/api/projects/{project}/data/presentation"
    metadata = json.loads((path / "project.json").read_text())
    metadata["filename"] = "tool.xlam"
    _write(path / "project.json", metadata)
    assert request(app, route, query="sheet=Data+%26+Source")[0] == 409
    metadata["filename"] = "sources.xlsx"
    _write(path / "project.json", metadata)
    (path / "index.sqlite").unlink()
    assert request(app, route, query="sheet=Data+%26+Source")[0] == 409
