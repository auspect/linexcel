"""The default time budget of an import is a setting of the server (batch 2 ter).

With the default 600 s an import may read sheets for 360 s; the public DGT
election table needs about 500 s. ``linexcel serve --import-seconds`` lets a
server that receives such workbooks raise the default of imports, graph
constructions and reports, which every user can still change per operation.
"""

import base64
import hashlib
import io
import json

import pytest
from openpyxl import Workbook

from linexcel.cli import _build_parser
from linexcel.web import create_app


def call(app, method, path, body=None):
    data = json.dumps(body or {}).encode()
    status = []
    result = b"".join(
        app(
            {
                "REQUEST_METHOD": method,
                "PATH_INFO": path,
                "QUERY_STRING": "",
                "CONTENT_TYPE": "application/json",
                "CONTENT_LENGTH": str(len(data)),
                "wsgi.input": io.BytesIO(data),
                "user": "alice",
            },
            lambda code, headers: status.append(code),
        )
    )
    return int(status[0].split()[0]), json.loads(result)


def workbook():
    book = Workbook()
    book.active["A1"] = 1
    stream = io.BytesIO()
    book.save(stream)
    return base64.b64encode(stream.getvalue()).decode()


@pytest.mark.parametrize("setting", [None, 1500])
def test_imports_without_a_budget_use_the_servers_default(tmp_path, setting):
    options = {} if setting is None else {"import_seconds": setting}
    app = create_app(tmp_path, resolve_user=lambda env: env.get("user"), **options)
    expected = 600 if setting is None else setting
    try:
        code, config = call(app, "GET", "/api/config")
        assert code == 200 and config["limits"]["importSeconds"] == expected
        code, created = call(
            app,
            "POST",
            "/api/projects",
            {"workbook": {"name": "a.xlsx", "data": workbook()}},
        )
        assert code == 200, created
        assert created["task"]["budget"]["seconds"] == expected
        # A file uploaded in parts, here through a local import, the same.
        path = tmp_path / "b.xlsx"
        path.write_bytes(base64.b64decode(workbook()))
        owner = hashlib.sha256(b"alice").hexdigest()
        uploaded = app.store.uploads.local(owner, {"path": str(path)})
        assert uploaded["task"]["budget"]["seconds"] == expected
        # A budget chosen for the operation still wins.
        code, chosen = call(
            app,
            "POST",
            "/api/projects",
            {
                "workbook": {"name": "c.xlsx", "data": workbook()},
                "budget": {"seconds": 90},
            },
        )
        assert code == 200 and chosen["task"]["budget"]["seconds"] == 90
    finally:
        app.close()


def test_the_serve_command_sets_the_default_import_budget():
    assert _build_parser().parse_args(["serve"]).import_seconds == 600
    args = _build_parser().parse_args(["serve", "--import-seconds", "1500"])
    assert args.import_seconds == 1500


@pytest.mark.parametrize("seconds", [0, 3601])
def test_the_default_import_budget_stays_within_the_operation_bounds(tmp_path, seconds):
    with pytest.raises(ValueError, match="import_seconds"):
        create_app(tmp_path, resolve_user=lambda env: "alice", import_seconds=seconds)
