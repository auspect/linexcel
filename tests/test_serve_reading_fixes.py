"""Local corrections from the batch 2 measurements (batch 2 bis).

Each test reproduces a finding of ``docs/critical_review.md`` section 9 on the
``linexcel serve`` path: through the HTTP coordinator where the behaviour is
visible there, otherwise through the function the worker or the coordinator
calls.
"""

import base64
import datetime as dt
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


def finish(app, task_id, seconds=120):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        owner = app.store.tasks[task_id]["owner"]
        task = app.store.get_task(owner, task_id)
        if task["status"] in {"failed", "succeeded", "cancelled"}:
            return task
        time.sleep(0.02)
    pytest.fail("worker did not finish")


def workbook_bytes(book):
    stream = io.BytesIO()
    book.save(stream)
    return stream.getvalue()


@pytest.fixture
def serve(tmp_path):
    app = create_app(tmp_path, resolve_user=lambda env: env.get("user"))

    def upload(data, name="book.xlsx"):
        if isinstance(data, Workbook):
            data = workbook_bytes(data)
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


# F07, F19: range inputs reach the engine in row chunks, with the same values.

ROWS = 25_000


def long_table_book():
    """A 25,000-row table read by ranges, with a formula inside one column,
    gaps, text, booleans, dates and errors."""
    book = Workbook()
    data = book.active
    data.title = "Data"
    for row in range(1, ROWS + 1):
        if row == 7:
            data.cell(row, 1, "=40+2")  # a formula inside the summed column
        elif row % 1000 != 0:  # one gap every 1,000 rows
            data.cell(row, 1, row)
        data.cell(row, 2, "x" if row % 3 == 0 else "y")
        data.cell(row, 3, row % 2 == 0)
    data["D1"] = dt.date(2024, 1, 31)
    data["D2"] = dt.date(2024, 2, 29)
    data["E1"] = "#DIV/0!"
    data["E1"].data_type = "e"
    calc = book.create_sheet("Calc")
    calc["A1"] = f"=SUM(Data!A1:A{ROWS})"
    calc["A2"] = f'=COUNTIF(Data!B1:B{ROWS},"x")'
    calc["A3"] = f"=SUMPRODUCT(--Data!C1:C{ROWS})"
    calc["A4"] = "=SUM(Data!D1:D2)"
    calc["A5"] = "=ISERROR(Data!E1)"
    return book


def expected_long_table():
    numbers = sum(r for r in range(1, ROWS + 1) if r != 7 and r % 1000 != 0)
    return {
        "Calc!A1": numbers + 42,
        "Calc!A2": ROWS // 3,
        "Calc!A3": ROWS // 2,
        "Calc!A4": 45322 + 45351,
        "Calc!A5": True,
    }


def test_range_inputs_keep_their_values_through_the_serve_path(serve):
    app, upload = serve
    project, task = upload(long_table_book())
    assert task["status"] == "succeeded", task
    for node, value in expected_long_table().items():
        result = evaluate(app, project, node)
        assert result["status"] == "completed", (node, result["diagnostics"])
        assert result["value"] == value, node


class RecordingWorkbook:
    """Wraps formualizer.Workbook and records how values reach it."""

    calls: list = []
    native_class = None

    def __init__(self, *args, **kwargs):
        self._native = RecordingWorkbook.native_class(*args, **kwargs)

    def set_value(self, sheet, row, column, value):
        RecordingWorkbook.calls.append(("set_value", sheet, row, column, 1))
        return self._native.set_value(sheet, row, column, value)

    def set_values_batch(self, sheet, row, column, data):
        RecordingWorkbook.calls.append(
            ("set_values_batch", sheet, row, column, len(data), len(data[0]))
        )
        return self._native.set_values_batch(sheet, row, column, data)

    def __getattr__(self, name):
        return getattr(self._native, name)


def test_range_inputs_are_loaded_in_bounded_chunks_never_over_a_formula(
    tmp_path, monkeypatch
):
    import formualizer

    from linexcel.lazy import evaluate_node
    from linexcel.lazy_store import build_index

    path = tmp_path / "long.xlsx"
    path.write_bytes(workbook_bytes(long_table_book()))
    index = tmp_path / "long.sqlite"
    build_index(path, index)
    RecordingWorkbook.native_class = formualizer.Workbook
    RecordingWorkbook.calls = []
    monkeypatch.setattr(formualizer, "Workbook", RecordingWorkbook)

    result = evaluate_node(path, "Calc!A1", index_path=index)

    assert result["status"] == "completed", result["diagnostics"]
    assert result["value"] == expected_long_table()["Calc!A1"]
    data_calls = [c for c in RecordingWorkbook.calls if c[1] == "Data"]
    per_cell = [c for c in data_calls if c[0] == "set_value"]
    batches = [c for c in data_calls if c[0] == "set_values_batch"]
    loaded = sum(c[4] for c in data_calls)
    assert loaded == ROWS - 1 - ROWS // 1000  # every input, no formula, no gap
    assert len(per_cell) <= 10, "single isolated inputs only"
    assert batches and all(rows <= 10_000 for *_, rows, _width in batches)
    for _, _, row, column, rows, *_ in data_calls:
        assert column == 1
        assert not row <= 7 < row + rows, "a batch must not write over A7"


def test_chunked_loading_matches_cell_by_cell_loading():
    """Differential check of the loader against the per-cell path."""
    import formualizer as fz

    from linexcel.lazy import _INPUT_ERROR_KINDS, _load_inputs

    inputs = []
    for row in range(1, 2501):
        if row != 100:  # S!A100 holds a formula, never an input
            inputs.append({"sheet": "S", "row": row, "column": 1, "value": row})
        if row % 7:
            inputs.append({"sheet": "S", "row": row, "column": 2, "value": row / 3})
    inputs += [
        {"sheet": "S", "row": 1, "column": 3, "value": "text"},
        {"sheet": "S", "row": 2, "column": 3, "value": ""},
        {"sheet": "S", "row": 3, "column": 3, "value": True},
        {"sheet": "S", "row": 4, "column": 3, "value": False},
        {"sheet": "S", "row": 5, "column": 3, "value": None},
        {"sheet": "S", "row": 6, "column": 3, "value": 2**40},
        {
            "sheet": "S",
            "row": 7,
            "column": 3,
            "value": "2024-01-31T00:00:00",
            "cached_comparison_value": 45322,
        },
    ]
    for offset, text in enumerate(_INPUT_ERROR_KINDS, 8):
        inputs.append(
            {
                "sheet": "S",
                "row": offset,
                "column": 3,
                "value": text,
                "sourceDataType": "e",
            }
        )
    inputs.append({"sheet": "Ext", "row": 2, "column": 2, "value": 5})
    inputs.sort(key=lambda n: (n["sheet"], n["row"], n["column"]))
    alias = {"Ext": "_linexcel_ext"}

    def engine():
        workbook = fz.Workbook()
        for sheet in ("S", "_linexcel_ext"):
            workbook.add_sheet(sheet)
        workbook.set_formula("S", 100, 1, "=1")  # inside a column, never loaded
        return workbook

    reference = engine()
    for node in inputs:
        value = node.get("cached_comparison_value", node.get("value"))
        if node.get("sourceDataType") == "e":
            value = {"type": "Error", "kind": _INPUT_ERROR_KINDS[value]}
        if value is not None:
            reference.set_value(
                alias.get(node["sheet"], node["sheet"]),
                node["row"],
                node["column"],
                value,
            )
    chunked = engine()
    _load_inputs(chunked, iter(inputs), alias, chunk_rows=1000)
    for sheet, rows, columns in (("S", 2501, 3), ("_linexcel_ext", 3, 3)):
        for row in range(1, rows):
            for column in range(1, columns + 1):
                assert repr(chunked.get_value(sheet, row, column)) == repr(
                    reference.get_value(sheet, row, column)
                ), (sheet, row, column)


# F26: reading a small selection must not walk the whole index.


def test_reading_a_small_selection_does_not_scan_the_index(tmp_path):
    from linexcel.lazy_store import WorkbookIndex, build_index

    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    for row in range(1, 30_001):
        sheet.cell(row, 1, row)
    book.create_sheet("Small")["A1"] = 1
    path = tmp_path / "wide.xlsx"
    path.write_bytes(workbook_bytes(book))
    index = tmp_path / "wide.sqlite"
    build_index(path, index)
    with WorkbookIndex(index) as store:
        store.prepare_input_selection()
        for node_id in ("Data!A5", "Data!A29999", "Small!A1"):
            store.select_input(node_id)
        steps = []
        store.db.set_progress_handler(lambda: steps.append(1), 100)
        selected = [node["id"] for node in store.selected_inputs()]
        store.db.set_progress_handler(None, 100)
    assert selected == ["Data!A5", "Data!A29999", "Small!A1"]
    # About 100 virtual-machine steps per callback: a walk of the 30,000-entry
    # position index costs thousands; three lookups and a sort cost a few.
    assert len(steps) < 50, len(steps)
