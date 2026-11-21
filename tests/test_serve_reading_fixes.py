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


# F24: a reader of the index holds one read transaction.


def test_an_index_reader_reads_one_snapshot_of_the_index(tmp_path):
    """Each serve read path opens a read-only WorkbookIndex. It must read one
    snapshot: a statement-by-statement reader takes the file lock for every
    lookup (about 270 µs each on Windows) and can mix two states of a file."""
    import sqlite3

    from linexcel.lazy_store import WorkbookIndex, build_index, graph_neighborhood

    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    sheet["A1"] = 1
    sheet["B1"] = "=A1*2"
    path = tmp_path / "small.xlsx"
    path.write_bytes(workbook_bytes(book))
    index = tmp_path / "small.sqlite"
    build_index(path, index)
    with WorkbookIndex(index) as reader:
        first = reader.db.execute("SELECT count(*) FROM edges").fetchone()[0]
        writer = sqlite3.connect(index, timeout=0)
        try:
            writer.execute("DELETE FROM edges")
            writer.commit()
            committed = True
        except sqlite3.OperationalError:
            committed = False
        finally:
            writer.close()
        second = reader.db.execute("SELECT count(*) FROM edges").fetchone()[0]
    assert first == second == 1
    assert not committed, "a writer must not change the file under a reader"
    around = graph_neighborhood(index, "S!B1", depth=1)
    assert {node["id"] for node in around["nodes"]} == {"S!A1", "S!B1"}


# F17: a part the workbook declares but the package lacks is a workbook fact.


def without_part(data, part):
    import zipfile

    source = zipfile.ZipFile(io.BytesIO(data))
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            if info.filename != part:
                target.writestr(info, source.read(info.filename))
    return stream.getvalue()


def text_book():
    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    sheet["A1"] = "Region"
    sheet["A2"] = "North"
    sheet["B2"] = 42
    sheet["B2"].number_format = "0.00"
    return workbook_bytes(book)


def snapshot_meta(app, project):
    code, snapshot = call(app, "GET", f"/api/projects/{project}")
    assert code == 200, snapshot
    return (snapshot.get("graph") or {}).get("meta") or snapshot.get("meta") or {}


def node(app, project, node_id):
    code, found = call(
        app, "GET", f"/api/projects/{project}/node", query="nodeId=" + node_id
    )
    assert code == 200, found
    return found.get("node", found)


def shared_string_book(with_part=True):
    """The text book with its strings moved to a shared-string table. The
    relationship and content type always declare the part; ``with_part``
    decides whether the part itself is in the package."""
    import zipfile

    source = zipfile.ZipFile(io.BytesIO(text_book()))
    stream = io.BytesIO()
    strings = (
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'count="2" uniqueCount="2"><si><t>Region</t></si><si><t>North</t></si></sst>'
    )
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            text = source.read(info.filename).decode("utf-8")
            if info.filename == "xl/worksheets/sheet1.xml":
                for index, value in enumerate(("Region", "North")):
                    text = text.replace(
                        f't="inlineStr"><is><t>{value}</t></is></c>',
                        f't="s"><v>{index}</v></c>',
                    )
                assert "inlineStr" not in text
            elif info.filename == "xl/_rels/workbook.xml.rels":
                text = text.replace(
                    "</Relationships>",
                    '<Relationship Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/sharedStrings" '
                    'Target="sharedStrings.xml" Id="rIdSst" /></Relationships>',
                )
            elif info.filename == "[Content_Types].xml":
                text = text.replace(
                    "</Types>",
                    '<Override PartName="/xl/sharedStrings.xml" ContentType="'
                    "application/vnd.openxmlformats-officedocument.spreadsheetml."
                    'sharedStrings+xml"/></Types>',
                )
            target.writestr(info, text)
        if with_part:
            target.writestr("xl/sharedStrings.xml", strings)
    return stream.getvalue()


def test_shared_string_control_book_imports_completely(serve):
    app, upload = serve
    project, task = upload(shared_string_book(with_part=True))
    assert task["status"] == "succeeded", task
    assert snapshot_meta(app, project)["status"] == "complete"
    assert node(app, project, "S!A2")["value"] == "North"


def test_missing_shared_strings_part_makes_the_import_partial(serve):
    app, upload = serve
    project, task = upload(shared_string_book(with_part=False))
    assert task["status"] == "succeeded", task
    meta = snapshot_meta(app, project)
    assert meta["status"] == "partial"
    messages = [
        d for d in meta["diagnostics"] if d.get("part") == "xl/sharedStrings.xml"
    ]
    assert messages and "missing" in messages[0]["message"]
    assert "2 text cell" in messages[0]["message"]
    cell = node(app, project, "S!A2")
    assert cell["value"] is None and cell["diagnostics"], "unreadable, never empty"


def test_missing_styles_part_makes_the_import_partial(serve):
    app, upload = serve
    project, task = upload(without_part(text_book(), "xl/styles.xml"))
    assert task["status"] == "succeeded", task
    meta = snapshot_meta(app, project)
    assert meta["status"] == "partial"
    assert any(
        d.get("part") == "xl/styles.xml" and "missing" in d["message"]
        for d in meta["diagnostics"]
    )
    assert node(app, project, "S!B2")["value"] is None


# F22: an operation stopped by the storage quota, or a missing part, says so.


def test_a_storage_quota_stop_says_what_happened(tmp_path):
    app = create_app(
        tmp_path, resolve_user=lambda env: env.get("user"), max_storage_mb=1
    )
    try:
        book = Workbook()
        sheet = book.active
        # Long formulas stay one record each (literals would fit in a grid);
        # under 10,000 cells, the import's own budget check never runs inside
        # the sheet, so the coordinator's quota stop ends the task (F21 only
        # publishes sheets read within the budget).
        for row in range(1, 9001):
            sheet.cell(row, 1, "=" + "+".join([str(row)] * 150))
        data = workbook_bytes(book)
        assert len(data) < 512 * 1024  # the upload fits, the index does not
        code, imported = call(
            app,
            "POST",
            "/api/projects",
            {"workbook": {"name": "big.xlsx", "data": base64.b64encode(data).decode()}},
        )
        assert code == 200, imported
        task = finish(app, imported["task"]["id"])
        assert task["status"] == "failed"
        assert task["error"]["kind"] == "storage_limit"
        message = task["error"]["message"]
        assert "Quota de stockage" in message and "--max-storage-mb" in message
        project = imported["project"]["id"]
        code, nodes = call(app, "GET", f"/api/projects/{project}/nodes")
        assert code == 409
        assert "le dernier import n’a pas abouti" in nodes["error"]["message"]
        # The Data tab, the only tab that asks the server, says the same.
        code, data = call(app, "GET", f"/api/projects/{project}/data")
        assert code == 409
        assert data["error"]["message"] == nodes["error"]["message"]
    finally:
        app.close()


def two_sheet_book():
    book = Workbook()
    first = book.active
    first.title = "First"
    first["A1"] = 1
    second = book.create_sheet("Second")
    for row in range(1, 50):
        second.cell(row, 1, row)
    return workbook_bytes(book)


def presentation(app, project, sheet):
    return call(
        app,
        "GET",
        f"/api/projects/{project}/data/presentation",
        query=f"sheet={sheet}&row=1&column=1&rows=20&columns=5",
    )


def test_a_missing_sheet_part_is_named_in_the_presentation(serve):
    app, upload = serve
    project, task = upload(without_part(two_sheet_book(), "xl/worksheets/sheet2.xml"))
    assert task["status"] == "succeeded", task
    code, first = presentation(app, project, "First")
    assert code == 200, first
    code, second = presentation(app, project, "Second")
    assert code == 409, second
    assert "absente du fichier" in second["error"]["message"]


def test_a_partly_read_sheet_is_named_in_the_presentation(serve):
    import zipfile

    app, upload = serve
    source = zipfile.ZipFile(io.BytesIO(two_sheet_book()))
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet2.xml":
                data = data[: len(data) // 2]
            target.writestr(info, data)
    project, task = upload(stream.getvalue())
    assert task["status"] == "succeeded", task
    code, second = presentation(app, project, "Second")
    assert code == 409, second
    assert "lue que partiellement" in second["error"]["message"]


# F25: the presentation streams styles.xml and bounds what it keeps.


def with_styles(data, section, entries):
    """Copy a workbook with ``entries`` appended to one styles.xml section."""
    import zipfile

    source = zipfile.ZipFile(io.BytesIO(data))
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            content = source.read(info.filename)
            if info.filename == "xl/styles.xml":
                text = content.decode("utf-8")
                closing = f"</{section}>"
                assert closing in text, section
                text = text.replace(closing, entries + closing, 1)
                content = text.encode("utf-8")
            target.writestr(info, content)
    return stream.getvalue()


def bold_book():
    from openpyxl.styles import Font

    book = Workbook()
    sheet = book.active
    sheet.title = "S"
    sheet["A1"] = "Total"
    sheet["A1"].font = Font(bold=True)
    sheet["B1"] = 3
    return workbook_bytes(book)


def test_named_styles_beyond_4_mib_do_not_block_the_presentation(serve):
    app, upload = serve
    named = (
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">'
        '<alignment horizontal="left" vertical="center" wrapText="1"/></xf>'
    ) * 45_000
    data = with_styles(bold_book(), "cellStyleXfs", named)
    project, task = upload(data)
    assert task["status"] == "succeeded", task
    code, window = presentation(app, project, "S")
    assert code == 200, window
    a1 = next(cell for cell in window["cells"] if cell["address"] == "A1")
    assert window["styles"][str(a1["styleId"])]["css"].get("fontWeight") == "bold"


def test_too_many_cell_formats_are_still_refused(serve):
    app, upload = serve
    formats = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' * 9_000
    project, task = upload(with_styles(bold_book(), "cellXfs", formats))
    assert task["status"] == "succeeded", task
    code, window = presentation(app, project, "S")
    assert code == 413, window
    assert window["error"]["message"] == "Trop de styles de présentation"


# F18: the Data tab copy is built by a background task, never in the request.


def data_book(rows=3000):
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    for row in range(1, rows + 1):
        sheet.cell(row, 1, f"item {row}")
        sheet.cell(row, 2, row * 1.5)
        sheet.cell(row, 3, row % 7)
    return book


def data_state(app, project, query=""):
    code, metadata = call(app, "GET", f"/api/projects/{project}/data", query=query)
    assert code == 200, metadata
    return metadata


def test_the_data_tab_answers_at_once_while_a_task_builds_its_copy(serve):
    app, upload = serve
    project, task = upload(data_book())
    assert task["status"] == "succeeded", task
    path = next(app.store.root.glob(f"*/projects/{project}"))
    first = data_state(app, project)
    assert first["state"] == "building" and not first["available"]
    assert not list(path.glob("source-data*.sqlite")), "built in the request"
    build = app.store.tasks[first["taskId"]]
    assert build["operation"] == "build_data" and build["projectId"] == project
    assert data_state(app, project)["taskId"] == first["taskId"], "one build"
    assert finish(app, first["taskId"])["status"] == "succeeded"
    ready = data_state(app, project)
    assert ready["available"] and ready["sheets"][0]["cellCount"] == 9000
    code, rows = call(
        app,
        "GET",
        f"/api/projects/{project}/datasette/source/sheet_001.json",
        query="_shape=objects&_size=2",
    )
    assert code == 200 and rows["rows"][0]["A"] == "item 1"


def _settled_storage(app, seconds=120):
    """Storage once background work (the library catalog) has finished."""
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if not any(app.store.root.rglob("*.building")):
            time.sleep(0.5)
            if not any(app.store.root.rglob("*.building")):
                return app.store._storage_bytes()
        time.sleep(0.2)
    pytest.fail("background work did not finish")


def test_a_failed_copy_says_why_and_can_be_prepared_again(tmp_path):
    data = workbook_bytes(data_book(20_000))
    upload = {"workbook": {"name": "d.xlsx", "data": base64.b64encode(data).decode()}}
    # What the workbook, its index and the library catalog take, measured on
    # a first server. The catalog is built after the import: with the smaller
    # index of schema 5 it no longer fits in a margin of 2 MiB.
    measure = create_app(tmp_path / "measure", resolve_user=lambda env: env.get("user"))
    try:
        code, imported = call(measure, "POST", "/api/projects", upload)
        assert code == 200, imported
        assert finish(measure, imported["task"]["id"])["status"] == "succeeded"
        used = _settled_storage(measure)
    finally:
        measure.close()
    # Room for all of that, not for a second copy of the cells.
    quota = int(used / 1048576) + 2
    app = create_app(
        tmp_path / "app",
        resolve_user=lambda env: env.get("user"),
        max_storage_mb=quota,
    )
    try:
        code, imported = call(app, "POST", "/api/projects", upload)
        assert code == 200, imported
        assert finish(app, imported["task"]["id"])["status"] == "succeeded"
        assert _settled_storage(app) < quota * 1048576
        project = imported["project"]["id"]
        first = data_state(app, project)
        assert finish(app, first["taskId"], seconds=300)["status"] == "failed"
        failed = data_state(app, project)
        assert failed["state"] == "failed" and failed["taskId"] == first["taskId"]
        assert "Quota de stockage" in failed["error"]
        code, refused = call(
            app, "GET", f"/api/projects/{project}/datasette/source/sheet_001.json"
        )
        assert code == 409 and "Quota de stockage" in refused["error"]["message"]
        again = data_state(app, project, query="retry=1")
        assert again["state"] == "building" and again["taskId"] != first["taskId"]
        finish(app, again["taskId"], seconds=300)
    finally:
        app.close()


def test_a_preparation_the_quota_cannot_start_says_so(tmp_path):
    """Asking for the view while storage is already full gives the quota
    message of the tab, not a bare 413 the tab cannot explain."""
    app = create_app(
        tmp_path / "app", resolve_user=lambda env: env.get("user"), max_storage_mb=2
    )
    try:
        data = workbook_bytes(data_book(50))
        code, imported = call(
            app,
            "POST",
            "/api/projects",
            {"workbook": {"name": "d.xlsx", "data": base64.b64encode(data).decode()}},
        )
        assert code == 200, imported
        assert finish(app, imported["task"]["id"])["status"] == "succeeded"
        project = imported["project"]["id"]
        # Another workbook's files fill the quota before the first opening.
        (app.store.root / "filler.bin").write_bytes(b"\0" * (2 * 1048576))
        for query in ("", "retry=1"):
            state = data_state(app, project, query=query)
            assert state["state"] == "failed", state
            assert "Quota de stockage du serveur atteint (2 Mio" in state["error"]
            assert "--max-storage-mb" in state["error"]
        assert not any(t["operation"] == "build_data" for t in app.store.tasks.values())
    finally:
        app.close()


def test_a_cancelled_preparation_is_not_reported_as_a_failure(serve):
    app, upload = serve
    project, task = upload(data_book(5000))
    assert task["status"] == "succeeded", task
    first = data_state(app, project)
    code, cancelled = call(app, "POST", f"/api/tasks/{first['taskId']}/cancel")
    assert code == 200, cancelled
    assert finish(app, first["taskId"])["status"] == "cancelled"
    state = data_state(app, project)
    assert state["state"] == "cancelled" and state["taskId"] == first["taskId"]
    assert "error" not in state
    code, refused = call(
        app, "GET", f"/api/projects/{project}/datasette/source/sheet_001.json"
    )
    assert code == 409 and "annulée" in refused["error"]["message"]
    again = data_state(app, project, query="retry=1")
    assert again["state"] == "building" and again["taskId"] != first["taskId"]
    assert finish(app, again["taskId"], seconds=300)["status"] == "succeeded"


def test_an_operation_the_quota_cannot_start_says_so(tmp_path):
    """Every operation refused because storage is full names the quota and
    the remedy, as the Data tab does, not the bare "storage full"."""
    app = create_app(
        tmp_path / "app", resolve_user=lambda env: env.get("user"), max_storage_mb=2
    )
    try:
        data = workbook_bytes(data_book(50))
        upload = {
            "workbook": {"name": "d.xlsx", "data": base64.b64encode(data).decode()}
        }
        code, imported = call(app, "POST", "/api/projects", upload)
        assert code == 200, imported
        assert finish(app, imported["task"]["id"])["status"] == "succeeded"
        project = imported["project"]["id"]
        (app.store.root / "filler.bin").write_bytes(b"\0" * (2 * 1048576))
        refused = [
            call(
                app,
                "POST",
                f"/api/projects/{project}/tasks",
                {"operation": "build_graph"},
            ),
            call(app, "POST", "/api/projects", upload),
        ]
        for code, answer in refused:
            assert code == 413, answer
            message = answer["error"]["message"]
            assert message.startswith("Quota de stockage du serveur atteint (2 Mio")
            assert "n’a pas pu démarrer" in message and "--max-storage-mb" in message
    finally:
        app.close()
