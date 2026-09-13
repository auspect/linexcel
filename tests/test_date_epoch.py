"""The workbook date system must apply before import, evaluation and steps."""

import datetime
import io
import re
import zipfile

import pytest
from openpyxl import Workbook
from openpyxl.utils.datetime import to_excel

from linexcel.analyzer import analyze_workbook
from linexcel.engine import _open_workbook, boot_engine
from linexcel.loader import _detect_epoch_1904, load_cached_values


def epoch_workbook(epoch_1904=False, *, boolean_text=None, broken_reference=False):
    workbook = Workbook()
    if epoch_1904:
        workbook.epoch = datetime.datetime(1904, 1, 1)
    sheet = workbook.active
    sheet.title = "Dates"
    sheet["A1"] = datetime.datetime(2026, 2, 1)
    sheet["B1"] = "=YEAR(A1)"
    sheet["C1"] = "=YEAR(D1)"
    sheet["D1"] = to_excel(datetime.datetime(2026, 2, 1), workbook.epoch)
    sheet["E1"] = "=YEAR(A1+1)+YEAR(DATE(2026,2,1))"
    sheet["F1"] = "=D1+1"
    sheet["F1"].number_format = "yyyy-mm-dd"
    sheet["G1"] = "=YEAR(DATE(2026,2,1))"
    if broken_reference:
        sheet["Z1"] = "=NoSheet!A1"
    stream = io.BytesIO()
    workbook.save(stream)
    rewritten = io.BytesIO()
    with zipfile.ZipFile(stream) as source, zipfile.ZipFile(rewritten, "w") as dest:
        for part in source.infolist():
            data = source.read(part.filename)
            if part.filename == "xl/worksheets/sheet1.xml":
                stored_serial = sheet["D1"].value + 1
                data = re.sub(
                    rb"<f>D1\+1</f><v(?:\s*/>|></v>)",
                    f"<f>D1+1</f><v>{stored_serial}</v>".encode(),
                    data,
                )
            if part.filename == "xl/workbook.xml" and boolean_text is not None:
                data = re.sub(rb' date1904="[^"]*"', b"", data)
                data = data.replace(
                    b"<workbookPr", f'<workbookPr date1904="{boolean_text}"'.encode(), 1
                )
            dest.writestr(part, data)
    return rewritten.getvalue()


@pytest.mark.parametrize("epoch_1904", [False, True])
def test_date_system_is_applied_before_the_engine_imports_cells(epoch_1904):
    engine = _open_workbook(epoch_workbook(epoch_1904), parallel=False)
    assert engine.get_value("Dates", 1, 1) == datetime.date(2026, 2, 1)
    serial = 44592 if epoch_1904 else 46054
    assert engine.get_value("Dates", 1, 4) == serial
    engine.evaluate_all()
    assert engine.get_value("Dates", 1, 2) == 2026
    assert engine.get_value("Dates", 1, 3) == 2026
    assert engine.get_value("Dates", 1, 7) == 2026


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("targeted", [False, True])
def test_date_formulas_and_scratch_steps_use_the_same_epoch(epoch_1904, targeted):
    targets = (
        ["Dates!B1", "Dates!C1", "Dates!E1", "Dates!F1", "Dates!G1"]
        if targeted
        else None
    )
    graph = analyze_workbook(epoch_workbook(epoch_1904), targets=targets)["graph"]
    nodes = {node.get("addr"): node for node in graph["nodes"] if node.get("formula")}
    assert nodes["B1"]["value"] == 2026
    assert nodes["C1"]["value"] == 2026
    assert nodes["G1"]["value"] == 2026
    assert nodes["F1"]["valueDate"] == "2026-02-02"
    assert nodes["E1"]["value"] == 4052
    steps = nodes["E1"]["steps"]
    assert steps["value"] == 4052
    assert [step["value"] for step in steps["children"]] == [2026, 2026]
    assert steps["children"][0]["children"][0]["value"] == "2026-02-02"
    assert steps["children"][1]["children"][0]["value"] == "2026-02-01"
    assert not any("differs from file" in w for w in graph["meta"]["warnings"])


@pytest.mark.parametrize(
    ("text", "expected"), [("1", True), ("true", True), ("0", False), ("false", False)]
)
def test_xml_boolean_spellings_are_shared_by_loader_and_engine(text, expected):
    data = epoch_workbook(expected, boolean_text=text)
    assert _detect_epoch_1904(data) is expected
    assert load_cached_values(data).epoch_1904 is expected
    engine = _open_workbook(data, parallel=False)
    assert engine.get_value("Dates", 1, 1) == datetime.date(2026, 2, 1)


def test_quarantine_retry_preserves_the_workbook_epoch():
    warnings = []
    session = boot_engine(epoch_workbook(True, broken_reference=True), warnings)
    assert session.quarantined
    assert session.engine.get_value("Dates", 1, 2) == 2026
    assert session.engine.get_value("Dates", 1, 3) == 2026
