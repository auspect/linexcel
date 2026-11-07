"""Negative date serials retain whole days when read from the source file."""

import datetime as dt
import io
import re
import zipfile

import pytest
from openpyxl import Workbook

from linexcel import ExecutionPolicy
from linexcel.analyzer import analyze_workbook
from linexcel.loader import (
    _load_cached_values_calamine,
    _load_cached_values_openpyxl,
    load_cached_values,
)


def workbook_bytes(epoch_1904, serial, number_format):
    workbook = Workbook()
    if epoch_1904:
        workbook.epoch = dt.datetime(1904, 1, 1)
    sheet = workbook.active
    sheet.title = "Dates"
    sheet["A1"] = serial
    sheet["B1"] = "=A1"
    sheet["C1"] = serial
    sheet["C1"].number_format = "[h]:mm:ss"
    sheet["D1"] = serial
    sheet["E1"] = 0.75
    sheet["E1"].number_format = "hh:mm:ss"
    sheet["F1"] = "18:00:00"
    sheet["G1"] = "=A1+1"
    for address in ("A1", "B1"):
        sheet[address].number_format = number_format
    source, output = io.BytesIO(), io.BytesIO()
    workbook.save(source)
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(output, "w") as result:
        for entry in original.infolist():
            content = original.read(entry)
            if entry.filename == "xl/worksheets/sheet1.xml":
                content, count = re.subn(
                    rb"<f>A1</f><v(?:\s*/>|></v>)",
                    f"<f>A1</f><v>{serial}</v>".encode(),
                    content,
                )
                assert count == 1
            result.writestr(entry, content)
    return output.getvalue()


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("serial", [-2.25, -1, -0.25])
@pytest.mark.parametrize("number_format", ["yyyy-mm-dd", "hh:mm:ss"])
def test_negative_dates_agree_across_cache_readers(epoch_1904, serial, number_format):
    data = workbook_bytes(epoch_1904, serial, number_format)
    epoch = dt.datetime(1904, 1, 1) if epoch_1904 else dt.datetime(1899, 12, 30)
    expected = epoch + dt.timedelta(days=serial)
    for reader in (_load_cached_values_calamine, _load_cached_values_openpyxl):
        cached = reader(data)
        for column in (1, 2):
            assert cached.get("Dates", 1, column) == expected
            assert cached.is_date("Dates", 1, column)
        assert cached.get("Dates", 1, 3) == dt.timedelta(days=serial)
        assert cached.get("Dates", 1, 4) == serial
        assert cached.get("Dates", 1, 5) == dt.time(18)
        assert cached.get("Dates", 1, 6) == "18:00:00"
        assert cached.get("Dates", 1, 7) is None


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
def test_negative_date_cache_preserves_retention_limit(epoch_1904, slow_reader):
    data = workbook_bytes(epoch_1904, -2.25, "yyyy-mm-dd")
    cached = load_cached_values(
        data, max_cells_per_sheet=1, max_dense_cells=0 if slow_reader else 100
    )
    assert len(cached) == 1
    epoch = dt.datetime(1904, 1, 1) if epoch_1904 else dt.datetime(1899, 12, 30)
    assert cached.get("Dates", 1, 1) == epoch - dt.timedelta(days=2.25)
    assert cached.get("Dates", 1, 2) is None
    assert cached.truncated_sheets == {"Dates"}


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
def test_analysis_reports_complete_negative_date_cache(
    epoch_1904, slow_reader, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = workbook_bytes(epoch_1904, -2.25, "yyyy-mm-dd hh:mm:ss")
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    expected = "1903-12-29 18:00:00" if epoch_1904 else "1899-12-27 18:00:00"
    assert node["cachedValue"] == expected
    assert node["value"] == -2.25
    assert node["valueSource"] == "engine"
