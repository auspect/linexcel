"""Temporal caches retain their calendar meaning across both source readers."""

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
from linexcel.values import readings_agree


@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize(
    ("serial", "cache_serial", "expected"),
    [
        (60 - 0.25 / 86_400_000, 61, "differ"),
        (61, 60 - 0.25 / 86_400_000, "differ"),
        (60 - 0.25 / 86_400_000, 60 - 0.25 / 86_400_000, "same"),
        (60 - 0.25 / 86_400_000, 60, "same"),
        (60, 60 - 0.25 / 86_400_000, "same"),
        (60 - 1 / 86_400_000, 60, "differ"),
        (61 - 0.25 / 86_400_000, 61, "same"),
        (61 - 0.49 / 86_400_000, 61, "same"),
        (61 - 0.51 / 86_400_000, 61, "differ"),
        (61 - 1 / 86_400_000, 61, "differ"),
        (61, 61 - 0.25 / 86_400_000, "same"),
        (61 - 0.25 / 86_400_000, 59, "differ"),
    ],
)
def test_leap_day_boundary_preserves_serial_distance(
    slow_reader, epoch_1904, serial, cache_serial, expected, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = workbook_bytes(
        epoch_1904, serial, "yyyy-mm-dd hh:mm:ss.000", cache_serial=cache_serial
    )
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    assert node["value"] == float(format(serial, ".16g"))
    assert node["cachedAgreement"] == expected
    assert any("Dates!B1" in w for w in graph["meta"]["warnings"]) == (
        expected == "differ"
    )
    if not epoch_1904 and 60 - 0.5 / 86_400_000 < cache_serial <= 60:
        assert node["cachedValue"] == cache_serial


@pytest.mark.parametrize(
    ("cached", "epoch_1904", "expected"),
    [
        (dt.datetime(1900, 3, 1), False, "same"),
        (dt.date(1900, 3, 1), False, "same"),
        (dt.datetime(1900, 3, 1, tzinfo=dt.UTC), False, "differ"),
        (dt.datetime(1900, 2, 28), False, "differ"),
        (dt.datetime(1900, 3, 2), False, "differ"),
        ("1900-03-01", False, "differ"),
        (dt.datetime(1900, 3, 1), None, "differ"),
        (dt.datetime(1900, 3, 1), True, "differ"),
    ],
)
def test_phantom_end_comparison_requires_typed_cache_and_calendar(
    cached, epoch_1904, expected
):
    assert (
        readings_agree(61 - 0.25 / 86_400_000, cached, None, epoch_1904=epoch_1904)
        == expected
    )


@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("offset_ms", [-4, -1, -0.25, 0, 0.25, 1, 4])
@pytest.mark.parametrize("number_format", ["yyyy-mm-dd hh:mm:ss.000", "General"])
def test_phantom_date_cache_uses_temporal_precision(
    slow_reader, offset_ms, number_format, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    serial = 60.5 + offset_ms / 86_400_000
    data = workbook_bytes(False, serial, number_format, cache_serial=60.5)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    differs = number_format != "General" and abs(offset_ms) > 0.5
    assert node["value"] == serial
    assert node["cachedValue"] == 60.5
    assert node["valueSource"] == "engine"
    assert "valueDate" not in node  # Do not invent a representable calendar day.
    assert node["cachedAgreement"] == ("differ" if differs else "same")
    assert (
        any("Dates!B1" in warning for warning in graph["meta"]["warnings"]) is differs
    )


@pytest.mark.parametrize(
    ("raw", "cached", "expected"),
    [
        (60.5 + 0.25 / 86_400_000, 60.5, "same"),
        (60.5 + 4 / 86_400_000, 60.5, "differ"),
        (True, True, "same"),
        (True, 1, "differ"),
        (1, True, "differ"),
        ("60.5", 60.5, "differ"),
        (60.5, "60.5", "differ"),
        (float("inf"), 60.5, "differ"),
        (float("nan"), 60.5, "differ"),
    ],
)
def test_numeric_date_comparison_preserves_types(raw, cached, expected):
    assert readings_agree(raw, cached, None, date_formatted=True) == expected


def workbook_bytes(epoch_1904, serial, number_format, *, cache_serial=None):
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
                stored = serial if cache_serial is None else cache_serial
                content, count = re.subn(
                    rb"<f>A1</f><v(?:\s*/>|></v>)",
                    f"<f>A1</f><v>{stored}</v>".encode(),
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
@pytest.mark.parametrize("number_format", ["yyyy-mm-dd hh:mm:ss.000", "hh:mm:ss.000"])
@pytest.mark.parametrize("remaining_ms", [0.1, 0.25, 1.0])
def test_first_midnight_cache_preserves_rounded_day(
    epoch_1904, number_format, remaining_ms
):
    serial = 1 - remaining_ms / 86_400_000
    data = workbook_bytes(epoch_1904, serial, number_format)
    if remaining_ms < 0.5:
        expected = dt.datetime(1904, 1, 2) if epoch_1904 else dt.datetime(1900, 1, 1)
    else:
        expected = dt.time(23, 59, 59, 999000)
    for reader in (_load_cached_values_calamine, _load_cached_values_openpyxl):
        cached = reader(data)
        for column in (1, 2):
            assert cached.get("Dates", 1, column) == expected
            assert cached.is_date("Dates", 1, column)
        duration = cached.get("Dates", 1, 3)
        assert isinstance(duration, dt.timedelta)
        assert abs(
            duration - dt.timedelta(days=1, milliseconds=-remaining_ms)
        ) <= dt.timedelta(microseconds=500)
        assert cached.get("Dates", 1, 4) == serial
        assert cached.get("Dates", 1, 5) == dt.time(18)
        assert cached.get("Dates", 1, 6) == "18:00:00"


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
def test_first_midnight_cache_preserves_retention_limit(epoch_1904, slow_reader):
    data = workbook_bytes(epoch_1904, 1 - 0.25 / 86_400_000, "hh:mm:ss.000")
    cached = load_cached_values(
        data, max_cells_per_sheet=1, max_dense_cells=0 if slow_reader else 100
    )
    expected = dt.datetime(1904, 1, 2) if epoch_1904 else dt.datetime(1900, 1, 1)
    assert cached.get("Dates", 1, 1) == expected
    assert len(cached) == 1
    assert cached.get("Dates", 1, 2) is None
    assert cached.truncated_sheets == {"Dates"}


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
def test_analysis_preserves_first_midnight_cache(epoch_1904, slow_reader, monkeypatch):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    serial = 1 - 0.25 / 86_400_000
    data = workbook_bytes(epoch_1904, serial, "yyyy-mm-dd hh:mm:ss.000")
    result = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))
    node = next(n for n in result["graph"]["nodes"] if n.get("addr") == "B1")
    assert node["cachedValue"] == ("1904-01-02" if epoch_1904 else "1900-01-01")
    assert node["value"] == serial
    assert node["valueSource"] == "engine"
    assert node["cachedAgreement"] == "same"
    assert node["valueDate"] == ("1904-01-01" if epoch_1904 else "1899-12-31")
    assert not any(
        "Dates!B1" in warning for warning in result["graph"]["meta"]["warnings"]
    )


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("remaining_ms", [0.25, 0.49, 0.51])
def test_first_midnight_display_anchor_preserves_rounding_boundary(
    epoch_1904, slow_reader, remaining_ms, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    serial = 1 - remaining_ms / 86_400_000
    data = workbook_bytes(epoch_1904, serial, "yyyy-mm-dd hh:mm:ss.000")
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    expected = (
        ("1904-01-01" if epoch_1904 else "1899-12-31")
        if remaining_ms < 0.5
        else "23:59:59.999000"
    )
    assert node["valueDate"] == expected
    assert node["value"] == serial
    assert node["cachedAgreement"] == "same"


@pytest.mark.parametrize("number_format", ["General", "[h]:mm:ss.000"])
@pytest.mark.parametrize("epoch_1904", [False, True])
def test_first_midnight_does_not_give_numbers_or_durations_a_date(
    number_format, epoch_1904
):
    serial = 1 - 0.25 / 86_400_000
    data = workbook_bytes(epoch_1904, serial, number_format)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    assert "valueDate" not in node


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("remaining_ms", [0.25, 0.49])
@pytest.mark.parametrize("number_format", ["yyyy-mm-dd hh:mm:ss.000", "hh:mm:ss.000"])
def test_first_midnight_display_with_time_only_cache(
    epoch_1904, slow_reader, remaining_ms, number_format, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    serial = 1 - remaining_ms / 86_400_000
    cache_serial = 1 - 1 / 86_400_000
    data = workbook_bytes(epoch_1904, serial, number_format, cache_serial=cache_serial)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    assert node["cachedValue"] == "23:59:59.999000"
    assert node["value"] == serial
    assert node["valueSource"] == "engine"
    assert node["cachedAgreement"] == "differ"
    assert any("Dates!B1" in warning for warning in graph["meta"]["warnings"])
    assert node["valueDate"] == ("1904-01-01" if epoch_1904 else "1899-12-31")


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("serial", [1, 1.5, -2.25, 0, 0.5])
@pytest.mark.parametrize("number_format", ["yyyy-mm-dd hh:mm:ss", "hh:mm:ss"])
def test_date_results_with_time_only_cache(
    epoch_1904, slow_reader, serial, number_format, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = workbook_bytes(epoch_1904, serial, number_format, cache_serial=0.5)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    expected = {
        1: "1904-01-02" if epoch_1904 else "1900-01-01",
        1.5: "1904-01-02" if epoch_1904 else "1900-01-01",
        -2.25: "1903-12-29" if epoch_1904 else "1899-12-27",
        0: "00:00:00",
        0.5: "12:00:00",
    }[serial]
    assert node["value"] == serial
    assert node["valueSource"] == "engine"
    assert node["cachedValue"] == "12:00:00"
    assert node["cachedAgreement"] == ("same" if serial == 0.5 else "differ")
    assert any("Dates!B1" in warning for warning in graph["meta"]["warnings"]) is (
        serial != 0.5
    )
    assert node["valueDate"] == expected


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("serial", [0, 0.5, 0.75])
@pytest.mark.parametrize("number_format", ["yyyy-mm-dd hh:mm:ss", "hh:mm:ss"])
def test_time_results_with_full_date_cache(
    epoch_1904, slow_reader, serial, number_format, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = workbook_bytes(epoch_1904, serial, number_format, cache_serial=1.5)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    assert node["value"] == serial
    assert node["valueSource"] == "engine"
    assert node["cachedValue"] == (
        "1904-01-02 12:00:00" if epoch_1904 else "1900-01-01 12:00:00"
    )
    assert node["cachedAgreement"] == "differ"
    assert any("Dates!B1" in warning for warning in graph["meta"]["warnings"])
    assert node["valueDate"] == (
        "1904-01-01"
        if epoch_1904
        else {0: "00:00:00", 0.5: "12:00:00", 0.75: "18:00:00"}[serial]
    )


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("serial", [0, 0.5])
@pytest.mark.parametrize("number_format", ["General", "[h]:mm:ss"])
def test_fractional_results_do_not_invent_time_format(
    epoch_1904, serial, number_format
):
    data = workbook_bytes(epoch_1904, serial, number_format, cache_serial=1.5)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    assert "valueDate" not in node
    assert node["cachedAgreement"] == "differ"


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
    assert node["valueDate"] == expected[:10]
    assert node["cachedAgreement"] == "same"


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("serial", [-40000.25, -2.25, -1, -0.25])
@pytest.mark.parametrize("offset_seconds", [-0.00025, 0.00025, 1, 86400])
def test_negative_date_comparison_preserves_rounding_and_real_gaps(
    epoch_1904, slow_reader, serial, offset_seconds, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = workbook_bytes(
        epoch_1904,
        serial + offset_seconds / 86400,
        "yyyy-mm-dd hh:mm:ss.000",
        cache_serial=serial,
    )
    result = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))
    node = next(n for n in result["graph"]["nodes"] if n.get("addr") == "B1")
    assert node["valueSource"] == "engine"
    assert node["cachedAgreement"] == ("same" if offset_seconds < 0.0005 else "differ")
    # Warnings and the serialized node must tell the same story.
    mismatch = any(
        "Dates!B1" in warning for warning in result["graph"]["meta"]["warnings"]
    )
    assert mismatch is (offset_seconds >= 0.0005)


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("serial", [-2.25, 0.75, 60.5, 45000.5])
@pytest.mark.parametrize(
    "number_format",
    ['yyyy-mm-dd "[h]"', 'yyyy-mm-dd "[mm]"', 'yyyy-mm-dd "[ss]"', r"yyyy-mm-dd \[h\]"],
)
def test_literal_elapsed_labels_preserve_cache_types(epoch_1904, serial, number_format):
    data = workbook_bytes(epoch_1904, serial, number_format)
    # A literal label must behave exactly like the same calendar format without it.
    plain = _load_cached_values_openpyxl(
        workbook_bytes(epoch_1904, serial, "yyyy-mm-dd")
    )
    for reader in (_load_cached_values_calamine, _load_cached_values_openpyxl):
        cached = reader(data)
        for column in (1, 2):
            expected = plain.get("Dates", 1, column)
            actual = cached.get("Dates", 1, column)
            assert actual == expected
            assert type(actual) is type(expected)
            assert cached.is_date("Dates", 1, column)
        assert cached.get("Dates", 1, 3) == dt.timedelta(days=serial)
        assert not cached.is_date("Dates", 1, 3)
        assert cached.get("Dates", 1, 4) == serial
        assert cached.get("Dates", 1, 5) == dt.time(18)
        assert cached.get("Dates", 1, 6) == "18:00:00"
        assert cached.get("Dates", 1, 7) is None


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
def test_literal_elapsed_label_preserves_retention_limit(epoch_1904, slow_reader):
    data = workbook_bytes(epoch_1904, 45000.5, 'yyyy-mm-dd "[ss]"')
    cached = load_cached_values(
        data, max_cells_per_sheet=1, max_dense_cells=0 if slow_reader else 100
    )
    assert len(cached) == 1
    expected = (
        dt.datetime(2027, 3, 16, 12) if epoch_1904 else dt.datetime(2023, 3, 15, 12)
    )
    assert cached.get("Dates", 1, 1) == expected
    assert cached.get("Dates", 1, 2) is None
    assert cached.truncated_sheets == {"Dates"}


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
def test_analysis_preserves_dates_with_literal_elapsed_labels(
    epoch_1904, slow_reader, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = workbook_bytes(epoch_1904, 45000.5, 'yyyy-mm-dd "[ss]"')
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    expected = "2027-03-16 12:00:00" if epoch_1904 else "2023-03-15 12:00:00"
    assert node["cachedValue"] == expected
    assert node["value"] == 45000.5
    assert node["valueSource"] == "engine"
    assert node["cachedAgreement"] == "same"


@pytest.mark.parametrize("serial", [-2.25, 60.5, 45000.5])
@pytest.mark.parametrize("number_format", ['[h]:mm:ss "[ss]"', '[mm]:ss "[h]"'])
def test_real_durations_with_literal_labels_remain_durations(serial, number_format):
    data = workbook_bytes(False, serial, number_format)
    for reader in (_load_cached_values_calamine, _load_cached_values_openpyxl):
        cached = reader(data)
        for column in (1, 2):
            assert cached.get("Dates", 1, column) == dt.timedelta(days=serial)
            assert not cached.is_date("Dates", 1, column)


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("offset_ms", [-1.0, -0.25, 0.25, 1.0])
@pytest.mark.parametrize("number_format", ["yyyy-mm-dd hh:mm:ss.000", "hh:mm:ss.000"])
def test_first_midnight_comparison_preserves_real_gaps(
    epoch_1904, slow_reader, offset_ms, number_format, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = workbook_bytes(
        epoch_1904, 1 + offset_ms / 86_400_000, number_format, cache_serial=1
    )
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(n for n in graph["nodes"] if n.get("addr") == "B1")
    differs = abs(offset_ms) > 0.5
    assert node["cachedAgreement"] == ("differ" if differs else "same")
    assert any("Dates!B1" in w for w in graph["meta"]["warnings"]) is differs
