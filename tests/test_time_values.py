"""Synthetic Excel time caches retain their meaning across JSON serialization."""

import datetime
import io
import math
import re
import zipfile
from functools import partial

import pytest
from openpyxl import Workbook

from linexcel import ExecutionPolicy
from linexcel.analyzer import analyze_workbook
from linexcel.loader import load_cached_values
from linexcel.values import readings_agree, serial_to_date_text, serial_to_time_text


@pytest.mark.parametrize("serialized", [False, True])
@pytest.mark.parametrize("offset", [-1000, -501, -500, -400, 0, 400, 500, 501, 1000])
def test_typed_time_readings_tolerate_only_cache_rounding(serialized, offset):
    raw = datetime.time(12)
    stored = (
        datetime.datetime(2026, 8, 7, 12) + datetime.timedelta(microseconds=offset)
    ).time()
    if serialized:
        raw, stored = str(raw), str(stored)
    expected = "same" if abs(offset) <= 500 else "differ"
    assert readings_agree(raw, stored, "12:00:00") == expected
    assert readings_agree(stored, raw, "12:00:00") == expected


@pytest.mark.parametrize("serialized", [False, True])
def test_typed_time_rounding_does_not_wrap_around_midnight(serialized):
    raw, stored = datetime.time(23, 59, 59, 999600), datetime.time.min
    if serialized:
        raw, stored = str(raw), str(stored)
    assert readings_agree(raw, stored, "23:59:59.999600") == "differ"


@pytest.mark.parametrize("serialized", [False, True])
def test_typed_time_rounding_does_not_assume_a_timezone(serialized):
    raw = datetime.time(12, tzinfo=datetime.UTC)
    stored = datetime.time(12)
    if serialized:
        raw, stored = str(raw), str(stored)
    assert readings_agree(raw, stored, "12:00:00") == "differ"
    assert readings_agree(stored, raw, "12:00:00") == "differ"


@pytest.mark.parametrize("serialized", [False, True])
@pytest.mark.parametrize(
    ("offset", "verdict"),
    [(0, "same"), (0.0004, "same"), (0.001, "differ"), (3600, "differ")],
)
def test_typed_datetime_readings_keep_hours_and_cache_precision(
    serialized, offset, verdict
):
    raw = datetime.datetime(2026, 8, 7, 12)
    stored = raw + datetime.timedelta(seconds=offset)
    if serialized:
        raw, stored = str(raw), str(stored)
    assert readings_agree(raw, stored, "2026-08-07") == verdict
    assert readings_agree(stored, raw, "2026-08-07") == verdict


@pytest.mark.parametrize("serialized", [False, True])
@pytest.mark.parametrize("hour", [0, 12])
def test_typed_date_is_midnight_not_any_time_on_the_same_day(serialized, hour):
    raw = datetime.datetime(2026, 8, 7, hour)
    stored = datetime.date(2026, 8, 7)
    if serialized:
        raw, stored = str(raw), str(stored)
    verdict = "same" if hour == 0 else "differ"
    assert readings_agree(raw, stored, "2026-08-07") == verdict
    assert readings_agree(stored, raw, "2026-08-07") == verdict


@pytest.mark.parametrize("serialized", [False, True])
def test_typed_datetime_rounding_across_midnight(serialized):
    raw = datetime.datetime(2026, 8, 7, 23, 59, 59, 999600)
    stored = datetime.datetime(2026, 8, 8)
    if serialized:
        raw, stored = str(raw), str(stored)
    assert readings_agree(raw, stored, "2026-08-07") == "same"


@pytest.mark.parametrize("serialized", [False, True])
def test_typed_datetime_does_not_invent_a_timezone(serialized):
    raw = datetime.datetime(2026, 8, 7, 12, tzinfo=datetime.UTC)
    stored = datetime.datetime(2026, 8, 7, 12)
    if serialized:
        raw, stored = str(raw), str(stored)
    assert readings_agree(raw, stored, "2026-08-07") == "differ"


def test_timestamp_text_without_metadata_is_not_coerced():
    assert (
        readings_agree("2026-08-07T12:00:00", "2026-08-07 12:00:00", None) == "differ"
    )


@pytest.mark.parametrize(
    ("serial", "stored"),
    [
        (0, datetime.time(0)),
        (0.5, datetime.time(12)),
        (0.25, datetime.time(6)),
        (0.5 + 0.123 / 86400, datetime.time(12, 0, 0, 123000)),
    ],
)
def test_raw_and_serialized_time_caches_agree(serial, stored):
    assert readings_agree(serial, stored, None) == "same"
    assert (
        readings_agree(serial, stored.isoformat(), serial_to_time_text(serial))
        == "same"
    )


@pytest.mark.parametrize("serial", [-1, 1, 1.5, float("nan"), float("inf"), True])
def test_a_time_of_day_does_not_discard_days_or_coerce_booleans(serial):
    assert serial_to_time_text(serial) is None
    assert readings_agree(serial, datetime.time(12), None) == "differ"


@pytest.mark.parametrize("stored", ["0", "00:00:00", "12:00:00"])
def test_time_shaped_text_without_temporal_metadata_stays_text(stored):
    assert readings_agree(0, stored, None) == "differ"


def test_a_true_time_difference_is_not_hidden():
    assert readings_agree(0.5, datetime.time(12, 1), None) == "differ"
    assert readings_agree(0.5, "12:01:00", "12:00:00") == "differ"


def test_no_timezone_is_assumed_for_excel_serials():
    assert readings_agree(0.5, datetime.time(12, tzinfo=datetime.UTC), None) == "differ"
    assert readings_agree(0.5, "12:00:00+00:00", "12:00:00") == "differ"


def _time_workbook(serial: float, cached_serial: float, epoch_1904: bool) -> bytes:
    wb = Workbook()
    if epoch_1904:
        wb.epoch = datetime.datetime(1904, 1, 1)
    ws = wb.active
    ws.title = "Clock"
    ws["A1"] = serial
    ws["B1"] = "=A1"
    ws["B1"].number_format = "hh:mm:ss.000"
    original = io.BytesIO()
    wb.save(original)
    output = io.BytesIO()
    with zipfile.ZipFile(original) as source, zipfile.ZipFile(output, "w") as dest:
        for part in source.infolist():
            payload = source.read(part.filename)
            if part.filename == "xl/worksheets/sheet1.xml":
                payload, count = re.subn(
                    rb"</f><v(?:\s*/>|></v>)",
                    f"</f><v>{cached_serial}</v>".encode(),
                    payload,
                )
                assert count == 1
            dest.writestr(part, payload)
    return output.getvalue()


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("serial", [0, 0.5, 0.5 + 0.123 / 86400])
@pytest.mark.parametrize("slow_reader", [False, True])
def test_real_in_memory_formula_time_cache_is_not_a_divergence(
    serial, epoch_1904, slow_reader, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = _time_workbook(serial, serial, epoch_1904)
    assert isinstance(load_cached_values(data).get("Clock", 1, 2), datetime.time)
    graph = partial(analyze_workbook, execution=ExecutionPolicy(isolated=False))(data)[
        "graph"
    ]
    node = next(node for node in graph["nodes"] if node["id"] == "c:Clock!B1")
    assert node["value"] == pytest.approx(serial)
    assert node["valueSource"] == "engine"
    assert node["cachedAgreement"] == "same"
    assert node["valueDate"] == serial_to_time_text(serial)
    assert not any("differs from file" in w for w in graph["meta"]["warnings"])


def test_real_in_memory_time_cache_detects_a_wrong_minute():
    graph = analyze_workbook(_time_workbook(0.5, 0.5 + 1 / 1440, False))["graph"]
    node = next(node for node in graph["nodes"] if node["id"] == "c:Clock!B1")
    assert node["cachedAgreement"] == "differ"
    assert any("differs from file" in w for w in graph["meta"]["warnings"])


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize(
    ("computed_fraction", "cached_fraction", "agreement"),
    [
        (0.75, 0.5, "differ"),
        (0.5, 0, "differ"),
        (0.5, 0.5, "same"),
        (0.5 + 0.123 / 86400, 0.5 + 0.123 / 86400, "same"),
        (0.5 + 0.0004 / 86400, 0.5, "same"),
        (0.5 + 0.001 / 86400, 0.5, "differ"),
    ],
)
def test_date_time_caches_compare_the_hour_as_well_as_the_day(
    epoch_1904, slow_reader, computed_fraction, cached_fraction, agreement, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    # The existing fixture's clock format also produces a datetime cache when
    # the serial contains whole days. Compare both reader implementations.
    day = 44592 if epoch_1904 else 46054
    data = _time_workbook(day + computed_fraction, day + cached_fraction, epoch_1904)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(node for node in graph["nodes"] if node["id"] == "c:Clock!B1")
    # OOXML's decimal serialization can shift the last binary float digit.
    assert node["value"] == pytest.approx(day + computed_fraction, rel=0, abs=1e-10)
    assert node["valueSource"] == "engine"
    assert node["valueDate"] == "2026-02-01"
    assert node["cachedAgreement"] == agreement
    assert any("differs from file" in w for w in graph["meta"]["warnings"]) == (
        agreement == "differ"
    )
    if agreement == "differ" and cached_fraction == 0.5:
        assert any(
            "differs from file value 2026-02-01 12:00:00" in warning
            for warning in graph["meta"]["warnings"]
        )


@pytest.mark.parametrize(
    "stored",
    [
        datetime.datetime(2026, 2, 1, 12, tzinfo=datetime.UTC),
        "2026-02-01 12:00:00+00:00",
    ],
)
def test_date_time_caches_do_not_assume_a_timezone(stored):
    assert readings_agree(46054.5, stored, "2026-02-01") == "differ"


def test_date_shaped_text_needs_temporal_metadata():
    assert readings_agree(46054.5, "2026-02-01 12:00:00", None) == "differ"


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("stored_type", ["datetime", "date", "text"])
@pytest.mark.parametrize(
    ("offset_seconds", "expected"),
    [
        (-86400, "differ"),
        (-0.001, "differ"),
        (-0.0004, "same"),
        (0.0004, "same"),
        (0.001, "differ"),
        (86400, "differ"),
    ],
)
def test_datetime_tolerance_spans_midnight(
    epoch_1904, stored_type, offset_seconds, expected
):
    midnight = 44593 if epoch_1904 else 46055
    raw = midnight + offset_seconds / 86400
    stored = datetime.datetime(2026, 2, 2)
    if stored_type == "date":
        stored = stored.date()
    elif stored_type == "text":
        stored = stored.isoformat(sep=" ")
    assert readings_agree(raw, stored, serial_to_date_text(raw, epoch_1904)) == expected


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize(
    ("offset_seconds", "expected"),
    [(-0.001, "differ"), (-0.0004, "same"), (0.0004, "same"), (0.001, "differ")],
)
def test_formula_cache_rounding_across_midnight(
    epoch_1904, slow_reader, offset_seconds, expected, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    midnight = 44593 if epoch_1904 else 46055
    data = _time_workbook(midnight + offset_seconds / 86400, midnight, epoch_1904)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(node for node in graph["nodes"] if node["id"] == "c:Clock!B1")
    assert node["valueSource"] == "engine"
    assert node["cachedAgreement"] == expected
    assert any("differs from file" in w for w in graph["meta"]["warnings"]) == (
        expected == "differ"
    )


@pytest.mark.parametrize("epoch_1904", [False, True])
def test_date_label_does_not_round_a_fraction_up_to_the_next_day(epoch_1904):
    raw = math.nextafter(2.0, 0.0)
    expected = "1904-01-02" if epoch_1904 else "1900-01-01"
    assert serial_to_date_text(raw, epoch_1904) == expected


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize(
    ("text", "fraction", "whole_date"),
    [
        ("00:00:00", 0, False),
        ("12:00:00", 0.5, False),
        ("2026-02-01", 0, True),
        ("2026-02-01 12:00:00", 0.5, True),
    ],
)
def test_formula_text_does_not_agree_with_identically_displayed_temporal_cache(
    epoch_1904, slow_reader, text, fraction, whole_date, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    day = (44592 if epoch_1904 else 46054) if whole_date else 0
    data = _time_workbook(text, day + fraction, epoch_1904)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(node for node in graph["nodes"] if node["id"] == "c:Clock!B1")
    assert node["value"] == node["cachedValue"] == text
    assert node["valueSource"] == "engine"
    assert node["valueKind"] == "text"
    assert "valueDate" not in node
    assert node["cachedAgreement"] == "differ"
    assert any(
        "Clock!B1:" in warning and "differs from file value" in warning
        for warning in graph["meta"]["warnings"]
    )


@pytest.mark.parametrize("stored_type", ["date", "datetime", "text", "time"])
@pytest.mark.parametrize("epoch_1904", [None, False, True])
def test_first_midnight_requires_calendar_and_typed_date(stored_type, epoch_1904):
    stored = {
        "date": datetime.date(1900, 1, 1),
        "datetime": datetime.datetime(1900, 1, 1),
        "text": "1900-01-01",
        "time": datetime.time(0),
    }[stored_type]
    expected = (
        "same"
        if epoch_1904 is False and stored_type in ("date", "datetime")
        else "differ"
    )
    assert (
        readings_agree(1 - 0.25 / 86_400_000, stored, None, epoch_1904=epoch_1904)
        == expected
    )


@pytest.mark.parametrize(
    "stored",
    [
        datetime.datetime(1900, 1, 1, tzinfo=datetime.UTC),
        datetime.datetime(1899, 12, 31),
        datetime.datetime(1900, 1, 1, microsecond=1000),
        datetime.datetime(1900, 1, 2),
        datetime.datetime(1904, 1, 2),
    ],
)
def test_first_midnight_does_not_hide_calendar_or_time_gaps(stored):
    assert (
        readings_agree(1 - 0.25 / 86_400_000, stored, None, epoch_1904=False)
        == "differ"
    )
