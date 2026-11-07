"""Elapsed Excel durations must not acquire calendar dates in reports."""

import datetime
import io
import re
import zipfile

import pytest
from openpyxl import Workbook

from linexcel import ExecutionPolicy
from linexcel.analyzer import analyze_workbook
from linexcel.loader import _is_date_format, load_cached_values
from linexcel.values import readings_agree


def duration_workbook(
    number_format, serial, epoch_1904, cached_serial=None, formula="=A1"
):
    if cached_serial is None:
        cached_serial = serial
    workbook = Workbook()
    if epoch_1904:
        workbook.epoch = datetime.datetime(1904, 1, 1)
    sheet = workbook.active
    sheet.title = "Durations"
    sheet["A1"] = serial
    sheet["B1"] = formula
    sheet["C1"] = "=A1+1"
    for address in ("A1", "B1", "C1"):
        sheet[address].number_format = number_format
    # A genuine date in the same workbook must retain its date metadata.
    sheet["D1"] = "=DATE(2026,2,1)"
    sheet["D1"].number_format = "yyyy-mm-dd"
    original, output = io.BytesIO(), io.BytesIO()
    workbook.save(original)
    with zipfile.ZipFile(original) as source, zipfile.ZipFile(output, "w") as dest:
        for part in source.infolist():
            payload = source.read(part.filename)
            if part.filename == "xl/worksheets/sheet1.xml":
                payload, count = re.subn(
                    rb"(<f>.*?</f>)<v(?:\s*/>|></v>)",
                    lambda match: match[1] + f"<v>{cached_serial}</v>".encode(),
                    payload,
                    count=1,
                )
                assert count == 1
            dest.writestr(part, payload)
    return output.getvalue()


@pytest.mark.parametrize(
    "number_format", ["[h]:mm:ss", "[hh]:mm", "[mm]:ss", "[ss].000"]
)
def test_elapsed_formats_are_not_calendar_formats(number_format):
    assert not _is_date_format(number_format)


@pytest.mark.parametrize(
    "number_format",
    ["yyyy-mm-dd", "hh:mm:ss", 'yyyy-mm-dd "[h]"', r"yyyy-mm-dd \[h\]", '0 "hours"'],
)
def test_calendar_and_clock_classification_is_preserved(number_format):
    assert _is_date_format(number_format) == (number_format != '0 "hours"')


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("number_format", ["[h]:mm:ss", "[mm]:ss", "[ss].000"])
@pytest.mark.parametrize("serial", [0.5, 1.5, 60.5])
def test_duration_inputs_and_formulas_never_get_calendar_labels(
    number_format, serial, epoch_1904, slow_reader, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = duration_workbook(number_format, serial, epoch_1904)
    cached = load_cached_values(data)
    for column in (1, 2, 3):
        assert not cached.is_date("Durations", 1, column)
    assert cached.get("Durations", 1, 1) == datetime.timedelta(days=serial)
    assert cached.get("Durations", 1, 2) == datetime.timedelta(days=serial)
    assert cached.get("Durations", 1, 3) is None

    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    nodes = {node["addr"]: node for node in graph["nodes"]}
    for address, expected in (("A1", serial), ("B1", serial), ("C1", serial + 1)):
        node = nodes[address]
        assert node["value"] == expected
        assert node["valueSource"] == "engine"
        assert "valueDate" not in node
        assert all(sample.get("date") is None for sample in node.get("values", []))
    assert nodes["D1"]["valueDate"] == "2026-02-01"


@pytest.mark.parametrize("serial", [-60.5, -1.5, -0.5, 0, 0.5, 1.5, 60.5, 1_000_000])
@pytest.mark.parametrize(
    ("offset_seconds", "expected"),
    [
        (0, "same"),
        (0.0004, "same"),
        (-0.0004, "same"),
        (0.002, "differ"),
        (1, "differ"),
        (86400, "differ"),
    ],
)
def test_duration_comparison_uses_elapsed_seconds(serial, offset_seconds, expected):
    stored = datetime.timedelta(days=serial)
    assert readings_agree(serial + offset_seconds / 86400, stored, None) == expected


@pytest.mark.parametrize(
    "raw",
    [
        True,
        False,
        None,
        float("nan"),
        float("inf"),
        float("-inf"),
        "1 day, 12:00:00",
        "36:00:00",
        "1.5",
    ],
)
def test_duration_comparison_does_not_coerce_non_numeric_results(raw):
    assert readings_agree(raw, datetime.timedelta(days=1.5), None) == "differ"


@pytest.mark.parametrize("text", ["1 day, 12:00:00", "36:00:00", "1.5"])
def test_duration_shaped_cache_text_is_not_a_duration(text):
    assert readings_agree(1.5, text, None) == "differ"


@pytest.mark.parametrize("epoch_1904", [False, True])
@pytest.mark.parametrize("slow_reader", [False, True])
@pytest.mark.parametrize("serial", [-1.5, 0, 1.5, 60.5])
@pytest.mark.parametrize(
    ("offset_seconds", "expected"),
    [(0, "same"), (0.0004, "same"), (0.002, "differ"), (86400, "differ")],
)
def test_duration_cache_verdict_and_warnings_agree(
    serial, epoch_1904, slow_reader, offset_seconds, expected, monkeypatch
):
    if slow_reader:
        monkeypatch.setattr("linexcel.loader.MAX_DENSE_CELLS", 0)
    data = duration_workbook(
        "[h]:mm:ss.000", serial, epoch_1904, serial + offset_seconds / 86400
    )
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(node for node in graph["nodes"] if node["id"] == "c:Durations!B1")
    assert node["valueSource"] == "engine"
    assert node["cachedAgreement"] == expected
    assert "valueDate" not in node
    assert any(
        "Durations!B1:" in w and "differs from file" in w
        for w in graph["meta"]["warnings"]
    ) == (expected == "differ")


@pytest.mark.parametrize(
    ("formula", "source", "expected"),
    [
        ('="1 day, 12:00:00"', "engine", "differ"),
        ("=TRUE()", "engine", "differ"),
        ("=NOW()", "volatile", "same"),
        ("=B1", "file", "same"),
    ],
)
def test_duration_comparison_preserves_result_type_and_cache_only_provenance(
    formula, source, expected
):
    data = duration_workbook("[h]:mm:ss", 1.5, False, formula=formula)
    graph = analyze_workbook(data, execution=ExecutionPolicy(isolated=False))["graph"]
    node = next(node for node in graph["nodes"] if node["id"] == "c:Durations!B1")
    assert node["valueSource"] == source
    assert node["cachedAgreement"] == expected
    assert node["cachedValue"] == "1 day, 12:00:00"
    assert any(
        "Durations!B1:" in w and "differs from file" in w
        for w in graph["meta"]["warnings"]
    ) == (expected == "differ")


@pytest.mark.parametrize("days", [-999_999_998, -1, 0, 1, 999_999_998])
@pytest.mark.parametrize("offset", [-1000, -501, -500, -250, 0, 250, 500, 501, 1000])
def test_typed_durations_tolerate_only_cache_rounding(days, offset):
    stored = datetime.timedelta(days=days)
    raw = stored + datetime.timedelta(microseconds=offset)
    expected = "same" if abs(offset) <= 500 else "differ"
    assert readings_agree(raw, stored, None) == expected
    assert readings_agree(stored, raw, None) == expected


def test_opposite_duration_limits_differ_without_overflow():
    assert (
        readings_agree(datetime.timedelta.min, datetime.timedelta.max, None) == "differ"
    )
    assert (
        readings_agree(datetime.timedelta.max, datetime.timedelta.min, None) == "differ"
    )
