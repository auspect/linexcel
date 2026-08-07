"""Tests for value resolution: guarded errors, resilience, dates, provenance."""

import datetime
import io
from typing import Any

import pytest
from openpyxl import Workbook

from linexcel.analyzer import (
    CachedValues,
    _Budget,
    _ValueResolver,
    analyze_workbook,
    load_cached_values,
    serial_to_date_text,
)

SHEET = "S"


def build(
    cells: dict[str, Any], formats: dict[str, str] | None = None, sheet: str = SHEET
) -> bytes:
    """Minimal in-memory workbook: one sheet, the given cells and formats."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for addr, value in cells.items():
        ws[addr] = value
    for addr, number_format in (formats or {}).items():
        ws[addr].number_format = number_format
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def graph_of(cells: dict[str, Any], formats: dict[str, str] | None = None) -> dict:
    return analyze_workbook(build(cells, formats), "values.xlsx")["graph"]


def node_of(graph: dict, node_id: str) -> dict:
    return next(n for n in graph["nodes"] if n["id"] == node_id)


class _StubEngine:
    """Engine stand-in: only reports stored values, never recalculates."""

    def __init__(self, values: dict[tuple[str, int, int], Any]):
        self._values = values

    def get_value(self, sheet: str, row: int, col: int) -> Any:
        return self._values.get((sheet, row, col))

    def get_formula(self, sheet: str, row: int, col: int) -> None:
        return None


def resolver_for(engine_values: dict, cached: CachedValues, warnings: list) -> tuple:
    resolver = _ValueResolver(
        _StubEngine(engine_values),
        {SHEET},
        cached,
        warnings,
        _Budget(0),
        scratch_ready=False,
    )
    return resolver


class TestGuardedErrors:
    def test_missing_sheet_guard_falls_back_to_the_second_argument(self):
        graph = graph_of({"A1": "=IFERROR(NOSHEET!A1, 456)"})
        node = node_of(graph, "c:S!A1")
        assert node["value"] == 456
        assert node["valueSource"] == "fallback"

    def test_native_error_guard_is_evaluated_by_the_engine(self):
        graph = graph_of({"A1": "=IFERROR(1/0, 42)"})
        node = node_of(graph, "c:S!A1")
        assert node["value"] == 42
        assert node["valueSource"] == "engine"

    def test_blank_operand_counts_as_zero(self):
        # A2 stays empty: Excel coerces the blank to 0, so the guard never fires
        graph = graph_of({"B1": "=IFERROR(A2+1, 7)"})
        assert node_of(graph, "c:S!B1")["value"] == 1

    def test_iserror_branch_is_evaluated(self):
        graph = graph_of({"A1": "=IF(ISERROR(1/0), 99, 0)"})
        assert node_of(graph, "c:S!A1")["value"] == 99

    def test_unguarded_missing_sheet_does_not_break_the_analysis(self):
        graph = graph_of({"A1": "=NOSHEET!A1"})
        node = node_of(graph, "c:S!A1")
        assert node["value"] is None
        assert "valueSource" not in node
        warnings = graph["meta"]["warnings"]
        assert any("Global evaluation incomplete" in w for w in warnings)


class TestResilience:
    """One bad reference used to blank out every formula of the workbook."""

    def test_good_formulas_keep_their_values_next_to_a_broken_one(self):
        graph = graph_of(
            {
                "A1": 1,
                "B1": 2,
                "A2": "=A1*2",
                "A3": "=IFERROR(NOSHEET!A1, 456)",
                "A4": "=SUM(A1:B1)",
            }
        )
        doubled = node_of(graph, "c:S!A2")
        summed = node_of(graph, "c:S!A4")
        assert doubled["value"] == 2
        assert doubled["valueSource"] == "engine"
        assert summed["value"] == 3
        assert summed["valueSource"] == "engine"

    def test_the_graph_still_describes_every_formula(self):
        graph = graph_of({"A1": 1, "A2": "=A1*2", "A3": "=NOSHEET!A1"})
        ids = {n["id"] for n in graph["nodes"]}
        assert {"c:S!A2", "c:S!A3"} <= ids
        assert graph["meta"]["stats"]["totalFormulas"] == 2


def build_multi(sheets: dict[str, dict[str, Any]]) -> bytes:
    """In-memory workbook with several sheets, in the given order."""
    wb = Workbook()
    first = True
    for title, cells in sheets.items():
        ws = wb.active if first else wb.create_sheet(title)
        ws.title = title
        first = False
        for addr, value in cells.items():
            ws[addr] = value
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# One unguarded broken reference makes the whole-workbook evaluation fail, and
# every value is then recovered cell by cell. A scratch evaluation only reads
# constants, so a formula pointing at another formula used to read 0: =A3+A4
# came back as 0 instead of 230, and =SUM(A1:A7) counted the two constants
# only. The precedents are now recovered first and fed back into the engine.
CHAINED = {
    "Inputs": {
        "A1": 10,
        "A2": 20,
        "A3": "=A1+A2",
        "A4": "=A1*A2",
        "A5": "=IFERROR(1/0,42)",
        "A6": "=IFERROR(NOSHEET!A1,456)",
        "A7": "=A3+A4",
        "B1": "=SUM(A1:A2)",
        "B2": "=IF(ISERROR(1/0),99,0)",
        "C1": 5,
        "C2": "=C1*2",
        "D1": 7,
        "D2": "=D1+1",
        "E1": datetime.date(2026, 8, 7),
        "E2": "=E1+1",
        "E3": "=E1+7",
        "E4": datetime.date(1999, 12, 31),
    },
    "Calc": {
        "A1": 100,
        "A2": "=Inputs!A1*10",
        "A3": "=Inputs!E1",
        "B1": "=SUM(Inputs!A1:A7)",
        "B2": "=Calc!A2+Calc!A1",
        "C1": "=IFERROR(Inputs!A5+1,0)",
        "C2": "=A1/B1",
        "D1": "=Inputs!E1+1",
    },
    "Errors": {
        "A1": "=1/0",
        "A2": "=NOSHEET!A1",
        "A3": "=IFERROR(A1+1,7)",
        "B1": 1,
        "B2": "=B1+B2",
    },
}


class TestChainedRefs:
    """A formula reading another formula, across sheets and through ranges."""

    def graph(self) -> dict:
        return analyze_workbook(build_multi(CHAINED), "chained.xlsx")["graph"]

    def test_same_sheet_chain_sums_the_computed_precedents(self):
        # =A3+A4 with A3==A1+A2 and A4==A1*A2: 30 + 200
        node = node_of(self.graph(), "c:Inputs!A7")
        assert node["value"] == 230
        assert node["valueSource"] == "engine"

    def test_range_over_formula_cells_is_summed_in_full(self):
        # Inputs!A1:A7 = 10 + 20 + 30 + 200 + 42 + 456 + 230, the guarded A6
        # contributing the value of its fallback branch
        node = node_of(self.graph(), "c:Calc!B1")
        assert node["value"] == 988
        assert node["valueSource"] == "engine"

    def test_cross_sheet_reference_to_a_guarded_formula(self):
        node = node_of(self.graph(), "c:Calc!C1")
        assert node["value"] == 43
        assert node["valueSource"] == "engine"

    def test_division_by_a_chained_range_sum(self):
        node = node_of(self.graph(), "c:Calc!C2")
        assert node["value"] == pytest.approx(100 / 988, abs=1e-6)
        assert node["valueSource"] == "engine"

    def test_guard_reading_a_cell_that_holds_an_error(self):
        # Errors!A1 is =1/0: the error travels to A3, whose guard catches it
        node = node_of(self.graph(), "c:Errors!A3")
        assert node["value"] == 7
        assert node["valueSource"] == "engine"

    def test_cross_sheet_constant_and_qualified_self_reference(self):
        graph = self.graph()
        assert node_of(graph, "c:Calc!A2")["value"] == 100
        assert node_of(graph, "c:Calc!B2")["value"] == 200

    def test_date_serial_travels_across_sheets(self):
        graph = self.graph()
        assert node_of(graph, "c:Calc!A3")["value"] == 46241
        assert node_of(graph, "c:Calc!D1")["value"] == 46242

    def test_the_static_date_keeps_its_valuedate(self):
        assert node_of(self.graph(), "i:Inputs!E1")["valueDate"] == "2026-08-07"

    def test_the_error_cell_stays_an_error(self):
        node = node_of(self.graph(), "c:Errors!A1")
        assert "'type': 'Error'" in node["value"]
        assert "'kind': 'Div'" in node["value"]

    def test_the_unguarded_broken_reference_claims_no_value(self):
        node = node_of(self.graph(), "c:Errors!A2")
        assert node["value"] is None
        assert "valueSource" not in node

    def test_a_self_referencing_formula_does_not_crash(self):
        # =B1+B2 written in B2: the cycle is cut, B2 reads as blank
        assert node_of(self.graph(), "c:Errors!B2")["value"] == 1

    def test_the_recovery_is_reported_in_the_warnings(self):
        warnings = self.graph()["meta"]["warnings"]
        assert any("Global evaluation incomplete" in w for w in warnings)
        assert any(w.startswith("Values recovered cell by cell:") for w in warnings)


class TestChainedViaGuarded:
    def test_guard_catches_the_error_of_a_referenced_cell(self):
        graph = graph_of({"A1": "=1/0", "A2": "=IFERROR(A1+1, 7)", "Z1": "=NOSHEET!A1"})
        assert node_of(graph, "c:S!A2")["value"] == 7

    def test_an_error_inside_a_range_poisons_the_sum(self):
        # Excel does not skip errors in SUM: SUM({#DIV/0!, 5}) is #DIV/0!, so
        # the guard fires and the fallback branch — not 6 — is what is shown
        graph = graph_of(
            {
                "A1": "=1/0",
                "A2": 5,
                "A3": "=IFERROR(SUM(A1:A2)+1, 0)",
                "Z1": "=NOSHEET!A1",
            }
        )
        node = node_of(graph, "c:S!A3")
        assert node["value"] == 0
        assert node["valueSource"] == "engine"


class TestSiblingIsolation:
    """One unguarded broken reference must not touch the cells around it."""

    def test_good_formulas_survive_on_every_sheet(self):
        graph = analyze_workbook(
            build_multi(
                {
                    "A": {"A1": 3, "A2": "=A1*4", "A3": "=A2+A1"},
                    "B": {"B1": "=NOSHEET!A1", "B2": "=A!A3*2"},
                    "C": {"C1": "=SUM(A!A1:A3)"},
                }
            ),
            "isolation.xlsx",
        )["graph"]
        assert node_of(graph, "c:A!A2")["value"] == 12
        assert node_of(graph, "c:A!A3")["value"] == 15
        assert node_of(graph, "c:B!B2")["value"] == 30
        assert node_of(graph, "c:C!C1")["value"] == 30
        assert node_of(graph, "c:B!B1")["value"] is None


class TestNestedGuard:
    def test_only_the_outer_guard_is_recovered(self):
        # Excel lets the inner IFERROR absorb the missing sheet — SUM(0) — and
        # returns 0. The engine cannot evaluate the expression at all, so the
        # recovery falls back to the branch of the outer guard. Documented
        # ceiling of the recovery, asserted here so a change is noticed.
        graph = graph_of({"A1": "=IFERROR(SUM(IFERROR(NOSHEET!A1,0)),1)"})
        node = node_of(graph, "c:S!A1")
        assert node["value"] == 1
        assert node["valueSource"] == "fallback"


class TestDates:
    def test_static_date_cell_is_reported_as_a_date(self):
        graph = graph_of({"A1": datetime.date(2026, 8, 7), "B1": "=A1+1"})
        node = node_of(graph, "i:S!A1")
        assert node["valueDate"] == "2026-08-07"
        # a midnight datetime serializes as the bare date, never with a time
        assert node["cachedValue"] == "2026-08-07"
        # the engine hands back the raw serial, which maps to the same date
        assert serial_to_date_text(node["value"]) == "2026-08-07"

    def test_date_formula_is_converted(self):
        graph = graph_of(
            {"A1": datetime.date(2026, 8, 7), "A2": "=A1+1"},
            {"A2": "yyyy-mm-dd"},
        )
        assert node_of(graph, "c:S!A2")["valueDate"] == "2026-08-08"

    def test_plain_number_has_no_date(self):
        graph = graph_of({"A1": 7, "A2": "=A1*2"})
        node = node_of(graph, "c:S!A2")
        assert node["value"] == 14
        assert "valueDate" not in node

    def test_serial_conversion(self):
        assert serial_to_date_text(46241.0) == "2026-08-07"
        assert serial_to_date_text(61.0) == "1900-03-01"
        # Jan/Feb 1900 are real dates: serial 1 is 1900-01-01
        assert serial_to_date_text(1.0) == "1900-01-01"
        assert serial_to_date_text(59.0) == "1900-02-28"
        # 1900's phantom leap day: no real date, so nothing is claimed
        assert serial_to_date_text(60.0) is None
        assert serial_to_date_text(0.0) is None
        assert serial_to_date_text(0.0, epoch_1904=True) == "1904-01-01"
        assert serial_to_date_text("2026-08-07") is None
        assert serial_to_date_text(True) is None

    def test_timestamp_keeps_its_time_of_day(self):
        warnings: list[str] = []
        cached = CachedValues(
            {(SHEET, 1, 1): datetime.datetime(2026, 8, 7, 15, 30)},
            {(SHEET, 1, 1)},
            False,
        )
        resolver = resolver_for({}, cached, warnings)
        assert resolver.describe(SHEET, 1, 1)["cachedValue"] == "2026-08-07 15:30:00"

    def test_cached_values_are_loaded_from_the_file(self):
        cached = load_cached_values(
            build({"A1": datetime.date(2026, 8, 7), "A2": 3}, {"A2": "0.00"})
        )
        assert cached.get(SHEET, 1, 1) == datetime.datetime(2026, 8, 7)
        assert cached.get(SHEET, 2, 1) == 3
        assert cached.is_date(SHEET, 1, 1)
        assert not cached.is_date(SHEET, 2, 1)
        assert cached.epoch_1904 is False


class TestProvenance:
    def test_formula_node_reports_the_engine(self):
        graph = graph_of({"A1": 2, "A2": "=A1*3"})
        assert node_of(graph, "c:S!A2")["valueSource"] == "engine"

    def test_guarded_node_reports_the_fallback(self):
        graph = graph_of({"A1": 1, "A2": "=IFERROR(NOSHEET!A1, 456)"})
        assert node_of(graph, "c:S!A2")["valueSource"] == "fallback"

    def test_static_cell_carries_the_file_value(self):
        graph = graph_of({"A1": datetime.date(2026, 8, 7), "B1": "=A1+1"})
        node = node_of(graph, "i:S!A1")
        assert node["cachedValue"] == "2026-08-07"
        assert node["values"][0]["source"] == "engine"
        assert node["values"][0]["date"] == "2026-08-07"

    def test_range_samples_carry_source_and_date(self):
        graph = graph_of({"A1": 1, "A2": 2, "A3": "=SUM(A1:A2)"})
        samples = node_of(graph, "i:S!A1:A2")["values"]
        assert [s["addr"] for s in samples] == ["A1", "A2"]
        assert {s["source"] for s in samples} == {"engine"}
        assert all(s["date"] is None for s in samples)

    def test_step_inputs_carry_the_date(self):
        graph = graph_of(
            {"A1": datetime.date(2026, 8, 7), "A2": "=A1+1"},
            {"A2": "yyyy-mm-dd"},
        )
        steps = node_of(graph, "c:S!A2")["steps"]
        assert steps["inputs"][0] == {
            "ref": "A1",
            "value": 46241.0,
            "date": "2026-08-07",
        }

    # A workbook written by openpyxl carries no cached value for formula cells,
    # and its constants always agree with what the engine recomputes, so the
    # divergence warning cannot be produced from a built file: it is exercised
    # here on the resolver itself, which is where the comparison happens.
    def test_date_divergence_is_warned(self):
        warnings: list[str] = []
        cached = CachedValues(
            {(SHEET, 1, 1): datetime.datetime(2026, 8, 7)}, {(SHEET, 1, 1)}, False
        )
        resolver = resolver_for({(SHEET, 1, 1): 46240.0}, cached, warnings)
        value, source, date_text = resolver.value(SHEET, 1, 1)
        assert (value, source, date_text) == (46240.0, "engine", "2026-08-06")
        assert warnings == [
            "S!A1: recalculated 46240.0 differs from file value 2026-08-07"
        ]

    def test_number_divergence_respects_the_float_tolerance(self):
        warnings: list[str] = []
        cached = CachedValues({(SHEET, 1, 1): 5.0}, set(), False)
        resolver = resolver_for({(SHEET, 1, 1): 5.0000000001}, cached, warnings)
        assert resolver.value(SHEET, 1, 1)[0] == 5.0000000001
        assert warnings == []

        warnings.clear()
        resolver = resolver_for({(SHEET, 1, 1): 6.0}, cached, warnings)
        assert resolver.value(SHEET, 1, 1)[0] == 6.0
        assert warnings == ["S!A1: recalculated 6.0 differs from file value 5.0"]

    def test_file_value_is_used_when_the_engine_has_nothing(self):
        warnings: list[str] = []
        cached = CachedValues({(SHEET, 2, 1): "KO"}, set(), False)
        resolver = resolver_for({}, cached, warnings)
        assert resolver.value(SHEET, 2, 1) == ("KO", "file", None)
        assert warnings == []
