"""Per-sheet publication within an import's budget (critical review F21).

An import used to publish one index at the end: a time or storage stop lost
every sheet already read. The reader now checks the task's budget before each
sheet and every 10,000 cells. The sheet it stops in is rolled back, and it and
the following sheets are published as not imported, with the reason. A
formula reading such a sheet is refused, never calculated on absent data.
"""

import pytest
from openpyxl import Workbook

from linexcel.lazy import evaluate_node
from linexcel.lazy_store import ImportBudget, WorkbookIndex, build_index


class _StopAfter(ImportBudget):
    """Exhausted from its ``checks + 1``-th check on, with ``reason``."""

    def __init__(self, checks, reason="time_budget"):
        super().__init__()
        self.checks, self.reason, self.calls = checks, reason, 0

    def exhausted(self, index_path, formulas=0):
        self.calls += 1
        return self.reason if self.calls > self.checks else None


def _book(tmp_path, big=0):
    book = Workbook()
    summary = book.active
    summary.title = "Summary"
    summary["A1"] = "=SUM(Data2!A1:A3)"
    summary["A2"] = "=Data1!A1*2"
    summary["A3"] = "=Data2!B1"
    first = book.create_sheet("Data1")
    first["A1"] = 21
    for row in range(1, big + 1):
        first.cell(row + 1, 2, row)
    second = book.create_sheet("Data2")
    for row in (1, 2, 3):
        second.cell(row, 1, row * 100)
    second["B1"] = "text"
    path = tmp_path / "book.xlsx"
    book.save(path)
    return path


def _details(index):
    with WorkbookIndex(index) as store:
        meta = store.meta
        return {s["name"]: s for s in meta["sheetDetails"]}, meta


def test_sheets_past_the_budget_are_published_as_not_imported(tmp_path):
    path = _book(tmp_path)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=_StopAfter(2))  # Summary, Data1, then stop
    details, meta = _details(index)
    assert [details[s]["status"] for s in ("Summary", "Data1", "Data2")] == [
        "complete",
        "complete",
        "not_imported",
    ]
    assert details["Data2"]["reason"] == "time_budget"
    assert "never blank" in details["Data2"]["diagnostics"][0]
    assert meta["status"] == "partial"
    assert {
        (d.get("sheet"), d.get("code"), d.get("reason")) for d in meta["diagnostics"]
    } >= {("Data2", "not_imported", "time_budget")}
    with WorkbookIndex(index) as store:
        assert store.count_cells("Data2", (1, 1, 3, 2)) == 1  # the referenced B1
        assert store["Data2!B1"]["provenance"] == "saved_blank"
        assert "not imported" in store["Data2!B1"]["diagnostics"][0]
        assert "not imported" in store["Data2!A1:A3"]["diagnostics"][0]
    refused = evaluate_node(path, "Summary!A1", index_path=index)
    assert refused["status"] != "completed"
    assert refused["value"] is None
    assert any("not imported" in d for d in refused["diagnostics"])
    assert evaluate_node(path, "Summary!A3", index_path=index)["value"] is None
    read = evaluate_node(path, "Summary!A2", index_path=index)
    assert (read["status"], read["value"]) == ("completed", 42)


def test_a_sheet_stopped_while_read_is_rolled_back(tmp_path):
    path = _book(tmp_path, big=12_000)
    index = tmp_path / "index.sqlite"
    # Summary's check, Data1's, then the check at Data1's 10,000th cell.
    build_index(path, index, budget=_StopAfter(2, "storage_budget"))
    details, _ = _details(index)
    assert details["Data1"]["status"] == "not_imported"
    assert details["Data1"]["reason"] == "storage_budget"
    assert (details["Data1"]["formulas"], details["Data1"]["inputs"]) == (0, 0)
    assert details["Data2"]["status"] == "not_imported"
    with WorkbookIndex(index) as store:
        # Nothing of Data1 was kept but the cell Summary references.
        assert store.count_cells("Data1", (1, 1, 20_000, 2)) == 1
        assert store["Data1!A1"]["provenance"] == "saved_blank"
        assert "Data1" not in store.grid.sheets
        assert "Data1!B2" not in store
    refused = evaluate_node(path, "Summary!A2", index_path=index)
    assert refused["status"] != "completed"


def test_a_task_budget_leaves_time_and_storage_for_the_rest_of_the_import(
    tmp_path,
):
    budget = ImportBudget.for_task(1000.0, 600, storage_bytes=8_000)
    assert budget.deadline == pytest.approx(1000 + 600 - ImportBudget.FINAL_SECONDS)
    assert budget.max_bytes == 4_000
    index = tmp_path / "index.sqlite"
    index.write_bytes(b"x" * 3_000)
    clock = iter([1100.0, 1100.0, 1100.0, 1600.0])
    budget.clock = lambda: next(clock)
    assert budget.exhausted(index) is None
    # 1,100 + 2,000,000 x 0.30 ms, the analysis at its fastest, > 1,595.
    assert budget.exhausted(index, formulas=2_000_000) == "time_budget"
    (tmp_path / "index.sqlite-journal").write_bytes(b"x" * 1_000)
    assert budget.exhausted(index) == "storage_budget"
    assert budget.exhausted(index) == "time_budget"


class _Clock:
    """One second per reading of the clock."""

    def __init__(self):
        self.now = 0.0

    def __call__(self):
        self.now += 1.0
        return self.now


def test_a_time_stop_estimates_the_budget_that_reads_every_sheet(tmp_path):
    """Batch 2 ter: "reimport with a larger time budget" said nothing of how
    large. The import records the reading rate observed until it stopped and
    the budget that would read every sheet at that rate."""
    import math
    import zipfile

    book = Workbook()
    book.active.title = "Small"
    book.active["A1"] = 1
    for name in ("Big1", "Big2"):
        sheet = book.create_sheet(name)
        for row in range(1, 12_501):
            sheet.cell(row, 1, row)
            sheet.cell(row, 2, f"v{row}")
    path = tmp_path / "book.xlsx"
    book.save(path)
    with zipfile.ZipFile(path) as archive:
        sizes = [
            archive.getinfo(f"xl/worksheets/sheet{n}.xml").file_size for n in (1, 2, 3)
        ]
    # Clock readings: start, before Small, before Big1, Big1 at 10,000 and
    # 20,000 cells, before Big2, Big2 at 10,000 cells: past the deadline.
    budget = ImportBudget(deadline=6.5, clock=_Clock(), seconds=100)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=budget)
    details, meta = _details(index)
    assert [details[s]["status"] for s in ("Small", "Big1", "Big2")] == [
        "complete",
        "complete",
        "not_imported",
    ]
    estimate = meta["importEstimate"]
    assert estimate["bytesTotal"] == sum(sizes)
    # Small and Big1 whole, then part of Big2 up to its stop.
    assert sizes[0] + sizes[1] < estimate["bytesRead"] < sum(sizes)
    assert estimate["readSeconds"] == 6.0
    # Past the deadline, the regions of Big1 were skipped too.
    assert estimate["cause"] == ["reading", "regions"]
    phases = estimate["phases"]
    # Reading at the observed rate over the whole workbook.
    assert phases["reading"] == pytest.approx(
        6.0 * sum(sizes) / estimate["bytesRead"], abs=0.05
    )
    # No formula; regions at the code's rate, never measured here.
    assert phases["analysis"] == 0
    assert not estimate["valueRateMeasured"]
    assert estimate["valueSeconds"] == ImportBudget.VALUE_SECONDS
    # The margin over the predicted phases, the final time, to the minute.
    needed = 1.25 * sum(phases.values()) + 5
    assert estimate["seconds"] == max(160, math.ceil(needed / 60) * 60)
    assert estimate["budgetSeconds"] == 100 and not estimate["beyondMaximum"]


class _Readings:
    """A clock that returns the given readings, then the last one."""

    def __init__(self, *readings):
        self.readings = list(readings)

    def __call__(self):
        return self.readings.pop(0) if len(self.readings) > 1 else self.readings[0]


def _formula_book(tmp_path, sheets=3, formulas=1_000):
    book = Workbook()
    book.active.title = "Values"
    book.active["A1"] = 2
    for index in range(1, sheets + 1):
        sheet = book.create_sheet(f"F{index}")
        for row in range(1, formulas + 1):
            sheet.cell(row, 1, f"=Values!A1*{row}")
    path = tmp_path / "formulas.xlsx"
    book.save(path)
    return path


def test_reading_stops_when_the_time_left_no_longer_covers_the_analysis(tmp_path):
    """Lot 2 ter, T3. Reading used to stop at a fixed 60 % of the budget,
    however many formulas it had read: the analysis of those formulas could
    then outlast the task. It now stops before a sheet as soon as the time
    left does not cover the predicted analysis of the formulas read."""
    path = _formula_book(tmp_path)
    budget = ImportBudget.for_task(0.0, 100)
    # 50 ms a formula: after F1, 50 s of reserve; after F2, 100 s.
    budget.reserve_seconds = 0.05
    # Readings: start, before Values, F1, F2, F3 (the reading clock only).
    budget.clock = _Readings(0.0, 1.0, 2.0, 3.0, 4.0)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=budget)
    details, _ = _details(index)
    assert [details[s]["status"] for s in ("Values", "F1", "F2", "F3")] == [
        "complete",
        "complete",
        "complete",
        "not_imported",
    ]
    assert details["F3"]["reason"] == "time_budget"


def test_a_values_sheet_may_take_most_of_the_budget(tmp_path):
    """Lot 2 ter, T4. A sheet of values needs no analysis: its reading may
    use the budget up to the final margin, not only 60 % of it."""
    book = Workbook()
    book.active.title = "Values"
    for row in range(1, 12_501):
        book.active.cell(row, 1, row)
        book.active.cell(row, 2, -row)
    book.create_sheet("After")["A1"] = 1
    path = tmp_path / "values.xlsx"
    book.save(path)
    budget = ImportBudget.for_task(0.0, 100)
    # Readings: start, before Values, its 10,000th and 20,000th cell (70 %
    # of the budget), before After.
    budget.clock = _Readings(0.0, 1.0, 35.0, 70.0, 72.0)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=budget)
    details, meta = _details(index)
    assert [details[s]["status"] for s in ("Values", "After")] == [
        "complete",
        "complete",
    ]
    assert "importEstimate" not in meta


def test_the_reserve_is_the_analysis_at_its_fastest_without_margin():
    """Reprise after the corpus runs. The reserve predicted the analysis at
    the cautious rate with the 1.25 margin (0.81 ms a formula), while the
    energy model analyzes at 0.41 to 0.52 ms: reading stopped with 29 to 40 %
    of the budget unused. It is now the fastest rate measured, 0.30 ms, with
    no margin; the cautious rate stays for proposals that measured none."""
    budget = ImportBudget()
    assert budget.reserve(0) == 0
    assert budget.reserve(1_000) == pytest.approx(0.30)
    assert budget.formula_seconds == pytest.approx(0.65e-3)
    assert ImportBudget(reserve_seconds=0.01).reserve(100) == pytest.approx(1.0)


def test_reading_goes_on_while_the_fastest_analysis_still_fits(tmp_path):
    """Reprise. With 24,000 formulas read and 15 s left before the deadline,
    the cautious reserve (19.5 s) stopped the reading; the analysis at its
    fastest (7.2 s) fits, so the third sheet is read too."""
    path = _formula_book(tmp_path, formulas=12_000)
    budget = ImportBudget.for_task(0.0, 100)  # deadline 95
    # Readings: start, before Values, F1, its 10,000th cell, F2, its
    # 10,000th cell, F3 (24,000 formulas read, at 80 s), its 10,000th cell.
    budget.clock = _Readings(0.0, 1.0, 2.0, 3.0, 60.0, 61.0, 80.0, 81.0, 82.0)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=budget)
    details, _ = _details(index)
    assert [details[s]["status"] for s in ("F1", "F2", "F3")] == ["complete"] * 3


def test_no_estimate_when_every_sheet_is_read(tmp_path):
    path = _book(tmp_path)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=ImportBudget(deadline=1e9, seconds=600))
    _, meta = _details(index)
    assert "importEstimate" not in meta


def test_every_index_records_its_phase_times_and_counts(tmp_path):
    """Lot 2 ter: each index says how long each phase took, on what."""
    path = _book(tmp_path)
    index = tmp_path / "index.sqlite"
    build_index(path, index)
    _, meta = _details(index)
    timings = meta["importTimings"]
    for phase in ("reading", "dependencies", "facts", "vba", "regions", "total"):
        assert timings[phase] >= 0
    assert (timings["formulasRead"], timings["formulasAnalyzed"]) == (3, 3)
    assert timings["valuesRead"] == 5 and not timings["readStopped"]
    assert timings["bytesRead"] == timings["bytesTotal"] > 0
    assert "importEstimate" not in meta


def _timings(**changes):
    timings = {
        "bytesRead": 1_000,
        "bytesTotal": 1_000,
        "readStopped": False,
        "reading": 30.0,
        "formulasRead": 200_000,
        "formulasAnalyzed": 100_000,
        "dependencies": 40.0,
        "valuesRead": 100_000,
        "regionValues": 100_000,
        "regions": 0.5,
        "rest": 1.0,
    }
    timings.update(changes)
    return timings


def test_the_proposal_counts_reading_analysis_and_regions():
    """Lot 2 ter, T2 in figures: the proposal used to count the reading
    alone. It now adds the analysis of every formula, at the rate measured
    on the import, the regions and the rest, with the margin."""
    budget = ImportBudget(seconds=60)
    proposal = budget.proposal(_timings())
    assert proposal["formulaRateMeasured"] and proposal["valueRateMeasured"]
    assert proposal["formulaSeconds"] == pytest.approx(0.4e-3)
    assert proposal["phases"] == {
        "reading": 30.0,
        "analysis": 80.0,
        "regions": 0.5,
        "rest": 1.0,
    }
    # 1.25 x 111.5 + 5 = 144.4 s: 180 s.
    assert proposal["seconds"] == 180 and not proposal["uncertain"]
    # A reading stopped at a quarter: reading and counts times four.
    stopped = budget.proposal(_timings(bytesRead=250, readStopped=True))
    assert stopped["phases"]["reading"] == 120.0
    assert stopped["phases"]["analysis"] == pytest.approx(320.0)
    assert stopped["formulasPredicted"] == 800_000


def test_an_estimate_from_too_little_work_is_uncertain():
    budget = ImportBudget(seconds=25)
    # Less than 20 % of the workbook read.
    assert budget.proposal(_timings(bytesRead=150, readStopped=True))["uncertain"]
    # Too few formulas analyzed to measure a rate that weighs most.
    few = budget.proposal(_timings(formulasAnalyzed=500, dependencies=0.2))
    assert not few["formulaRateMeasured"]
    assert few["formulaSeconds"] == ImportBudget.FORMULA_SECONDS
    assert few["uncertain"]
    assert few["processedShare"] == pytest.approx(500 / 200_000)
    # Regions never described while they weigh most.
    values = budget.proposal(
        _timings(
            formulasRead=0,
            formulasAnalyzed=0,
            valuesRead=10**9,
            regions=0.0,
            regionValues=0,
        )
    )
    assert not values["valueRateMeasured"] and values["uncertain"]
    # Beyond the server's bound, it says so.
    assert values["seconds"] == 3600 and values["beyondMaximum"]


def _formula_sheet(sheet, formulas=25_000):
    sheet.title = "Formulas"
    for row in range(1, formulas + 1):
        sheet.cell(row, 1, f"=ROW()*{row % 7}")


@pytest.mark.parametrize("values_first", [False, True])
def test_a_formula_sheet_that_can_be_read_is_not_rolled_back(tmp_path, values_first):
    """Lot 2 ter, reprise. Inside a sheet, the reserve counted the formulas of
    that very sheet: a sheet that could be read within the budget was rolled
    back and published "not imported" as soon as their predicted analysis
    no longer fitted. The analysis safety net would have published all its
    values and analyzed what time allowed. Inside a sheet the reserve now
    protects only the sheets already read whole."""
    book = Workbook()
    if values_first:
        book.active.title = "Values"
        book.active["A1"] = 1
        _formula_sheet(book.create_sheet())
    else:
        _formula_sheet(book.active)
    path = tmp_path / "formulas.xlsx"
    book.save(path)
    budget = ImportBudget.for_task(0.0, 30)  # deadline 25
    # Readings: start, before each sheet, Formulas at its 10,000th and
    # 20,000th cell: 10 s, when 20,000 formulas predict 16.25 s of analysis.
    readings = [0.0, 1.0] + ([2.0] if values_first else []) + [3.0, 10.0]
    budget.clock = _Readings(*readings)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=budget)
    details, meta = _details(index)
    assert details["Formulas"]["status"] == "complete", details["Formulas"]
    assert details["Formulas"]["formulas"] == 25_000
    # The clock stays at 10 s: the analysis had the time to finish.
    assert meta["importTimings"]["formulasNotAnalyzed"] == 0
    assert "importEstimate" not in meta


def test_what_has_no_safety_net_stays_cautious():
    """Reprise. Region description cannot be stopped once started and the
    final steps cannot be stopped at all: a miss there publishes nothing.
    They keep the cautious rate with the margin, and the 5 s final time."""
    budget = ImportBudget(deadline=100.0, clock=lambda: 0.0)
    # 16 M values: 88 s at 5.5 us, 110 s with the 1.25 margin.
    assert not budget.regions_fit(16_000_000)
    assert budget.regions_fit(14_000_000)  # 96.25 s with the margin
    assert ImportBudget.FINAL_SECONDS == 5.0
    assert ImportBudget.for_task(0.0, 60).deadline == 55.0


def test_the_proposed_budget_reads_every_sheet_when_reimported(tmp_path):
    """Reprise, the invariant that was missing. On the corpus, 25 s proposed
    180 s, and the reimport with 180 s still stopped reading. The proposal
    used the analysis rate measured on the import, the reserve its own: the
    proposed budget, less the final time, must cover the whole reading plus
    the reserve of every formula. Reimported with the proposal, at the same
    simulated rates, the workbook is read whole."""
    path = _formula_book(tmp_path, sheets=6, formulas=12_000)

    def run(seconds, name):
        budget = ImportBudget.for_task(0.0, seconds)
        budget.reserve_seconds = 2e-3  # 24 s of reserve a sheet
        budget.clock = _Clock()  # 1 s a reading: reading, analysis, regions
        index = tmp_path / name
        build_index(path, index, budget=budget)
        return _details(index)

    details, meta = run(60, "first.sqlite")
    assert any(d["status"] == "not_imported" for d in details.values())
    proposed = meta["importEstimate"]["seconds"]
    details, _ = run(proposed, "again.sqlite")
    assert {d["status"] for d in details.values()} - {"partial"} == {"complete"}, (
        proposed,
        {n: d["status"] for n, d in details.items()},
    )


def test_the_final_margin_grows_with_the_formulas_read(tmp_path):
    """Diagnostic of the end of an import stopped by the safety net. Marking
    the sheets, links, regions, counts and publication take longer as the
    formulas grow: 1.2 s after the deadline for 150,000 formulas, 3.3 s for
    450,000, past the 5 s kept for a million. The final margin now grows
    with the formulas read (never under 5 s), and the analysis stops that
    much earlier."""
    path = _formula_book(tmp_path, sheets=1, formulas=1_000)
    budget = ImportBudget.for_task(0.0, 60)  # 55 s with the 5 s margin
    # 10 ms of final work a formula here: 1,000 formulas keep 10 s.
    budget.final_seconds_per_formula = 0.01
    # Reading early; the analysis then reads 52 s, past 60 - 10 s.
    budget.clock = _Readings(0.0, 1.0, 2.0, 52.0)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=budget)
    details, meta = _details(index)
    assert details["F1"]["status"] == "partial"
    assert meta["importTimings"]["formulasNotAnalyzed"] == 1_000
    assert budget.final_seconds() == pytest.approx(10.0)
    # Never under the fixed 5 s.
    assert ImportBudget().final_seconds() == ImportBudget.FINAL_SECONDS
