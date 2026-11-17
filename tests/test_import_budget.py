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

    def exhausted(self, index_path):
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
    assert budget.deadline == pytest.approx(1000 + 0.6 * 600)
    assert budget.max_bytes == 4_000
    index = tmp_path / "index.sqlite"
    index.write_bytes(b"x" * 3_000)
    clock = iter([1100.0, 1100.0, 1400.0])
    budget.clock = lambda: next(clock)
    assert budget.exhausted(index) is None
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
    needed = 6.0 * sum(sizes) / estimate["bytesRead"]
    assert estimate["seconds"] == max(
        160, math.ceil(needed / ImportBudget.READ_TIME_SHARE * 1.25 / 60) * 60
    )
    assert estimate["budgetSeconds"] == 100 and not estimate["beyondMaximum"]
    # The reading share, so that the interface can explain the numbers.
    assert estimate["readShare"] == ImportBudget.READ_TIME_SHARE


def test_no_estimate_when_every_sheet_is_read(tmp_path):
    path = _book(tmp_path)
    index = tmp_path / "index.sqlite"
    build_index(path, index, budget=ImportBudget(deadline=1e9, seconds=600))
    _, meta = _details(index)
    assert "importEstimate" not in meta
