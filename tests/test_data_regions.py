"""Data regions: formula-free blocks of literal cells, one ``data`` node each.

The detection rule is an heuristic (``data_regions``). These tests pin its
cases on small hand-built workbooks: an Excel table with a calculated column,
a block under a title with a total row below it, a column with a formula in
the middle of its values, a block too small to count, and the link a range
node gets to the region it reads, with the exact sub-range.
"""

from openpyxl import Workbook
from openpyxl.worksheet.table import Table

from linexcel.data_regions import REGION_MIN_CELLS
from linexcel.lazy import evaluate_node
from linexcel.lazy_store import WorkbookIndex, build_index, graph_neighborhood

ROWS = 2500  # 4 columns of 2,500 rows: 10,000 cells, the smallest region


def _build(tmp_path, book):
    path = tmp_path / "book.xlsx"
    book.save(path)
    build_index(path, tmp_path / "index.sqlite")
    return path, tmp_path / "index.sqlite"


def _regions(index):
    with WorkbookIndex(index) as store:
        return {
            node["id"]: node for node in store.values() if node["kind"] == "data"
        }, store.meta


def test_a_table_region_leaves_out_its_calculated_column(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Ledger"
    sheet.append(["Date", "Account", "Amount", "Rate", "Taxed"])
    for row in range(2, ROWS + 2):
        sheet.append(
            [45000 + row, f"A{row % 7}", row * 1.5, 0.2, f"=C{row}*(1+D{row})"]
        )
    sheet.add_table(Table(displayName="tblLedger", ref=f"A1:E{ROWS + 1}"))
    sheet["G1"] = "=SUM(tblLedger[Amount])"
    sheet["G2"] = f"=SUM(C2:C{ROWS + 1})"
    sheet["G3"] = "=SUM(C10:D20)"
    path, index = _build(tmp_path, book)
    regions, meta = _regions(index)
    assert list(regions) == [f"data:Ledger!A2:D{ROWS + 1}"]
    region = regions[f"data:Ledger!A2:D{ROWS + 1}"]
    assert region["regionKind"] == "excel_table"
    assert region["table"] == "tblLedger"
    assert region["headers"] == ["Date", "Account", "Amount", "Rate"]
    assert region["excludedFormulaColumns"] == ["Taxed"]
    assert (region["rowCount"], region["columnCount"]) == (ROWS, 4)
    assert region["cells"] == ROWS * 4
    assert [c["type"] for c in region["schema"]] == [
        "number",
        "text",
        "number",
        "number",
    ]
    assert region["calculable"] is True
    assert meta["dataRegions"]["regions"] == 1
    assert meta["nodeCounts"]["data"] == 1
    with WorkbookIndex(index) as store:
        whole = store[f"Ledger!C2:C{ROWS + 1}"]
        part = store["Ledger!C10:D20"]
    assert whole["dataRegions"] == [
        {
            "node": f"data:Ledger!A2:D{ROWS + 1}",
            "subRange": f"Ledger!C2:C{ROWS + 1}",
            "bounds": [2, 3, ROWS + 1, 3],
            "coversWholeRange": True,
        }
    ]
    assert part["dataRegions"][0]["subRange"] == "Ledger!C10:D20"
    # The link is a relation the neighbourhood shows; the calculation is
    # unchanged and still reads every cell.
    neighbours = {n["id"] for n in graph_neighborhood(index, "Ledger!C10:D20")["nodes"]}
    assert f"data:Ledger!A2:D{ROWS + 1}" in neighbours
    result = evaluate_node(path, "Ledger!G2", index_path=index)
    assert result["status"] == "completed", result["diagnostics"]
    assert result["value"] == sum(row * 1.5 for row in range(2, ROWS + 2))
    assert result["coverage"]["inputCells"] == ROWS


def test_a_block_keeps_its_title_out_and_its_total_row_below(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet["A1"] = "Population by region"
    sheet.append([])
    sheet.append(["Code", "Region", "2023", "2024"])
    for row in range(ROWS):
        sheet.append([f"R{row:05d}", f"Name {row}", row, row + 1])
    total = ROWS + 4
    sheet[f"C{total}"] = f"=SUM(C4:C{total - 1})"
    _, index = _build(tmp_path, book)
    regions, _ = _regions(index)
    # The total under column C does not exclude it (rule ★).
    assert list(regions) == [f"data:Data!A4:D{total - 1}"]
    region = regions[f"data:Data!A4:D{total - 1}"]
    assert region["regionKind"] == "formula_free_columns"
    assert region["headerRow"] == 3
    assert region["headers"] == ["Code", "Region", "2023", "2024"]
    assert region["preambleRows"] == [1, 2]
    assert region["preview"]["rows"][0] == [4, "R00000", "Name 0", 0, 1]


def test_a_column_with_a_formula_among_its_values_stays_per_cell(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Mixed"
    for row in range(1, ROWS + 1):
        sheet.append([row, row, row, row, row])
    sheet["C100"] = "=C99+1"
    _, index = _build(tmp_path, book)
    regions, _ = _regions(index)
    # C is mixed; A:B and D:E are too narrow alone (2 x 2,500 cells).
    assert regions == {}
    book = Workbook()
    sheet = book.active
    sheet.title = "Mixed"
    for row in range(1, ROWS + 1):
        sheet.append([row, row, row, row, row, row])
    sheet["A100"] = "=A99+1"
    tmp = tmp_path / "second"
    tmp.mkdir()
    _, index = _build(tmp, book)
    regions, _ = _regions(index)
    assert list(regions) == [f"data:Mixed!B1:F{ROWS}"]


def test_a_block_under_the_threshold_has_no_region(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Small"
    for row in range(1, ROWS):  # one row short of the threshold
        sheet.append([row, row, row, row])
    _, index = _build(tmp_path, book)
    regions, meta = _regions(index)
    assert (ROWS - 1) * 4 < REGION_MIN_CELLS
    assert regions == {}
    assert meta["dataRegions"]["regions"] == 0
    assert "data" not in meta["nodeCounts"]


def test_a_formula_inside_the_rectangle_cuts_it(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Cut"
    for row in range(1, 3 * ROWS + 1):
        sheet.append([row, row, row, row] if row <= ROWS else [row, row, row])
    # Under D's last value, so D is not mixed; but A:D spans 3 x ROWS rows
    # and this formula is inside: the rectangle is cut at its row.
    sheet[f"D{ROWS + 1}"] = "=SUM(D1:D2)"
    _, index = _build(tmp_path, book)
    regions, _ = _regions(index)
    assert list(regions) == [f"data:Cut!A{ROWS + 2}:D{3 * ROWS}"]
    assert regions[f"data:Cut!A{ROWS + 2}:D{3 * ROWS}"]["schema"][3]["type"] == "empty"
