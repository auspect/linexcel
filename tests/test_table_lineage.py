"""The graph of ``analyze`` follows structured table references.

Batch calculation already resolved declared table references, but the graph
kept each one as an opaque node: ``=SUM(SalesTable[Qty])`` had no precedent
cell, and a targeted analysis stopped at the reference. The precedents
expected below are read off the workbook layout, cell by cell.
"""

import io
import zipfile

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table

from linexcel import ExecutionPolicy, analyze, structured_refs
from linexcel.refs import Rect


def saved(book):
    stream = io.BytesIO()
    book.save(stream)
    return stream.getvalue()


def sales_book(more=None):
    """Sales!A1:D5 is ``SalesTable``: a header row, three data rows whose
    ``Total`` column is calculated, and a totals row. Report reads it;
    ``more`` adds formulas to Report by address."""
    book = Workbook()
    sales = book.active
    sales.title = "Sales"
    sales.append(["Product", "Qty", "Price", "Total"])
    for product, qty, price in [("a", 2, 10), ("b", 3, 20), ("c", 4, 30)]:
        sales.append([product, qty, price, "=[@Qty]*[@Price]"])
    sales.append(["Total", None, None, "=SUBTOTAL(109,[Total])"])
    sales.add_table(Table(displayName="SalesTable", ref="A1:D5", totalsRowCount=1))
    report = book.create_sheet("Report")
    report["B1"] = "=SUM(SalesTable[Qty])"
    report["B2"] = "=SUM(SalesTable[Total])"
    report["B3"] = "=SalesTable[[#Totals],[Total]]"
    report["B4"] = "=SUM(SalesTable[Missing])"
    for address, formula in (more or {}).items():
        report[address] = formula
    return saved(book)


def graph(data, targets=None):
    result = analyze(data, targets=targets, execution=ExecutionPolicy(isolated=False))
    edges = {(edge["source"], edge["target"]) for edge in result.edges}
    return result, edges


def precedents(edges, node):
    return {source for source, target in edges if target == node}


def test_a_table_column_reference_reads_the_cells_of_the_column():
    result, edges = graph(sales_book())
    # Qty is Sales!B2:B4, three literals; Total is Sales!D2:D4, one copied
    # formula; the totals row of Total is the formula cell Sales!D5.
    assert precedents(edges, "c:Report!B1") == {"i:Sales!B2:B4"}
    assert precedents(edges, "c:Report!B2") == {"g:Sales!D2#3"}
    assert precedents(edges, "c:Report!B3") == {"c:Sales!D5"}
    # The totals row reads its own table's column, named without the table.
    assert precedents(edges, "c:Sales!D5") == {"g:Sales!D2#3"}


def test_a_current_row_reference_reads_its_row_for_each_copied_cell():
    result, edges = graph(sales_book())
    # D2 reads B2 and C2, D3 reads B3 and C3, D4 reads B4 and C4.
    assert precedents(edges, "g:Sales!D2#3") == {"i:Sales!B2:B4", "i:Sales!C2:C4"}


def column_book(formulas, name="Orders"):
    """Sales!A1:B<n> is one table, Qty and Total: ``formulas`` are the Total
    column from row 2 down, and Qty counts from 1."""
    book = Workbook()
    sheet = book.active
    sheet.title = "Sales"
    sheet.append(["Qty", "Total"])
    for qty, formula in enumerate(formulas, 1):
        sheet.append([qty, formula])
    sheet.add_table(Table(displayName=name, ref=f"A1:B{len(formulas) + 1}"))
    return saved(book)


def test_each_copied_formula_reads_the_rows_of_its_own_cells():
    # Written as Excel saves it. B2:B3 hold one formula, B4:B5 another: each
    # group reads Qty on its own two rows, not the whole column.
    twice, thrice = "=Orders[[#This Row],[Qty]]*2", "=Orders[[#This Row],[Qty]]*3"
    result, edges = graph(column_book([twice, twice, thrice, thrice]))
    assert precedents(edges, "g:Sales!B2#2") == {"i:Sales!A2:A3"}
    assert precedents(edges, "g:Sales!B4#2") == {"i:Sales!A4:A5"}


def test_many_separate_rows_of_a_column_are_read_as_one_range():
    # Two formulas alternate down forty rows: each holds every other row. One
    # precedent per row would put forty nodes on the graph; like a copied A1
    # formula, each group reads one range from its first row to its last.
    result, edges = graph(column_book(["=[@Qty]*3", "=[@Qty]*2"] * 20))
    assert precedents(edges, "g:Sales!B2#20") == {"i:Sales!A2:A40"}
    assert precedents(edges, "g:Sales!B3#20") == {"i:Sales!A3:A41"}
    inputs = [node for node in result.nodes if node["kind"] == "input"]
    assert len(inputs) == 2


def test_a_range_between_two_table_references_stays_unresolved():
    # Jan:Mar reads Feb too, which neither end names. Linking the two ends
    # alone would show a formula that reads less than it does.
    book = Workbook()
    sales = book.active
    sales.title = "Sales"
    for values in [("Product", "Jan", "Feb", "Mar"), ("a", 1, 2, 3), ("b", 4, 5, 6)]:
        sales.append(values)
    sales.add_table(Table(displayName="SalesTable", ref="A1:D3"))
    report = book.create_sheet("Report")
    report["B1"] = "=SUM(SalesTable[Jan]:SalesTable[Mar])"
    result, edges = graph(saved(book))
    assert precedents(edges, "c:Report!B1") == {
        "x:SalesTable[Jan]",
        "x:SalesTable[Mar]",
    }


def test_a_reference_to_a_table_without_cells_stays_named():
    # The table declares one data row the sheet holds no cell for: there is
    # no cell to link to, and the formula must not look as if it read nothing.
    book = Workbook()
    sheet = book.active
    sheet.title = "Input"
    sheet.append(["Item", "Qty", "Price"])
    sheet.add_table(Table(displayName="Orders", ref="A1:C2"))
    book.create_sheet("Report")["B1"] = "=SUM(Orders[Qty])"
    result, edges = graph(saved(book))
    assert precedents(edges, "c:Report!B1") == {"x:Orders[Qty]"}


def test_a_targeted_analysis_keeps_a_table_reference_it_did_not_trace():
    # The unknown column makes the engine keep the saved value of B5, so the
    # targeted trace never reaches the Total column: it stays named.
    data = sales_book({"B5": "=SUM(SalesTable[Total])+SUM(SalesTable[Missing])"})
    result, edges = graph(data, targets=["Report!B5"])
    assert precedents(edges, "c:Report!B5") == {
        "x:SalesTable[Total]",
        "x:SalesTable[Missing]",
    }


def test_a_reference_the_resolver_refuses_stays_an_opaque_node():
    result, edges = graph(sales_book())
    assert precedents(edges, "c:Report!B4") == {"x:SalesTable[Missing]"}
    opaque = {node["id"] for node in result.nodes if node["kind"] == "opaque"}
    assert opaque == {"x:SalesTable[Missing]"}


def test_a_targeted_analysis_goes_upstream_through_a_table_reference():
    result, edges = graph(sales_book(), targets=["Report!B2"])
    assert edges == {
        ("g:Sales!D2#3", "c:Report!B2"),
        ("i:Sales!B2:B4", "g:Sales!D2#3"),
        ("i:Sales!C2:C4", "g:Sales!D2#3"),
    }
    assert {node["id"] for node in result.nodes} == {
        "c:Report!B2",
        "g:Sales!D2#3",
        "i:Sales!B2:B4",
        "i:Sales!C2:C4",
    }


def two_tables_book():
    """Two tables side by side on one sheet, with the same calculated column:
    Left is A1:C3 and Right is F1:H3; C2:C3 and H2:H3 hold one copied formula."""
    book = Workbook()
    sheet = book.active
    sheet.title = "Sales"
    for row, values in enumerate([("Qty", "Price", "Total"), (2, 10), (3, 20)], 1):
        for offset, value in enumerate(values):
            sheet.cell(row, 1 + offset, value)
            sheet.cell(row, 6 + offset, value)
    for row in (2, 3):
        sheet.cell(row, 3, "=[@Qty]*[@Price]")
        sheet.cell(row, 8, "=[@Qty]*[@Price]")
    sheet.add_table(Table(displayName="Left", ref="A1:C3"))
    sheet.add_table(Table(displayName="Right", ref="F1:H3"))
    return saved(book)


def test_one_formula_copied_in_two_tables_reads_each_tables_own_columns():
    result, edges = graph(two_tables_book())
    # The four cells are one node; it reads Qty and Price of each table and
    # nothing of the columns between them (C to E).
    assert precedents(edges, "g:Sales!C2#4") == {
        "i:Sales!A2:A3",
        "i:Sales!B2:B3",
        "i:Sales!F2:F3",
        "i:Sales!G2:G3",
    }


def test_two_tables_side_by_side_are_not_a_long_chain():
    # Thirty cells hold one formula, fifteen in each table. None reads another:
    # the graph must not warn about thirty linked calculation steps.
    book = Workbook()
    sheet = book.active
    sheet.title = "Sales"
    for row in range(1, 17):
        values = ("Qty", "Price", "Total") if row == 1 else (row, 10)
        for offset, value in enumerate(values):
            sheet.cell(row, 1 + offset, value)
            sheet.cell(row, 6 + offset, value)
        if row > 1:
            sheet.cell(row, 3, "=[@Qty]*[@Price]")
            sheet.cell(row, 8, "=[@Qty]*[@Price]")
    sheet.add_table(Table(displayName="Left", ref="A1:C16"))
    sheet.add_table(Table(displayName="Right", ref="F1:H16"))
    result, edges = graph(saved(book))
    assert precedents(edges, "g:Sales!C2#30") == {
        "i:Sales!A2:A16",
        "i:Sales!B2:B16",
        "i:Sales!F2:F16",
        "i:Sales!G2:G16",
    }
    assert not [w for w in result.warnings if "Long dependency chains" in w]


def test_a_running_total_down_a_table_column_is_a_long_chain():
    # Each cell reads the one above it through the table: thirty steps.
    formulas = ["=[@Qty]"] + ["=[@Qty]+INDEX([Total],ROW()-2)"] * 30
    result, edges = graph(column_book(formulas))
    assert [w for w in result.warnings if "Long dependency chains" in w]


def test_one_formula_in_two_tables_resolves_where_its_column_exists():
    # Left has a Qty column, Right does not: the cells of Left read their
    # rows, and what the cells of Right asked for stays named.
    book = Workbook()
    sheet = book.active
    sheet.title = "Sales"
    for row in range(1, 7):
        left = ("Qty", "Total") if row == 1 else (row, "=[@Qty]*2")
        right = ("Amount", "Total") if row == 1 else (row, "=[@Qty]*2")
        for offset in (0, 1):
            sheet.cell(row, 1 + offset, left[offset])
            sheet.cell(row, 4 + offset, right[offset])
    sheet.add_table(Table(displayName="Left", ref="A1:B6"))
    sheet.add_table(Table(displayName="Right", ref="D1:E6"))
    result, edges = graph(saved(book))
    assert precedents(edges, "g:Sales!B2#10") == {"i:Sales!A2:A6", "x:[@Qty]"}


def test_a_copy_outside_its_table_keeps_the_reference_unresolved():
    book = Workbook()
    sheet = book.active
    sheet.title = "Sales"
    for values in [("Qty", "Double"), (2, "=[@Qty]*2"), (3, "=[@Qty]*2")]:
        sheet.append(values)
    sheet.add_table(Table(displayName="Orders", ref="A1:B3"))
    sheet["B6"] = "=[@Qty]*2"  # below the table: no current row there
    result, edges = graph(saved(book))
    # B2 and B3 read A2 and A3; B6 reads nothing the graph can name.
    assert precedents(edges, "g:Sales!B2#3") == {"i:Sales!A2:A3", "x:[@Qty]"}


def test_invalid_table_metadata_keeps_opaque_nodes_and_says_why():
    source = zipfile.ZipFile(io.BytesIO(sales_book()))
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            content = source.read(info.filename)
            if info.filename.startswith("xl/tables/"):
                assert b' ref="A1:D5"' in content
                content = content.replace(b' ref="A1:D5"', b' ref="nonsense"', 1)
            target.writestr(info, content)
    result, edges = graph(stream.getvalue())
    assert precedents(edges, "c:Report!B1") == {"x:SalesTable[Qty]"}
    assert any("Table name or rectangle is invalid" in w for w in result.warnings)


# A refusal that cannot depend on the cell holding the formula is asked about
# once, however many cells hold it. That covers a reference that reads the
# same cells from anywhere, and a current-row reference that names its table
# when what is refused is the table or the column, not the row.


def counted(data):
    resolve = structured_refs.batch_table_resolver(data)[0]
    calls = []

    def counting(text, sheet, row=None, column=None):
        calls.append((text, row, column))
        return resolve(text, sheet, row, column)

    return counting, calls


CELLS = [(row, 4) for row in range(2, 5)] + [(row, 9) for row in range(1, 200)]


@pytest.mark.parametrize(
    "text",
    [
        "SalesTable[Missing]",  # unknown column
        "Unknown[Qty]",  # unknown table
        "SalesTable[[#Headers],[#Totals],[Qty]]",  # rows that do not touch
        "SalesTable[@Missing]",  # current row of an unknown column
        "Unknown[@Qty]",  # current row of an unknown table
    ],
)
def test_a_refusal_that_cannot_depend_on_the_cell_is_asked_about_once(text):
    resolve, calls = counted(sales_book())
    assert structured_refs.member_rectangles(resolve, text, "Sales", CELLS) == (
        [],
        False,
    )
    assert len(calls) == 1, calls


@pytest.mark.parametrize(
    "text",
    [
        "[@Qty]",  # which table depends on the cell
        "SalesTable[@Qty]",  # which row depends on the cell
    ],
)
def test_a_refusal_that_depends_on_the_cell_does_not_stop_the_others(text):
    resolve, calls = counted(sales_book())
    # Column 9 is outside the table, on every row from 1 to 199; then come
    # the three cells of the Total column, asked last.
    cells = CELLS[3:] + CELLS[:3]
    rectangles, complete = structured_refs.member_rectangles(
        resolve, text, "Sales", cells
    )
    assert not complete
    assert Rect("Sales", 2, 2, 4, 2) in rectangles


def every_other_row(count):
    """Cells of the Total column (B) of ``column_book``, one row in two."""
    return [(2 + 2 * index, 2) for index in range(count)]


def test_separate_rows_are_kept_apart_up_to_a_limit_then_read_as_one_range():
    resolve, _calls = counted(column_book(["=[@Qty]"] * 60))
    limit = structured_refs.MAX_ROW_RUNS
    # Up to the limit: one rectangle per row, exactly what the cells read.
    exact, complete = structured_refs.member_rectangles(
        resolve, "[@Qty]", "Sales", every_other_row(limit)
    )
    assert complete
    assert exact == [
        Rect("Sales", 2 + 2 * index, 1, 2 + 2 * index, 1) for index in range(limit)
    ]
    # One more: a single range from the first row read to the last.
    wide, complete = structured_refs.member_rectangles(
        resolve, "[@Qty]", "Sales", every_other_row(limit + 1)
    )
    assert complete
    assert wide == [Rect("Sales", 2, 1, 2 + 2 * limit, 1)]


def test_member_rectangles_of_table_references():
    resolve, calls = counted(sales_book())
    cells = [(2, 4), (3, 4), (4, 4)]
    # Not a table reference: a defined name, an ordinary cell.
    assert structured_refs.member_rectangles(resolve, "Rate", "Sales", cells) is None
    assert (
        structured_refs.member_rectangles(resolve, "Report!A1", "Sales", cells) is None
    )
    # A column does not depend on the reading cell: one question.
    del calls[:]
    assert structured_refs.member_rectangles(
        resolve, "SalesTable[Qty]", "Report", cells
    ) == (
        [Rect("Sales", 2, 2, 4, 2)],
        True,
    )
    assert len(calls) == 1
    # The current row: one cell per member, merged where the rows touch.
    assert structured_refs.member_rectangles(resolve, "[@Qty]", "Sales", cells) == (
        [Rect("Sales", 2, 2, 4, 2)],
        True,
    )
    assert structured_refs.member_rectangles(
        resolve, "[@Qty]", "Sales", [(2, 4), (4, 4)]
    ) == (
        [Rect("Sales", 2, 2, 2, 2), Rect("Sales", 4, 2, 4, 2)],
        True,
    )
    # A member on the totals row has no current row: incomplete.
    assert structured_refs.member_rectangles(
        resolve, "[@Qty]", "Sales", [(2, 4), (5, 4)]
    ) == (
        [Rect("Sales", 2, 2, 2, 2)],
        False,
    )
