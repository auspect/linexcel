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


def sales_book():
    """Sales!A1:D5 is ``SalesTable``: a header row, three data rows whose
    ``Total`` column is calculated, and a totals row. Report reads it."""
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


# The resolver is asked once for a reference no cell can resolve, however
# many cells hold the formula: a refusal must not cost one call per cell.


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
        "[@Missing]",  # current row of an unknown column
        "SalesTable[@Missing]",
    ],
)
def test_a_reference_no_cell_can_resolve_is_asked_about_once_or_twice(text):
    resolve, calls = counted(sales_book())
    assert structured_refs.member_rectangles(resolve, text, "Sales", CELLS) == (
        [],
        False,
    )
    assert len(calls) <= 2, calls


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
