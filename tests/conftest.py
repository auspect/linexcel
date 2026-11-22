import io

import pytest


def build_lineage_workbook() -> bytes:
    """Rich workbook for lineage tests: stretched formulas,
    cross-sheet references, defined name, composite functions."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws["A1"], ws["B1"], ws["C1"], ws["D1"] = "Product", "Qty", "Price", "Revenue"
    for r in range(2, 102):
        ws.cell(row=r, column=1, value=f"P{r - 1}")
        ws.cell(row=r, column=2, value=r % 7 + 1)
        ws.cell(row=r, column=3, value=10.5 + (r % 13))
        ws.cell(row=r, column=4, value=f"=B{r}*C{r}")  # stretched 100 times
    ws["A1"].comment = Comment("Exported product category", "Data team")
    ws.freeze_panes = "A2"
    ws.column_dimensions["C"].hidden = True
    ws.merge_cells("F1:G1")
    ws["F1"] = "Presentation context"

    syn = wb.create_sheet("Summary")
    syn["B1"] = "=SUM(Sales!D2:D101)"
    syn["B2"] = "=ROUND(AVERAGE(Sales!D2:D101), 2)"
    syn["B3"] = (
        "=IF(SUM(Sales!D2:D101)>TargetRate, "
        'CONCATENATE("OK: ", ROUND(B1/1000,1), "k"), "KO")'
    )

    params = wb.create_sheet("Params")
    params["A1"] = 5000
    wb.defined_names.add(DefinedName("TargetRate", attr_text="Params!$A$1"))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture()
def lineage_excel() -> bytes:
    return build_lineage_workbook()


@pytest.fixture(autouse=True, scope="session")
def slow_browser():
    """Run every browser page on a slowed processor when asked to.

    ``LINEXCEL_CPU_THROTTLE=6`` divides the speed of each Chromium page by
    six. A test that reads the page before the page has finished passes on a
    quiet machine and fails on a busy CI runner; slowed, it fails here too,
    every time. Use it to reproduce such a failure and to prove its fix.
    """
    import os

    rate = float(os.environ.get("LINEXCEL_CPU_THROTTLE") or 1)
    if rate <= 1:
        yield
        return
    from playwright.sync_api import Browser

    new_page = Browser.new_page

    def throttled(self, *args, **kwargs):
        page = new_page(self, *args, **kwargs)
        session = page.context.new_cdp_session(page)
        session.send("Emulation.setCPUThrottlingRate", {"rate": rate})
        return page

    Browser.new_page = throttled
    try:
        yield
    finally:
        Browser.new_page = new_page
