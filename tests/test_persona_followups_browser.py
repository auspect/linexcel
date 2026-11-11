"""Follow-ups from the 2026-11-10 first-user trial, on a real local server."""

import io

import pytest
from openpyxl import Workbook
from test_library_search_transitions import book_bytes, scope_ready
from test_library_search_transitions import browser as browser
from test_library_search_transitions import live_library as live_library

playwright = pytest.importorskip("playwright.sync_api")


def model_bytes():
    book = Workbook()
    sheet = book.active
    sheet.title = "Summary"
    sheet.append(["Region", "Revenue"])
    for region, value in [("North", 100), ("South", 200), ("West", 300)]:
        sheet.append([region, value])
    sheet["B6"] = "=SUM(B2:B4)"
    sheet["B7"] = "=B6*2"
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


def test_escape_leaves_search_after_focus_moves_to_the_page(browser, live_library):
    app, url, upload = live_library
    project = upload("alpha.xlsx", book_bytes())
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    try:
        page.goto(f"{url}/?project={project['id']}")
        page.locator(".library-search-trigger").click()
        scope_ready(page, url)
        # A click on plain text leaves focus on the page body, not in the view.
        page.locator(".library-search-heading p").click()
        page.evaluate("document.activeElement.blur()")
        assert page.evaluate("document.activeElement === document.body")
        page.keyboard.press("Escape")
        playwright.expect(page.locator("#library-search-view")).to_be_hidden()
        playwright.expect(page.locator(".library-search-trigger")).to_be_focused()
    finally:
        page.close()


@pytest.mark.parametrize("width", [1440, 390])
def test_see_in_graph_opens_the_cell_neighbourhood_not_the_overview(
    browser, live_library, width
):
    app, url, upload = live_library
    project = upload("model.xlsx", model_bytes())
    page = browser.new_page(
        viewport={"width": width, "height": 844 if width == 390 else 900}
    )
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    try:
        page.goto(f"{url}/?project={project['id']}")
        # Visit the graph first, as the trial did: a lazy workbook may open on
        # its overview, which then drilled a later selection into groups.
        page.locator("#tab-graph").click()
        page.locator("#graph-fit").click()
        # The overview is active: it offers the way back to a neighbourhood.
        playwright.expect(page.locator("#graph-neighborhood")).to_be_visible(
            timeout=20000
        )
        page.locator("#tab-nodes").click()
        page.locator("#search").fill("B7")
        page.locator('#node-list [data-node="Summary!B7"]').click()
        page.locator("#detail-graph").click()
        playwright.expect(page.locator("#view-graph")).to_be_visible()
        # The explicit request for this cell shows its neighbourhood: the
        # overview's "back to neighbourhood" control is not offered.
        playwright.expect(page.locator("#graph-neighborhood")).to_be_hidden()
        nodes = page.locator(".react-flow__node")
        playwright.expect(nodes.filter(has_text="Summary · B7")).to_have_count(1)
        # Its precedent is a card of its own, not hidden in a "rest of workbook" group.
        playwright.expect(nodes.filter(has_text="Summary · B6")).to_have_count(1)
        assert not nodes.filter(has_text="Reste du classeur").count()
        assert not errors
    finally:
        page.close()


def wide_bytes():
    book = Workbook()
    sheet = book.active
    sheet.title = "Wide"
    for row in range(1, 81):
        for column in range(1, 41):
            sheet.cell(row, column, row * column)
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


@pytest.mark.parametrize("size", [(824, 768), (1440, 900), (390, 844)])
def test_data_grid_and_its_scrollbars_fit_in_the_panel(browser, live_library, size):
    app, url, upload = live_library
    project = upload("wide.xlsx", wide_bytes())
    page = browser.new_page(viewport={"width": size[0], "height": size[1]})
    try:
        page.goto(f"{url}/?project={project['id']}")
        page.locator("#tab-data").click()
        if page.locator("#data-options").is_visible():
            page.locator("#data-options > summary").click()
        page.locator("#data-sheet").select_option(label="Wide")
        playwright.expect(page.locator("#excel-window table")).to_be_visible()
        page.evaluate(
            "document.getElementById('excel-window').scrollIntoView({block:'start'})"
        )
        box = page.evaluate("""() => {
          const grid = document.getElementById('excel-window').getBoundingClientRect();
          const panel = document.getElementById('view-data');
          return {top: grid.top, bottom: grid.bottom, height: grid.height,
            viewport: innerHeight, panel: panel.clientHeight,
            scrollsX: document.getElementById('excel-window').scrollWidth
              > document.getElementById('excel-window').clientWidth};
        }""")
        # The grid scrolls sideways. Its scrollbar is at its bottom edge, which
        # must be on screen together with its top: at 824x768 it never was.
        assert box["scrollsX"]
        assert box["height"] <= box["panel"]
        assert box["top"] >= 0 and box["bottom"] <= box["viewport"]
    finally:
        page.close()


def test_retention_line_follows_the_interface_language(browser, live_library):
    app, url, upload = live_library
    project = upload("alpha.xlsx", book_bytes())
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    try:
        page.goto(f"{url}/?project={project['id']}")
        retention = page.locator("#retention")
        playwright.expect(retention).to_contain_text("Limite des fichiers importés : ")
        page.locator("#ui-language").select_option("en")
        # The line is composed from numbers, so the page dictionary cannot
        # retranslate it: it stayed French until the page was reloaded.
        playwright.expect(retention).to_contain_text("Upload limit: ")
        assert "Limite" not in retention.text_content()
        page.locator("#ui-language").select_option("fr")
        playwright.expect(retention).to_contain_text("Limite des fichiers importés : ")
    finally:
        page.close()


def kpi_bytes():
    book = Workbook()
    sheet = book.active
    sheet.title = "Summary"
    sheet["B5"], sheet["C5"] = "Indicateur", "Réel"
    sheet["B8"], sheet["C8"] = "Marge Brute Commerciale", 500000
    for row in range(21, 41):
        sheet.cell(row, 2, f"Ligne {row}")
        sheet.cell(row, 3, row * 10)
    sheet["B45"], sheet["C45"] = "Total des lignes", "=SUM(C21:C40)"
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


def test_label_search_shows_each_occurrence_value_and_a_short_member_list(
    browser, live_library
):
    app, url, upload = live_library
    project = upload("kpi.xlsx", kpi_bytes())
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    try:
        page.goto(f"{url}/?project={project['id']}")
        page.locator(".library-search-trigger").click()
        scope_ready(page, url)
        query = page.locator("#library-search-query")
        query.fill("Marge Brute Commerciale")
        query.press("Enter")
        card = page.locator(".library-result").first
        playwright.expect(card).to_be_visible()
        buttons = card.locator(".library-occurrences button, button").all_inner_texts()
        # The number sits beside its address; the label does not repeat itself.
        assert "Summary!C8 · 500000" in buttons
        assert "Summary!B8" in buttons
        # Both occurrences are listed and the count is complete: nothing to page.
        assert not any("page suivante" in text for text in buttons)
        query.fill("Total des lignes")
        query.press("Enter")
        total = page.locator(".library-result").first
        total.get_by_role("button", name="Summary!C45", exact=False).first.click()
        members = page.locator(".library-revision-navigation span").first
        playwright.expect(members).to_contain_text("autres cellules")
        text = members.inner_text()
        # Twelve coordinates, then a count; a range is not prefixed twice.
        assert text.count(",") == 11
        assert "Summary!Summary!" not in text
    finally:
        page.close()


def test_range_and_uncached_formula_cards_say_what_they_hold(browser, live_library):
    app, url, upload = live_library
    project = upload("model.xlsx", model_bytes())
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    try:
        page.goto(f"{url}/?project={project['id']}")
        page.locator("#tab-graph").click()
        page.locator("#graph-search").fill("Summary!B6")
        page.locator("#graph-results").get_by_text("Summary · B6").first.click()
        page.wait_for_function(
            """() => document.querySelector('#graph-canvas')._flow?.getNodes()
              .some(node => node.data.original.kind === 'range')"""
        )
        cards = page.locator("#graph-canvas").evaluate(
            """e => Object.fromEntries(e._flow.getNodes().map(node =>
              [node.data.title, [node.data.subtitle, node.data.valueLabel]]))"""
        )
        # A range has no value of its own: "Not available" read as a failure.
        assert cards["Summary!B2:B4"] == ["3 cellules · sans valeur propre", ""]
        # openpyxl saves no cached result for a formula.
        assert cards["Summary · B6"][1] == "Aucune valeur enregistrée"
        page.locator("#ui-language").select_option("en")
        page.wait_for_function(
            """() => document.querySelector('#graph-canvas')._flow.getNodes()
              .some(node => node.data.subtitle === '3 cells · no value of its own')"""
        )
    finally:
        page.close()
