"""Direct search and explicit return preserve the workbook exploration context."""

import os
from pathlib import Path

import pytest
from test_library_search_transitions import book_bytes, scope_ready
from test_library_search_transitions import browser as browser
from test_library_search_transitions import live_library as live_library

playwright = pytest.importorskip("playwright.sync_api")


@pytest.mark.parametrize("width", [1440, 390])
def test_direct_search_filters_and_return_preserve_workbook(
    browser, live_library, width
):
    app, url, upload = live_library
    project = upload("alpha.xlsx", book_bytes())
    page = browser.new_page(
        viewport={"width": width, "height": 844 if width == 390 else 900}
    )
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    try:
        page.goto(f"{url}/?project={project['id']}")
        page.locator("#tab-sheets").click()
        playwright.expect(page.locator("#tab-sheets")).to_have_attribute(
            "aria-selected", "true"
        )
        page.locator(".library-search-trigger").click()
        scope_ready(page, url)
        query = page.locator("#library-search-query")
        playwright.expect(query).to_be_focused()
        advanced = page.locator(".library-advanced")
        assert advanced.get_attribute("open") is None
        assert page.locator(".library-kind-filter").is_hidden()
        assert not page.get_by_role("button", name="Poser une question à l’IA").count()
        back = page.get_by_role("button", name="Retour au classeur", exact=True)
        box = back.bounding_box()
        assert box and box["y"] >= 48 and box["y"] + box["height"] < 400
        query.fill("ALPHA")
        query.press("Enter")
        page.locator(".library-result").first.wait_for()
        assert advanced.get_attribute("open") is None
        assert not any(task["operation"] == "ask" for task in app.store.tasks.values())
        capture_dir = os.environ.get("M3_CAPTURE_DIR")
        if capture_dir:
            target = Path(capture_dir)
            target.mkdir(parents=True, exist_ok=True)
            page.locator(".workspace > main").evaluate("e => e.scrollTop=0")
            page.screenshot(path=str(target / f"direct-search-{width}.png"))
        advanced.locator("summary").first.click()
        page.locator(".library-scope button").filter(has_text="Classeurs").click()
        page.locator(".library-scope select").select_option(project["id"])
        advanced.locator("summary").first.click()
        playwright.expect(advanced.locator("summary").first).to_contain_text(
            "alpha.xlsx"
        )
        page.locator(".workspace > main").evaluate("e => e.scrollTop=0")
        back.click()
        playwright.expect(page.locator("#view-sheets")).to_be_visible()
        playwright.expect(page.locator("#tab-sheets")).to_have_attribute(
            "aria-selected", "true"
        )
        playwright.expect(page.locator(".library-search-trigger")).to_be_focused()
        page.locator(".library-search-trigger").click()
        playwright.expect(query).to_have_value("ALPHA")
        query.press("Escape")
        playwright.expect(page.locator("#view-sheets")).to_be_visible()
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
        assert not errors
    finally:
        page.close()
