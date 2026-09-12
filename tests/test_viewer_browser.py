"""Runtime regressions for graph interaction and responsive report rendering.

Run with ``uv sync --extra screenshots``, ``uv run playwright install chromium``
and ``uv run pytest tests/test_viewer_browser.py``. Set
``LINEXCEL_BROWSER_CHANNEL=msedge`` to use an installed Edge on Windows.
"""

import os
from pathlib import Path

import pytest

from linexcel.viewer import render_html

playwright = pytest.importorskip("playwright.sync_api")


@pytest.fixture(scope="module")
def browser():
    with playwright.sync_playwright() as driver:
        browser = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        yield browser
        browser.close()


@pytest.fixture
def report(browser, tmp_path):
    pages = []

    def open_report(graph, *, width=1440, height=900):
        graph = {
            **graph,
            "meta": {
                "stats": {
                    "totalFormulas": 0,
                    "totalNodes": len(graph["nodes"]),
                    "totalEdges": len(graph["edges"]),
                },
                **graph.get("meta", {}),
            },
        }
        page = browser.new_page(viewport={"width": width, "height": height})
        pages.append(page)
        errors = []
        page.on("pageerror", lambda error: errors.append(str(error)))
        path = tmp_path / f"report-{len(pages)}.html"
        path.write_text(render_html(graph), encoding="utf-8")
        page.goto(path.as_uri())
        page.wait_for_function("document.querySelector('#lin-cy')._cyreg?.cy != null")
        page.evaluate("window.cy = document.querySelector('#lin-cy')._cyreg.cy")
        page.wait_for_timeout(100)
        return page, errors

    yield open_report
    for page in pages:
        page.close()


def filtered_graph():
    return {
        "nodes": [
            {"id": "i", "kind": "input", "sheet": "In", "label": "Input"},
            {
                "id": "c",
                "kind": "cell",
                "sheet": "Out",
                "label": "Output",
                "formula": "=In!A1",
                "value": 2,
                "cachedValue": 1,
                "valueSource": "engine",
                "cachedAgreement": "differ",
            },
            {"id": "u", "kind": "input", "sheet": "Other", "label": "Unrelated"},
        ],
        "edges": [{"id": "e", "source": "i", "target": "c", "kind": "ref"}],
    }


def visible_ids(page):
    return page.evaluate("cy.nodes(':visible').map(n => n.id()).sort()")


def capture(page, name):
    directory = os.environ.get("LINEXCEL_VIEWER_ARTIFACTS")
    if directory:
        target = Path(directory)
        target.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(target / f"{name}.png"))


def test_sheet_kind_and_discrepancy_filters_compose(report):
    page, errors = report(filtered_graph())
    page.locator("#lin-rail-kinds button").filter(has_text="Source data").click()
    assert visible_ids(page) == ["c"]
    page.select_option("#lin-sheet-filter", "Out")
    assert visible_ids(page) == ["c"]
    page.select_option("#lin-sheet-filter", "__all__")
    assert visible_ids(page) == ["c"]
    page.locator("#lin-rail-kinds button").filter(has_text="Source data").click()
    page.click("#lin-diffs-only")
    assert visible_ids(page) == ["c"]
    page.select_option("#lin-sheet-filter", "Other")
    assert visible_ids(page) == []
    page.select_option("#lin-sheet-filter", "__all__")
    page.wait_for_function("cy.getElementById('c').visible()", timeout=3000)
    assert visible_ids(page) == ["c"]
    assert not errors


def test_search_does_not_select_hidden_results(report):
    page, errors = report(filtered_graph())
    page.select_option("#lin-sheet-filter", "Out")
    page.fill("#lin-search", "Unrelated")
    page.press("#lin-search", "Enter")
    assert page.locator("#lin-search-status").inner_text() == "No matches"
    assert page.evaluate("cy.nodes(':selected').length") == 0
    assert not errors


def test_group_discrepancy_preserves_the_representative_reading(report):
    group = {
        "id": "g",
        "label": "Group",
        "kind": "group",
        "formula": "=B1*2",
        "r1c1": "=RC[1]*2",
        "count": 2,
        "bbox": "A1:A2",
        "value": 4,
        "cachedValue": 4,
        "valueSource": "engine",
        "cachedAgreement": "same",
        "groupCachedAgreement": "differ",
        "samples": [
            {
                "addr": "A1",
                "value": 4,
                "cachedValue": 4,
                "source": "engine",
                "cachedAgreement": "same",
            },
            {
                "addr": "A2",
                "value": 6,
                "cachedValue": 4,
                "source": "engine",
                "cachedAgreement": "differ",
            },
        ],
    }
    page, errors = report({"nodes": [group], "edges": []})
    page.click("#lin-diffs-only")
    assert visible_ids(page) == ["g"]
    page.fill("#lin-search", "Group")
    page.press("#lin-search", "Enter")
    assert "differs from the file" in page.locator(".lin-verdict").inner_text()
    assert page.locator(".lin-vtable tbody tr").count() == 2
    assert page.locator(".lin-vtable tr.is-diff th").inner_text() == "A2"
    assert page.evaluate("cy.nodes(':selected').length") == 1
    assert not errors


def test_dense_search_updates_detail_once_per_frame(report):
    graph = {
        "nodes": [
            {"id": str(i), "kind": "cell", "label": f"Match {i}"} for i in range(400)
        ],
        "edges": [],
    }
    page, errors = report(graph)
    aspect = page.evaluate("""() => {
        const box = cy.nodes().boundingBox();
        return box.w / box.h;
    }""")
    assert 0.2 < aspect < 5, "Isolated cards must fill an area, not one thin rank"
    # Count graph-wide scans during the synchronous selection transaction.
    # The ceiling is independent of node count; timing thresholds vary by CI.
    calls = page.evaluate("""() => {
        const original = cy.nodes;
        let calls = 0;
        cy.nodes = function (...args) { calls++; return original.apply(this, args); };
        const input = document.querySelector('#lin-search');
        input.value = 'Match';
        input.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'Enter', bubbles: true
        }));
        cy.nodes = original;
        return calls;
    }""")
    assert calls < 20
    assert page.evaluate("cy.nodes(':selected').length") == 400
    page.wait_for_timeout(600)
    capture(page, "dense-search")
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
@pytest.mark.parametrize("dark", [False, True])
def test_report_controls_fit_viewport(report, width, height, dark):
    page, errors = report(filtered_graph(), width=width, height=height)
    if dark:
        page.click("#lin-theme")
    page.fill("#lin-search", "Output")
    page.press("#lin-search", "Enter")
    page.wait_for_timeout(100)
    assert page.evaluate("cy.getElementById('c').style('opacity')") == "1"
    for selector in ("#lin-search", "#lin-theme", ".lin-close"):
        box = page.locator(selector).bounding_box()
        assert box is not None
        assert box["x"] >= 0 and box["x"] + box["width"] <= width
    page.wait_for_timeout(600)
    capture(page, f"report-{width}-{'dark' if dark else 'light'}")
    assert not errors
