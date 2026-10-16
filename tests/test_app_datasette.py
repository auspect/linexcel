"""Data navigation uses the protected Datasette surface without stale frames."""

import json
import os
from pathlib import Path
from urllib.parse import urlsplit

import pytest

playwright = pytest.importorskip("playwright.sync_api")
APP = Path(__file__).parents[1] / "src/linexcel/assets/app.html"


@pytest.fixture(scope="module")
def browser():
    with playwright.sync_playwright() as driver:
        instance = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        yield instance
        instance.close()


def open_app(browser, viewport, *, payload=None, delayed=False, language="fr"):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    requests, errors, pending = [], [], []
    page.on("pageerror", lambda error: errors.append(str(error)))
    projects = [{"id": "p", "name": "Public data"}, {"id": "q", "name": "Add-in"}]
    data = payload or {
        "available": True,
        "sourceSha256": "a" * 64,
        "url": "/api/projects/p/datasette/source",
        "cellsUrl": "/api/projects/p/datasette/source/cells",
        "sheets": [
            {
                "sheet": "Country <2023>",
                "table": "sheet_001",
                "url": "/api/projects/p/datasette/source/sheet_001",
                "rowCount": 7,
                "cellCount": 21,
            },
            {
                "sheet": "Inputs",
                "table": "sheet_002",
                "url": "/api/projects/p/datasette/source/sheet_002",
                "rowCount": 2,
                "cellCount": 5,
            },
        ],
    }

    def fulfill(route, body):
        route.fulfill(body=json.dumps(body), content_type="application/json")

    def handle(route):
        path = urlsplit(route.request.url).path
        requests.append(path)
        if path == "/":
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
        elif path.endswith("cytoscape.min.js"):
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
        elif path == "/api/config":
            fulfill(route, {"limits": {"memoryMb": 2048, "uploadMb": 64}})
        elif path == "/api/projects":
            fulfill(route, {"projects": projects})
        elif path in {"/api/projects/p", "/api/projects/q"}:
            project = projects[0 if path.endswith("/p") else 1]
            fulfill(
                route,
                {
                    "project": project,
                    "graph": {"nodes": [], "edges": [], "meta": {}},
                    "tasks": [],
                },
            )
        elif path == "/api/projects/p/data":
            if delayed:
                pending.append(route)
            else:
                fulfill(route, data)
        elif path == "/api/projects/q/data":
            fulfill(route, {"available": False, "reason": "macro_only", "sheets": []})
        elif "/datasette/" in path:
            route.fulfill(
                body="<html><body><h1>Datasette</h1><table><tr><td>497087</td></tr></table></body></html>",
                content_type="text/html",
            )
        else:
            fulfill(route, {})

    page.route("**/*", handle)
    page.goto("http://localhost/?lang=" + language)
    page.locator("[data-project='p']").click()
    page.locator("#tab-data").click()
    return page, requests, errors, pending, data


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
@pytest.mark.parametrize("language", ["fr", "en"])
def test_sheet_and_provenance_navigation_keep_the_frame_when_switching_tabs(
    browser, viewport, language
):
    page, requests, errors, _, _ = open_app(browser, viewport, language=language)
    page.frame_locator("#data-frame").locator("h1").wait_for()
    assert page.locator("#data-sheet option").first.inner_text() == "Country <2023>"
    page.locator("#data-sheet").select_option(label="Inputs")
    page.wait_for_function(
        "document.getElementById('data-frame').src.endsWith('/sheet_002')"
    )
    assert "5" in page.locator("#data-source").inner_text()
    page.locator("#tab-nodes").click()
    page.locator("#tab-data").click()
    assert requests.count("/api/projects/p/data") == 1
    assert page.locator("#data-sheet option:checked").inner_text() == "Inputs"
    page.locator("#data-sheet").select_option("")
    page.wait_for_function(
        "document.getElementById('data-frame').src.endsWith('/cells')"
    )
    assert page.locator("#data-open").get_attribute("href").endswith("/cells")
    assert not page.evaluate("document.documentElement.scrollWidth > innerWidth")
    assert not errors
    page.close()


def test_late_data_reply_cannot_replace_the_next_projects_frame(browser):
    page, _, errors, pending, data = open_app(browser, (1440, 900), delayed=True)
    page.wait_for_function(
        "document.getElementById('data-status').textContent.includes('Chargement')"
    )
    page.locator("[data-project='q']").click()
    page.locator("#data-status").filter(has_text="complément").wait_for()
    assert len(pending) == 1
    pending[0].fulfill(body=json.dumps(data), content_type="application/json")
    page.wait_for_timeout(100)
    assert page.locator("#data-frame").is_hidden()
    assert page.locator("#data-frame").get_attribute("src") is None
    assert "complément" in page.locator("#data-status").inner_text()
    assert not errors
    page.close()


def test_multiple_wide_sheets_keep_separate_filtered_cells_views(browser):
    payload = {
        "available": True,
        "cellsUrl": "/api/projects/p/datasette/source/cells",
        "sheets": [
            {
                "sheet": "Wide One",
                "table": "cells",
                "url": "/api/projects/p/datasette/source/cells?sheet__exact=Wide+One",
                "rowCount": 1,
                "cellCount": 2,
            },
            {
                "sheet": "Wide Two",
                "table": "cells",
                "url": "/api/projects/p/datasette/source/cells?sheet__exact=Wide+Two",
                "rowCount": 3,
                "cellCount": 4,
            },
        ],
    }
    page, _, errors, _, _ = open_app(browser, (390, 844), payload=payload)
    page.locator("#data-sheet option").first.wait_for(state="attached")
    page.locator("#data-sheet").select_option(label="Wide Two")
    assert "sheet__exact=Wide+Two" in page.locator("#data-frame").get_attribute("src")
    assert "Wide Two" in page.locator("#data-source").inner_text()
    assert "4 cellules" in page.locator("#data-source").inner_text()
    assert not errors
    page.close()


@pytest.mark.parametrize(
    "url",
    [
        "https://other.example/data",
        "/api/projects/q/datasette/source",
        "javascript:alert(1)",
    ],
)
def test_only_the_current_projects_datasette_url_can_be_embedded(browser, url):
    page, requests, errors, _, _ = open_app(
        browser, (390, 844), payload={"available": True, "url": url, "sheets": []}
    )
    page.locator("#data-status").filter(has_text="illisible").wait_for()
    assert page.locator("#data-frame").is_hidden()
    assert not any("/datasette/" in path for path in requests)
    assert not errors
    page.close()
