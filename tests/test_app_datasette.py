"""Data navigation uses the protected Datasette surface without stale frames."""

import json
import os
from pathlib import Path
from urllib.parse import urlsplit

import pytest
from graph_test_helpers import open_data_options, open_workspace_menu

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


def open_app(
    browser,
    viewport,
    *,
    payload=None,
    delayed=False,
    language="fr",
    presentation=None,
    project=None,
    data_error=None,
    limits=None,
):
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
        if (
            route.request.method == "POST"
            and route.request.url.endswith("/tasks")
            and route.request.post_data_json.get("operation") == "build_graph"
        ):
            route.fulfill(
                status=409,
                json={"error": {"message": "Overview unavailable in fixture"}},
            )
            return
        path = urlsplit(route.request.url).path
        requests.append(path)
        if path == "/":
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
        elif path.endswith(("app-graph.js", "app-graph.css")):
            route.fulfill(
                body=(APP.parent / route.request.url.rsplit("/", 1)[-1]).read_text(
                    encoding="utf-8"
                ),
                content_type=(
                    "text/css"
                    if route.request.url.endswith(".css")
                    else "application/javascript"
                ),
            )
        elif path == "/api/config":
            fulfill(
                route,
                {"limits": {"memoryMb": 2048, "uploadMb": 64, **(limits or {})}},
            )
        elif path == "/api/projects":
            fulfill(route, {"projects": projects})
        elif path in {"/api/projects/p", "/api/projects/q"}:
            chosen = projects[0 if path.endswith("/p") else 1]
            fulfill(
                route,
                {
                    "project": chosen,
                    "graph": {"nodes": [], "edges": [], "meta": {}},
                    "tasks": [],
                    **(project if path.endswith("/p") and project else {}),
                },
            )
        elif path == "/api/projects/p/data":
            if data_error:
                route.fulfill(
                    status=data_error[0],
                    json={"error": {"message": data_error[1]}},
                )
            elif delayed:
                pending.append(route)
            else:
                fulfill(route, data)
        elif path == "/api/projects/p/data/presentation":
            if callable(presentation):
                presentation(route)
            else:
                fulfill(route, presentation or {})
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
    open_workspace_menu(page)
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
    open_data_options(page)
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
    open_workspace_menu(page)
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
    open_data_options(page)
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


def presentation_payload():
    return {
        "sheet": "Country <2023>",
        "sourceSha256": "a" * 64,
        "window": {"row": 1, "column": 1, "rows": 60, "columns": 20},
        "extent": {"rows": 100, "columns": 4},
        "defaults": {"rowHeight": 20, "columnWidth": 64},
        "styles": {
            "1": {
                "css": {"fontWeight": "bold", "backgroundColor": "#123456"},
                "numberFormat": "#,##0.00",
            }
        },
        "columnDimensions": {"3": {"hidden": True}},
        "rowDimensions": {"3": {"hidden": True}},
        "merges": [{"row": 1, "column": 1, "lastRow": 1, "lastColumn": 2}],
        "cells": [
            {
                "row": 1,
                "column": 1,
                "address": "A1",
                "value": "<script>",
                "display": "<script>",
                "styleId": 1,
                "numberFormat": "General",
            },
            {
                "row": 2,
                "column": 2,
                "address": "B2",
                "value": 1234.5,
                "display": "1,234.50",
                "styleId": 1,
                "numberFormat": "#,##0.00",
                "formatStatus": "formatted",
            },
            {
                "row": 2,
                "column": 4,
                "address": "D2",
                "value": None,
                "display": "",
                "formula": "=SUM(B2:C2)",
                "numberFormat": "0.00",
                "valueStatus": "missing",
            },
            {
                # A text cell whose shared-string part is missing (F17).
                "row": 2,
                "column": 1,
                "address": "A2",
                "value": None,
                "display": "",
                "numberFormat": "General",
                "valueStatus": "unreadable",
            },
            {
                "row": 4,
                "column": 1,
                "address": "A4",
                "value": 42,
                "display": "42",
                "numberFormat": "0.00E+00",
                "formatStatus": "unsupported",
            },
        ],
        "notices": [],
    }


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
@pytest.mark.parametrize("language", ["fr", "en"])
def test_excel_presentation_preserves_source_details_and_datasette(
    browser, viewport, language
):
    data = {
        "available": True,
        "presentationAvailable": True,
        "sourceSha256": "a" * 64,
        "cellsUrl": "/api/projects/p/datasette/source/cells",
        "sheets": [
            {
                "sheet": "Country <2023>",
                "url": "/api/projects/p/datasette/source/sheet_001",
            }
        ],
    }
    page, requests, errors, _, _ = open_app(
        browser,
        viewport,
        payload=data,
        language=language,
        presentation=presentation_payload(),
    )
    try:
        page.locator("#excel-window td[data-cell='B2']").wait_for()
        assert page.locator("#data-frame").is_hidden()
        assert (
            page.locator("#excel-window td[data-cell='A1']").inner_text() == "<script>"
        )
        assert (
            page.locator("#excel-window td[data-cell='A1']").get_attribute("colspan")
            == "2"
        )
        assert (
            page.locator("#excel-window td[data-cell='B2']").inner_text() == "1,234.50"
        )
        assert page.locator("#excel-window td[data-cell='C2']").count() == 0
        assert page.locator("#excel-window td[data-cell='A3']").count() == 0
        assert (
            page.locator("#excel-window td[data-cell='B2']").evaluate(
                "el => getComputedStyle(el).backgroundColor"
            )
            == "rgb(18, 52, 86)"
        )
        page.locator("#excel-window td[data-cell='D2']").click()
        assert "SUM(B2:C2)" in page.locator("#excel-cell-detail").inner_text()
        assert page.locator("#excel-window td[data-cell='D2']").inner_text() == "∅"
        # An unreadable stored value is unknown, never an empty cell (F17).
        unreadable = page.locator("#excel-window td[data-cell='A2']")
        assert unreadable.inner_text() == "∅"
        label = (
            "Value unreadable in the file"
            if language == "en"
            else "Valeur illisible dans le fichier"
        )
        assert label in unreadable.get_attribute("title")
        unreadable.click()
        assert label in page.locator("#excel-cell-detail").inner_text()
        page.locator("#excel-window td[data-cell='B2']").click()
        page.keyboard.press("ArrowLeft")
        assert unreadable.evaluate("el => el === document.activeElement")
        page.keyboard.press("ArrowRight")
        assert "1234.5" in page.locator("#excel-cell-detail").inner_text()
        assert (
            "unsupported" in page.locator("#excel-notices").inner_text()
            if language == "en"
            else "pas pris en charge" in page.locator("#excel-notices").inner_text()
        )
        assert page.locator("#calculation-policy").is_hidden()
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
        page.locator("#data-mode").select_option("datasette")
        page.frame_locator("#data-frame").locator("h1").wait_for()
        assert page.locator("#data-excel").is_hidden()
        assert not errors
        assert not any(path.endswith("/tasks") for path in requests)
    finally:
        page.close()


def test_late_presentation_reply_cannot_cross_projects(browser):
    pending = []
    data = {
        "available": True,
        "presentationAvailable": True,
        "sourceSha256": "a" * 64,
        "sheets": [
            {
                "sheet": "Country <2023>",
                "url": "/api/projects/p/datasette/source/sheet_001",
            }
        ],
    }
    page, _, errors, _, _ = open_app(
        browser, (390, 844), payload=data, presentation=pending.append
    )
    try:
        page.wait_for_function(
            "document.querySelector('#excel-window')"
            ".getAttribute('aria-busy') === 'true'"
        )
        open_workspace_menu(page)
        page.locator("[data-project='q']").click()
        page.locator("#data-status").filter(has_text="complément").wait_for()
        assert len(pending) == 1
        pending[0].fulfill(
            body=json.dumps(presentation_payload()), content_type="application/json"
        )
        page.wait_for_timeout(100)
        assert page.locator("#data-excel").is_hidden()
        assert page.locator("#excel-window td").count() == 0
        assert not errors
    finally:
        page.close()


def test_main_tabs_support_keyboard_navigation_and_scoped_calculation_notice(browser):
    page, _, errors, _, _ = open_app(browser, (390, 844))
    try:
        page.locator("#tab-data").focus()
        page.keyboard.press("End")
        assert page.locator("#tab-diagnostics").get_attribute("aria-selected") == "true"
        assert page.locator("#calculation-policy").is_hidden()
        page.keyboard.press("ArrowRight")
        assert page.locator("#tab-nodes").get_attribute("aria-selected") == "true"
        assert (
            page.locator("#calculation-policy").locator("xpath=..").get_attribute("id")
            == "view-nodes"
        )
        assert page.locator("[data-tab][tabindex='0']").count() == 1
        page.keyboard.press("ArrowLeft")
        assert page.locator("#tab-diagnostics").get_attribute("aria-selected") == "true"
        page.keyboard.press("Home")
        assert page.locator("#tab-nodes").get_attribute("aria-selected") == "true"
        assert not errors
    finally:
        page.close()


BUILDING = {
    "available": False,
    "state": "building",
    "taskId": "t1",
    "presentationAvailable": False,
    "sheets": [],
}


@pytest.mark.parametrize(
    ("viewport", "language", "text"),
    [
        ((1440, 900), "fr", "Préparation de la vue des données en arrière-plan"),
        ((390, 844), "en", "Preparing the data view in the background"),
    ],
)
def test_a_data_view_in_preparation_is_announced_then_opens_by_itself(
    browser, viewport, language, text
):
    """The copy is built by a task (batch 2 bis, F18): the tab says so and
    asks again until the view is available, without a manual refresh."""
    page, requests, errors, pending, data = open_app(
        browser, viewport, delayed=True, language=language
    )
    task = {
        "id": "t1",
        "projectId": "p",
        "operation": "build_data",
        "nodeId": None,
        "status": "running",
        "progress": {"phase": "build_data", "fraction": None},
        "createdAt": 1,
    }
    page.route(
        "**/api/tasks/t1",
        lambda route: route.fulfill(
            body=json.dumps({"task": task}), content_type="application/json"
        ),
    )
    for _ in range(50):
        if pending:
            break
        page.wait_for_timeout(50)
    pending.pop(0).fulfill(body=json.dumps(BUILDING), content_type="application/json")
    page.locator("#data-status").filter(has_text=text).wait_for()
    assert page.locator("#data-controls").is_hidden()
    # The task the server started is shown in Operations at once.
    page.locator("#tab-tasks").click()
    label, phase = (
        ("Vue des données", "Préparation de la vue des données")
        if language == "fr"
        else ("Data view", "Preparing the data view")
    )
    page.locator("#task-list .task").filter(has_text=label).wait_for()
    assert phase in page.locator("#task-list").inner_text()
    assert "build_data" not in page.locator("#task-list").inner_text()
    page.locator("#tab-data").click()
    for _ in range(80):  # the tab polls on its own, about every 2 s
        if pending:
            break
        page.wait_for_timeout(50)
    assert pending, "the tab did not ask again"
    pending.pop(0).fulfill(body=json.dumps(data), content_type="application/json")
    page.frame_locator("#data-frame").locator("h1").wait_for()
    assert requests.count("/api/projects/p/data") == 2
    assert not page.evaluate("document.documentElement.scrollWidth > innerWidth")
    assert not errors
    page.close()


@pytest.mark.parametrize(
    ("language", "prefix", "advice"),
    [
        (
            "fr",
            "Préparation de la vue des données impossible. Quota",
            "Actualiser les données relance la préparation.",
        ),
        (
            "en",
            "The data view could not be prepared. Server",
            "Refresh data starts the preparation again.",
        ),
    ],
)
def test_a_failed_data_view_says_why_and_refresh_prepares_it_again(
    browser, language, prefix, advice
):
    page, _, errors, pending, data = open_app(
        browser, (1440, 900), delayed=True, language=language
    )
    urls = []
    page.on("request", lambda request: urls.append(request.url))
    for _ in range(50):
        if pending:
            break
        page.wait_for_timeout(50)
    failed = {
        "available": False,
        "state": "failed",
        "taskId": "t1",
        "error": "Quota de stockage du serveur atteint (2048 Mio, classeurs, index "
        "et vues compris) : l’opération a été arrêtée et son résultat n’a pas été "
        "conservé. Supprimez des classeurs ou relancez le serveur avec une valeur "
        "--max-storage-mb plus grande.",
        "presentationAvailable": False,
        "sheets": [],
    }
    pending.pop(0).fulfill(body=json.dumps(failed), content_type="application/json")
    status = page.locator("#data-status")
    status.filter(has_text=prefix).wait_for()
    assert advice in status.inner_text()
    reason = (
        "Quota de stockage du serveur atteint"
        if language == "fr"
        else "Server storage quota reached (2048 MiB"
    )
    assert reason in status.inner_text()
    assert "error" in (status.get_attribute("class") or "")
    page.wait_for_timeout(2500)
    assert not pending, "a failed preparation is not retried by itself"
    page.locator("#data-refresh").click()
    for _ in range(50):
        if pending:
            break
        page.wait_for_timeout(50)
    assert pending and any("/data?retry=1" in url for url in urls)
    pending.pop(0).fulfill(body=json.dumps(data), content_type="application/json")
    page.frame_locator("#data-frame").locator("h1").wait_for()
    assert not errors
    page.close()


NO_INDEX = (
    "Aucun index publié : le dernier import n’a pas abouti. "
    "Consultez Opérations pour la raison."
)
STOPPED_IMPORT = {
    "id": "i1",
    "projectId": "p",
    "operation": "import",
    "nodeId": None,
    "status": "failed",
    "error": {
        "kind": "storage_limit",
        "message": "Quota de stockage du serveur atteint",
    },
    "createdAt": 1,
}


@pytest.mark.parametrize(
    ("viewport", "language", "text", "heading", "wrong"),
    [
        (
            (1440, 900),
            "fr",
            "Aucun index publié",
            "Import non abouti",
            ("Import complet", "Réessayez"),
        ),
        (
            (390, 844),
            "en",
            "No index published",
            "Import did not complete",
            ("Import complete", "try again", "Importez", "Aucun"),
        ),
    ],
)
def test_after_a_stopped_import_the_tabs_say_that_no_index_was_published(
    browser, viewport, language, text, heading, wrong
):
    """Batch 2 bis, F22: no tab may suggest a complete import or a retry that
    cannot help when the last import stopped before publishing its index."""
    page, _, errors, _, _ = open_app(
        browser,
        viewport,
        language=language,
        project={"graph": None, "tasks": [STOPPED_IMPORT]},
        data_error=(409, NO_INDEX),
    )
    try:
        status = page.locator("#data-status")
        status.filter(has_text=text).wait_for()
        # The same neutral notice as the other tabs: nothing failed here.
        assert "error" not in (status.get_attribute("class") or "")
        views = {"data": status.inner_text()}
        for name in ("nodes", "graph", "diagnostics"):
            page.locator(f"#tab-{name}").click()
            view = page.locator(f"#view-{name}")
            view.filter(has_text=text).wait_for()
            # Visible text only: a message in a collapsed section is not shown.
            views[name] = view.inner_text()
            assert text in views[name], (name, views[name])
        assert heading in page.locator("#diagnostics h3").first.inner_text()
        for name, shown in views.items():
            for word in wrong:
                assert word not in shown, (name, word, shown)
        assert not page.evaluate("document.documentElement.scrollWidth > innerWidth")
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize(
    ("answer", "english"),
    [
        (
            {
                "available": False,
                "state": "failed",
                "presentationAvailable": False,
                "sheets": [],
                "error": "Quota de stockage du serveur atteint (37 Mio, classeurs, "
                "index et vues compris) : la préparation de la vue des données n’a "
                "pas pu démarrer. Supprimez des classeurs ou relancez le serveur "
                "avec une valeur --max-storage-mb plus grande.",
            },
            "Server storage quota reached (37 MiB, workbooks, indexes and views "
            "included): the data view preparation could not start. Delete "
            "workbooks or restart the server with a larger --max-storage-mb value.",
        ),
        (
            (413, "Stockage plein ; supprimez des classeurs"),
            "Storage full; delete workbooks",
        ),
        (
            (
                413,
                "Quota de stockage du serveur atteint (2 Mio, classeurs, index et "
                "vues compris) : l’opération n’a pas pu démarrer. Supprimez des "
                "classeurs ou relancez le serveur avec une valeur --max-storage-mb "
                "plus grande.",
            ),
            "Server storage quota reached (2 MiB, workbooks, indexes and views "
            "included): the operation could not start. Delete workbooks or restart "
            "the server with a larger --max-storage-mb value.",
        ),
    ],
)
def test_storage_quota_messages_are_translated(browser, answer, english):
    """Batch 2 bis, F22: the quota messages the Data tab can show reach an
    English interface in English."""
    failed = isinstance(answer, dict)
    page, _, errors, _, _ = open_app(
        browser,
        (1440, 900),
        language="en",
        payload=answer if failed else None,
        data_error=None if failed else answer,
    )
    try:
        status = page.locator("#data-status")
        status.filter(has_text=english).wait_for()
        assert "Quota de" not in status.inner_text()
        assert "Stockage" not in status.inner_text()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize(
    ("viewport", "language", "text"),
    [
        (
            (1440, 900),
            "fr",
            "Préparation de la vue des données annulée. "
            "Actualiser les données la relance.",
        ),
        (
            (390, 844),
            "en",
            "Data view preparation cancelled. Refresh data starts it again.",
        ),
    ],
)
def test_a_cancelled_data_view_is_not_shown_as_a_failure(
    browser, viewport, language, text
):
    """Batch 2 bis, F18: cancelling the preparation in Operations is the
    user's choice; the tab says so plainly and Refresh data starts it again."""
    page, _, errors, pending, data = open_app(
        browser, viewport, delayed=True, language=language
    )
    urls = []
    page.on("request", lambda request: urls.append(request.url))
    try:
        for _ in range(50):
            if pending:
                break
            page.wait_for_timeout(50)
        cancelled = {
            "available": False,
            "state": "cancelled",
            "taskId": "t1",
            "presentationAvailable": False,
            "sheets": [],
        }
        pending.pop(0).fulfill(
            body=json.dumps(cancelled), content_type="application/json"
        )
        status = page.locator("#data-status")
        status.filter(has_text=text).wait_for()
        assert "cancelled" not in status.inner_text().replace(text, "")
        assert "error" not in (status.get_attribute("class") or "")
        open_data_options(page)
        page.locator("#data-refresh").click()
        for _ in range(50):
            if pending:
                break
            page.wait_for_timeout(50)
        assert pending and any("/data?retry=1" in url for url in urls)
        pending.pop(0).fulfill(body=json.dumps(data), content_type="application/json")
        page.frame_locator("#data-frame").locator("h1").wait_for()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize(
    ("language", "advice", "overclaim"),
    [
        (
            "fr",
            "les valeurs qui en dépendent restent inconnues (∅), jamais vides",
            "sont signalées",
        ),
        (
            "en",
            "the values that depend on it stay unknown (∅), never empty",
            "flagged",
        ),
    ],
)
def test_the_file_advice_says_only_what_the_interface_shows(
    browser, language, advice, overclaim
):
    """Batch 2 bis, F17: the advice for a missing part must not promise
    flags the sheets do not carry; unreadable values are shown as ∅."""
    meta = {
        "status": "partial",
        "sheetDetails": [
            {"name": "Data", "status": "complete", "formulas": 0, "diagnostics": []}
        ],
        "diagnostics": [
            {
                "part": "xl/sharedStrings.xml",
                "message": "Shared-string part declared by the workbook is "
                "missing: 3 text cell(s) are unreadable, never empty values.",
            }
        ],
    }
    page, _, errors, _, _ = open_app(
        browser,
        (1440, 900),
        language=language,
        project={"graph": {"nodes": [], "edges": [], "meta": meta}},
    )
    try:
        page.locator("#tab-diagnostics").click()
        view = page.locator("#diagnostics")
        view.filter(has_text=advice).wait_for()
        assert advice in view.inner_text()
        assert overclaim not in view.inner_text()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize(
    ("language", "message", "retry"),
    [
        (
            "fr",
            "La partie XML de cette feuille est absente du fichier",
            "Réessayez",
        ),
        ("en", "This sheet's XML part is missing from the file", "Retry"),
    ],
)
def test_a_sheet_the_file_lacks_is_not_offered_a_retry(
    browser, language, message, retry
):
    """Batch 2 bis, F22: showing the area again cannot bring back a part
    the file lacks, so the presentation gives the reason alone."""
    data = {
        "available": True,
        "presentationAvailable": True,
        "sourceSha256": "a" * 64,
        "cellsUrl": "/api/projects/p/datasette/source/cells",
        "sheets": [
            {"sheet": "Lost", "url": "/api/projects/p/datasette/source/sheet_001"}
        ],
    }
    reason = (
        "La partie XML de cette feuille est absente du fichier : elle n’a pas pu "
        "être lue à l’import. Les autres feuilles restent consultables ; voir "
        "Diagnostics."
    )
    page, _, errors, _, _ = open_app(
        browser,
        (1440, 900),
        payload=data,
        language=language,
        presentation=lambda route: route.fulfill(
            status=409, json={"error": {"message": reason}}
        ),
    )
    try:
        open_data_options(page)
        status = page.locator("#excel-status")
        status.filter(has_text=message).wait_for()
        assert retry not in status.inner_text()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize(
    ("language", "label"),
    [("fr", "Construction du graphe"), ("en", "Graph construction")],
)
def test_operations_name_a_running_operation_phase_in_words(browser, language, label):
    """Batch 2 bis: a running operation reports its own name as its phase;
    Operations shows the label, never the internal name."""
    running = {
        "id": "g1",
        "projectId": "p",
        "operation": "build_graph",
        "nodeId": None,
        "status": "running",
        "progress": {"phase": "build_graph", "fraction": None},
        "createdAt": 1,
    }
    page, _, errors, _, _ = open_app(
        browser, (390, 844), language=language, project={"tasks": [running]}
    )
    try:
        page.locator("#tab-tasks").click()
        tasks = page.locator("#task-list")
        tasks.locator(".task").filter(has_text=label).wait_for()
        assert "build_graph" not in tasks.inner_text()
        assert tasks.inner_text().count(label) >= 2, tasks.inner_text()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize(
    ("reason", "english"),
    [
        (
            "Sources modifiées ; importez une nouvelle version",
            "Sources changed; import a new version",
        ),
        (
            "Sources modifiées pendant la lecture ; réessayez avec Afficher la zone.",
            "Sources changed while reading; retry with Show area.",
        ),
        (
            "Cet add-in ne contient pas de feuille de calcul",
            "This add-in contains no worksheet",
        ),
    ],
)
def test_presentation_refusals_reach_english_in_english(browser, reason, english):
    """Batch 2 bis, F22: a 409 reason is shown alone, so it must be
    translated, and the one a retry can help says so itself."""
    data = {
        "available": True,
        "presentationAvailable": True,
        "sourceSha256": "a" * 64,
        "cellsUrl": "/api/projects/p/datasette/source/cells",
        "sheets": [
            {"sheet": "Data", "url": "/api/projects/p/datasette/source/sheet_001"}
        ],
    }
    page, _, errors, _, _ = open_app(
        browser,
        (1440, 900),
        payload=data,
        language="en",
        presentation=lambda route: route.fulfill(
            status=409, json={"error": {"message": reason}}
        ),
    )
    try:
        status = page.locator("#excel-status")
        status.filter(has_text=english).wait_for()
        assert status.inner_text().strip() == english
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize("language", ["fr", "en"])
def test_a_presentation_reason_is_visible_on_a_phone(browser, language):
    """Batch 2 bis, F22: at 390 px the reason must not wait behind the
    collapsed data options; the user would see an empty grid."""
    data = {
        "available": True,
        "presentationAvailable": True,
        "sourceSha256": "a" * 64,
        "cellsUrl": "/api/projects/p/datasette/source/cells",
        "sheets": [
            {"sheet": "Lost", "url": "/api/projects/p/datasette/source/sheet_001"}
        ],
    }
    reason = (
        "La partie XML de cette feuille est absente du fichier : elle n’a pas pu "
        "être lue à l’import. Les autres feuilles restent consultables ; voir "
        "Diagnostics."
    )
    page, _, errors, _, _ = open_app(
        browser,
        (390, 844),
        payload=data,
        language=language,
        presentation=lambda route: route.fulfill(
            status=409, json={"error": {"message": reason}}
        ),
    )
    try:
        text = "absente du fichier" if language == "fr" else "missing from the file"
        status = page.locator("#excel-status")
        status.filter(has_text=text).wait_for(state="attached")
        assert not page.locator("#data-options").evaluate("element => element.open")
        assert status.is_visible()
        assert not page.evaluate("document.documentElement.scrollWidth > innerWidth")
        assert not errors
    finally:
        page.close()


def test_the_import_budget_field_starts_at_the_servers_default(browser):
    """Batch 2 ter: ``linexcel serve --import-seconds`` sets the budget the
    "Import and graph" field proposes; the user can still change it."""
    page, _, errors, _, _ = open_app(
        browser, (1440, 900), limits={"importSeconds": 1500}
    )
    try:
        assert page.locator("#import-seconds").input_value() == "1500"
        assert not errors
    finally:
        page.close()


ESTIMATE_META = {
    "status": "partial",
    "sheetDetails": [
        {"name": "Small", "status": "complete", "formulas": 0, "diagnostics": []},
        {
            "name": "Feuil1",
            "status": "not_imported",
            "reason": "time_budget",
            "formulas": 0,
            "diagnostics": [],
        },
    ],
    "diagnostics": [],
    "importEstimate": {
        "readSeconds": 355.2,
        "bytesRead": 563_000_000,
        "bytesTotal": 796_018_668,
        "budgetSeconds": 600,
        "seconds": 1080,
        "beyondMaximum": False,
        "readShare": 0.6,
    },
}


@pytest.mark.parametrize(
    ("viewport", "language", "rate", "button"),
    [
        (
            (1440, 900),
            "fr",
            "La lecture des feuilles dispose de 60 % du budget de l’import, soit "
            "360 s sur 600 s : elle s’est arrêtée après 355 s, avec 71 % de leur "
            "contenu lu.",
            "Réimporter avec un budget de 1080 s",
        ),
        (
            (390, 844),
            "en",
            "Reading the sheets gets 60% of the import's budget, 360 s of 600 s: "
            "it stopped after 355 s, with 71% of their content read.",
            "Reimport with a budget of 1080 s",
        ),
    ],
)
def test_a_time_stop_proposes_a_reimport_with_an_estimated_budget(
    browser, viewport, language, rate, button
):
    """Batch 2 ter: a sheet not read within the import's time budget gets a
    reimport proposal with the budget estimated from the observed rate."""
    posted = []
    page, _, errors, _, _ = open_app(
        browser,
        viewport,
        language=language,
        project={
            "graph": {"nodes": [], "edges": [], "lazy": True, "meta": ESTIMATE_META}
        },
    )
    page.on(
        "request",
        lambda request: (
            posted.append(request.post_data_json)
            if request.method == "POST" and request.url.endswith("/projects/p/tasks")
            else None
        ),
    )
    try:
        page.locator("#tab-diagnostics").click()
        proposal = page.locator("#import-estimate")
        proposal.wait_for()
        assert rate in proposal.inner_text()
        page.get_by_role("button", name=button).click()
        for _ in range(50):
            if any(body.get("operation") == "import" for body in posted):
                break
            page.wait_for_timeout(50)
        imports = [body for body in posted if body.get("operation") == "import"]
        assert imports and imports[0]["budget"]["seconds"] == 1080
        assert imports[0]["force"] is True
        assert page.locator("#import-seconds").input_value() == "1080"
        assert not page.evaluate("document.documentElement.scrollWidth > innerWidth")
        assert not errors
    finally:
        page.close()
