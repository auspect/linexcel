"""Screens of an import cut by its time budget (lot 2 ter, budget step,
reprise after the independent visual review).

A mounted application with a partial index: a formula not analyzed with a
saved value, one without, and the refused calculation the server returns
for them, as the import and lazy.py produce them.
"""

import json
import os
from pathlib import Path

import pytest
from graph_test_helpers import open_workspace_menu

from linexcel.lazy_store import NOT_ANALYZED

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


def _formula(cell, cached):
    return {
        "id": f"Third!{cell}",
        "sheet": "Third",
        "cell": cell,
        "kind": "cell",
        "formula": f"=B{cell[1:]}*3",
        "cachedValue": cached,
        "cached_value": cached,
        "valueSource": "saved_cache",
        "dependencies": [],
        "diagnostics": [NOT_ANALYZED],
    }


META = {
    "status": "partial",
    "sheetDetails": [
        {"name": "First", "status": "complete", "formulas": 1, "diagnostics": []},
        {
            "name": "Third",
            "status": "partial",
            "reason": "time_budget",
            "formulas": 2,
            "unanalyzedFormulas": 2,
            "diagnostics": [],
        },
    ],
    "diagnostics": [],
    "importEstimate": {
        "readSeconds": 3.0,
        "bytesRead": 50_000,
        "bytesTotal": 50_000,
        "budgetSeconds": 20,
        "seconds": 80,
        "beyondMaximum": False,
        "cause": ["reading", "analysis", "regions"],
        "formulasAnalyzed": 17,
        "formulasPredicted": 243,
        "processedShare": 0.07,
        "uncertain": True,
    },
}
GRAPH = {
    "nodes": [
        _formula("A7", None),
        _formula("A8", 24),
        {
            "id": "Third!B7",
            "sheet": "Third",
            "cell": "B7",
            "kind": "input",
            "value": 70,
            "cachedValue": 70,
            "valueSource": "saved_input",
            "dependencies": [],
            "diagnostics": [],
        },
        {
            "id": "First!A1",
            "sheet": "First",
            "cell": "A1",
            "kind": "cell",
            "formula": "=Third!B7*2",
            "cachedValue": 140,
            "valueSource": "saved_cache",
            "dependencies": ["Third!B7"],
            "diagnostics": [],
        },
    ],
    "edges": [{"source": "Third!B7", "target": "First!A1"}],
    "meta": META,
}


def open_app(browser, viewport, language, *, tasks=None, graph=GRAPH):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    project = {"id": "p1", "name": "model.xlsx", "revision": 1}
    created = []

    def handle(route):
        request = route.request
        if request.url.endswith(("app-graph.js", "app-graph.css")):
            route.fulfill(
                body=(APP.parent / request.url.rsplit("/", 1)[-1]).read_text(
                    encoding="utf-8"
                ),
                content_type="text/css"
                if request.url.endswith(".css")
                else "application/javascript",
            )
            return
        path = request.url.split("/mounted/", 1)[-1].split("?", 1)[0]
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/config":
            body = {"limits": {"memoryMb": 1024, "uploadMb": 64}}
        elif path == "api/projects":
            body = {"projects": [project]}
        elif path == "api/projects/p1":
            body = {"project": project, "graph": graph, "tasks": tasks or []}
        elif path == "api/projects/p1/tasks" and request.method == "POST":
            body_in = request.post_data_json
            if body_in.get("operation") != "evaluate":
                route.fulfill(status=409, json={"error": {"message": "fixture"}})
                return
            node_id = body_in.get("nodeId")
            task = {
                "id": str(len(created)),
                "operation": "evaluate",
                "nodeId": node_id,
                "status": "succeeded",
                "result": {
                    "nodeId": node_id,
                    "status": "unsupported",
                    "value": None,
                    "diagnostics": [f"{node_id}: {NOT_ANALYZED}"],
                    "steps": [],
                },
            }
            created.append(task)
            body = {"task": task}
        elif path.startswith("api/tasks/"):
            task = next(t for t in created if t["id"] == path.rsplit("/", 1)[-1])
            body = {"task": task}
        else:
            route.fulfill(status=404, json={"error": {"message": "fixture"}})
            return
        route.fulfill(body=json.dumps(body), content_type="application/json")

    page.route("**/*", handle)
    page.goto(f"http://localhost/mounted/?lang={language}")
    open_workspace_menu(page)
    page.locator("[data-project='p1']").click()
    page.locator("[data-node]").first.wait_for()
    return page, errors


def select(page, node_id):
    if page.locator("#node-back-list").is_visible():
        page.locator("#node-back-list").click()
    page.locator(f"[data-node='{node_id}']").first.click()
    page.locator("#node-detail").wait_for()


VIEWPORTS = [(1440, 900), (390, 844)]


@pytest.mark.parametrize("viewport", VIEWPORTS)
def test_the_refused_result_badge_is_translated(browser, viewport):
    """U1. "Linexcel result · Non pris en charge" in the English interface:
    the status of the inspector summary did not go through translation."""
    page, errors = open_app(browser, viewport, "en")
    try:
        select(page, "Third!A7")
        summary = page.locator(".inspector-summary")
        summary.get_by_text("Linexcel result · Unsupported").wait_for()
        assert "Non pris en charge" not in page.locator("body").inner_text()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize("viewport", VIEWPORTS)
def test_a_formula_not_analyzed_is_explained_in_french(browser, viewport):
    """U2. The refusal of a formula not analyzed showed the server's English
    message in the French interface."""
    page, errors = open_app(browser, viewport, "fr")
    try:
        select(page, "Third!A7")
        detail = page.locator("#node-detail")
        detail.get_by_text(
            "Third!A7: Formule non analysée dans le budget de temps de l’import",
            exact=False,
        ).first.wait_for()
        assert "Formula not analyzed" not in detail.inner_text()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize("language", ["fr", "en"])
def test_a_saved_value_is_said_shown_only_when_the_file_holds_one(browser, language):
    """U3. "Its saved value is shown" beside "Saved cache: Not available":
    the message now says the saved value stays shown when the file holds one,
    and the cell shows it when it does."""
    page, errors = open_app(browser, (1440, 900), language)
    when = (
        "quand le fichier en porte une"
        if language == "fr"
        else "when the file holds one"
    )
    try:
        for node_id, cache in (("Third!A7", None), ("Third!A8", "24")):
            select(page, node_id)
            detail = page.locator("#node-detail")
            detail.get_by_text(when, exact=False).first.wait_for()
            text = detail.inner_text()
            assert "Its saved value is shown" not in text
            if cache:
                assert cache in page.locator(".inspector-summary").inner_text()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize(
    ("viewport", "language", "note"),
    [
        (
            (1440, 900),
            "fr",
            "Liste peut-être incomplète : une formule non importée ou non analysée "
            "n’y figure pas.",
        ),
        (
            (390, 844),
            "en",
            "This list may be incomplete: a formula not imported or not analyzed "
            "is missing from it.",
        ),
    ],
)
def test_the_uses_of_a_node_warn_that_they_may_be_incomplete(
    browser, viewport, language, note
):
    """U4. The warning that listed uses may be incomplete stood only in
    Diagnostics; it now stands beside the list of uses of a node."""
    page, errors = open_app(browser, viewport, language)
    try:
        select(page, "Third!B7")
        # The lists of sources and uses stand in the card's provenance.
        page.locator("#node-detail .inspector-evidence > summary").click()
        page.locator("#node-detail").get_by_text(note).first.wait_for()
        # The graph's notice, below the dependents it draws, says it too.
        # The graph package hides that notice at phone width (app-graph.css,
        # frontend/, outside this step): it is checked as text there.
        page.locator("#tab-graph").click()
        notice = page.locator("#graph-notice")
        if viewport[0] > 900:
            notice.get_by_text(note.split(":")[0], exact=False).wait_for()
        else:
            page.wait_for_function(
                "text => document.querySelector('#graph-notice')"
                ".textContent.includes(text)",
                arg=note.split(":")[0],
            )
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize(
    ("language", "texts"),
    [
        (
            "fr",
            [
                "Lecture des feuilles arrêtée par le budget de 20 s après 3 s, "
                "avec 99 % de leur contenu lu.",
                "Analyse des dépendances arrêtée par le budget de 20 s : 17 formules "
                "analysées sur 243.",
                "Régions de données sautées faute de temps dans le budget de 20 s.",
            ],
        ),
        (
            "en",
            [
                "Reading the sheets stopped by the 20 s budget after 3 s, with 99% "
                "of their content read.",
                "Dependency analysis stopped by the 20 s budget: 17 of 243 formulas "
                "analyzed.",
                "Data regions skipped for lack of time within the 20 s budget.",
            ],
        ),
    ],
)
def test_the_stops_name_the_budget_used(browser, language, texts):
    """U5. "stopped by the budget" did not say which budget: the user could
    not compare it with the budget proposed."""
    page, errors = open_app(browser, (1440, 900), language)
    try:
        page.locator("#tab-diagnostics").click()
        page.locator("#import-estimate").wait_for()
        shown = page.locator("#import-estimate").inner_text()
        for text in texts:
            assert text in shown, (text, shown)
        assert not errors
    finally:
        page.close()


def _timed_out(previous):
    return {
        "id": "t1",
        "projectId": "p1",
        "operation": "import",
        "status": "failed",
        "budget": {"seconds": 20, "memoryMb": 1024},
        "createdAt": 1,
        "error": {
            "kind": "timeout",
            "message": "Import arrêté : le budget de 20 s s’est écoulé à une étape "
            "inconnue. Aucun nouvel index n’a été publié ; l’index précédent reste "
            "consultable. Essayez un budget plus grand, par exemple le double "
            "(40 s) : ce n’est pas une estimation.",
            "suggestedSeconds": 40,
            "previousIndex": previous,
        },
    }


@pytest.mark.parametrize(
    ("language", "heading", "sentence"),
    [
        (
            "fr",
            "Index précédent · Import complet",
            "Les feuilles ci-dessous sont celles de l’index précédent",
        ),
        (
            "en",
            "Previous index · Import complete",
            "The sheets below are those of the previous index",
        ),
    ],
)
def test_after_a_timeout_the_sheets_shown_are_said_to_be_the_previous_index(
    browser, language, heading, sentence
):
    """U6. "Import complete" followed the "Import stopped" notice, and nothing
    said that the sheets listed were those of the previous index."""
    complete = {
        "nodes": GRAPH["nodes"][2:],
        "edges": GRAPH["edges"],
        "meta": {
            "status": "complete",
            "sheetDetails": [
                {
                    "name": "First",
                    "status": "complete",
                    "formulas": 1,
                    "diagnostics": [],
                }
            ],
            "diagnostics": [],
        },
    }
    page, errors = open_app(
        browser, (1440, 900), language, tasks=[_timed_out(True)], graph=complete
    )
    try:
        page.locator("#tab-diagnostics").click()
        page.locator("#import-timeout").wait_for()
        text = page.locator("#diagnostics").inner_text()
        assert heading in text and sentence in text, text
        assert not errors
    finally:
        page.close()
