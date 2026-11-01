"""Source evidence must stay distinct and escaped on both serve surfaces."""

import json
import os
from pathlib import Path

import pytest
from graph_test_helpers import open_workspace_menu

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


def open_graph(browser, viewport, node, *, macro_only=False):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    graph = {
        "nodes": [node],
        "edges": [],
        "formulaPatterns": [],
        "meta": {"sheets": [] if macro_only else ["Lookups"], "macroOnly": macro_only},
    }

    def route(request):
        path = request.request.url.split("http://localhost/", 1)[-1]
        if path == "":
            request.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path.endswith("cytoscape.min.js"):
            request.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
            return
        data = {}
        if path == "api/config":
            data = {"limits": {"memoryMb": 2048, "uploadMb": 64}}
        elif path == "api/projects":
            data = {"projects": [{"id": "p", "name": "Source review"}]}
        elif path == "api/projects/p":
            data = {
                "project": {"id": "p", "name": "Source review"},
                "graph": graph,
                "tasks": [],
            }
        elif path == "api/projects/p/tasks":
            body = request.request.post_data_json
            assert not macro_only  # Static VBA access must never execute code.
            assert body["operation"] == "evaluate"
            assert body["nodeId"] == node["id"]
            assert body["force"] is True
            data = {
                "task": {
                    "id": "targeted-evaluation",
                    "operation": "evaluate",
                    "nodeId": node["id"],
                    "status": "succeeded",
                    "result": {"status": "unsupported"},
                }
            }
        request.fulfill(body=json.dumps(data), content_type="application/json")

    page.route("**/*", route)
    page.goto("http://localhost/")
    open_workspace_menu(page)
    page.locator("[data-project='p']").click()
    page.locator("#node-list [data-node]").click()
    page.locator("#node-detail h2").wait_for()
    return page, errors


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_original_and_canonical_formulas_are_distinct_in_detail_and_reader(
    browser, viewport
):
    original = "=" + "(" * 40 + "1+2" + ")" * 40
    node = {
        "id": "Lookups!B22",
        "kind": "cell",
        "sheet": "Lookups",
        "cell": "B22",
        "formula": "=1 + 2",
        "sourceFormula": original,
        "canonicalFormula": "=1 + 2",
        "formulaProvenance": {
            "xmlText": original[1:],
            "status": "stored_xml",
            "workbookSha256": "a" * 64,
            "worksheetPart": "xl/worksheets/sheet1.xml",
        },
    }
    page, errors = open_graph(browser, viewport, node)
    assert page.locator("#node-detail [data-formula-source]").inner_text() == original
    assert (
        page.locator("#node-detail [data-formula-canonical]").inner_text() == "=1 + 2"
    )
    page.locator("#detail-graph").click()
    assert (
        page.locator("#graph-preview .graph-reader-heading .formula").inner_text()
        == original
    )
    page.locator("#graph-panel-calculation > .formula-proof > summary").click()
    assert (
        page.locator("#graph-panel-calculation [data-formula-source]").inner_text()
        == original
    )
    assert (
        page.locator("#graph-panel-calculation [data-formula-canonical]").inner_text()
        == "=1 + 2"
    )
    assert not page.evaluate("document.documentElement.scrollWidth>innerWidth")
    assert not errors
    page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_shared_slave_does_not_present_translated_formula_as_xml_source(
    browser, viewport
):
    node = {
        "id": "Lookups!B2",
        "kind": "cell",
        "sheet": "Lookups",
        "cell": "B2",
        "formula": "=A2+1",
        "canonicalFormula": "=A2+1",
        "sourceFormula": None,
        "formulaProvenance": {
            "status": "shared_translation",
            "xmlText": "",
            "shared": {"role": "slave", "masterCell": "B1"},
        },
    }
    page, errors = open_graph(browser, viewport, node)
    assert page.locator("#node-detail [data-formula-source]").count() == 0
    assert page.locator("#node-detail [data-formula-canonical]").inner_text() == "=A2+1"
    assert "Aucun texte de formule" in page.locator("#node-detail").inner_text()
    page.locator("#detail-graph").click()
    assert (
        page.locator("#graph-preview .graph-reader-heading .formula").inner_text()
        == "=A2+1"
    )
    page.locator("#graph-panel-calculation > .formula-proof > summary").click()
    assert page.locator("#graph-panel-calculation [data-formula-source]").count() == 0
    assert (
        "Aucun texte de formule"
        in page.locator("#graph-panel-calculation").inner_text()
    )
    assert not errors
    page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_macro_source_is_escaped_bounded_and_has_no_recalculation_or_captures(
    browser, viewport
):
    code = (
        'Public Sub Run()\n  x = "<img src=x onerror=window.pwned=1>"\nEnd Sub\n'
        + "'a" * 4000
    )
    node = {
        "id": "vp:Addin.Run",
        "kind": "vba",
        "name": "Addin.Run",
        "module": "Addin",
        "procKind": "Sub",
        "lines": [1, 4],
        "code": code,
        "vbaProvenance": {
            "codeTruncated": True,
            "fileSha256": "b" * 64,
            "totalCodeCharacters": len(code),
        },
    }
    page, errors = open_graph(browser, viewport, node, macro_only=True)
    assert page.locator("#node-detail [data-vba-code]").inner_text() == code[:6000]
    assert page.locator("#node-detail img").count() == 0
    assert page.locator("#recalculate").count() == 0
    assert "Aucune macro exécutée" in page.locator("#node-detail").inner_text()
    page.locator("#detail-graph").click()
    assert (
        page.locator("#graph-panel-calculation [data-vba-code]").inner_text()
        == code[:6000]
    )
    assert not page.locator("#graph-recalculate").is_visible()
    assert page.locator("#graph-preview img").count() == 0
    page.locator("#tab-sheets").click()
    assert "sans analyse des cellules" in page.locator("#sheet-list").inner_text()
    page.locator("#tab-captures").click()
    assert page.locator("#capture").is_disabled()
    assert "aucune capture" in page.locator("#capture-results").inner_text()
    assert not page.evaluate("document.documentElement.scrollWidth>innerWidth")
    assert not errors
    page.close()
