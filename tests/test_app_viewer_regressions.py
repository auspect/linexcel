"""Browser regressions for expanded serve controls and workbook navigation."""

import json
import os
import time
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

import pytest
from graph_test_helpers import open_graph_options

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


@pytest.fixture(params=[(1440, 900), (390, 844)])
def app(browser, request):
    width, height, *response_delays = request.param
    project_response_delay_ms = response_delays[0] if response_delays else 0
    page = browser.new_page(viewport={"width": width, "height": height})
    page.add_init_script("localStorage.setItem('linexcel-manual-recalculation','true')")
    if project_response_delay_ms:
        page.add_init_script(
            """(() => {
                const originalFetch = window.fetch;
                window.fetch = async (...args) => {
                    const response = await originalFetch(...args);
                    const path = new URL(args[0], location.href).pathname;
                    if (path === '/api/projects/p') {
                        await new Promise(resolve => setTimeout(resolve, DELAY_MS));
                    }
                    return response;
                };
            })();""".replace("DELAY_MS", str(project_response_delay_ms))
        )
    project = {"id": "p", "name": "Navigation", "revision": 1}
    nodes = [
        {"id": "First!A1", "sheet": "First", "cell": "A1", "kind": "input", "value": 1},
        {
            "id": "Second!A1",
            "sheet": "Second",
            "cell": "A1",
            "kind": "input",
            "value": 2,
        },
        {
            "id": "Second!A2",
            "sheet": "Second",
            "cell": "A2",
            "kind": "cell",
            "formula": "=A1*2",
        },
        {"id": "Slow!A1", "sheet": "Slow", "cell": "A1", "kind": "input", "value": 3},
    ]
    graph = {
        "lazy": True,
        "nodes": nodes[:1],
        "edges": [],
        "formulaPatterns": [],
        "meta": {"sheets": ["First", "Second", "Slow", "Empty"], "nodeCount": 3},
        "pagination": {"offset": 0, "limit": 200, "total": len(nodes)},
    }
    requests, errors, delayed = [], [], []
    page.on("pageerror", lambda error: errors.append(str(error)))

    def fulfill(route, body):
        route.fulfill(body=json.dumps(body), content_type="application/json")

    def handle(route):
        path = urlsplit(route.request.url).path
        params = parse_qs(urlsplit(route.request.url).query, keep_blank_values=True)
        payload = (
            route.request.post_data_json if route.request.method == "POST" else None
        )
        requests.append((path, payload if isinstance(payload, dict) else params))
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
            fulfill(route, {"projects": [project]})
        elif path == "/api/projects/p":
            fulfill(route, {"project": project, "graph": graph, "tasks": []})
        elif path == "/api/projects/p/nodes":
            sheet = params.get("sheet", [""])[0]
            kind = params.get("kind", [""])[0]
            result = [
                n
                for n in nodes
                if (not sheet or n["sheet"] == sheet)
                and (not kind or n["kind"] == kind)
            ]
            data = {
                "nodes": result,
                "pagination": {"offset": 0, "limit": 200, "total": len(result)},
            }
            if sheet == "Slow":
                delayed.append((route, data))
            else:
                fulfill(route, data)
        elif path == "/api/projects/p/neighborhood":
            node_id = params["nodeId"][0]
            ids = {node_id, "Second!A1"} if node_id == "Second!A2" else {node_id}
            edges = (
                [{"source": "Second!A1", "target": "Second!A2"}]
                if node_id == "Second!A2"
                else []
            )
            fulfill(
                route, {"nodes": [n for n in nodes if n["id"] in ids], "edges": edges}
            )
        elif path == "/api/projects/p/tasks":
            if payload["operation"] == "evaluate":
                assert payload["force"] is True
                fulfill(
                    route,
                    {
                        "task": {
                            "id": "evaluate-" + str(len(requests)),
                            "nodeId": payload["nodeId"],
                            "operation": "evaluate",
                            "status": "succeeded",
                            "result": {
                                "status": "unsupported",
                                "diagnostics": ["Fixture engine unavailable"],
                            },
                        }
                    },
                )
            else:
                assert payload["operation"] == "build_graph"
                fulfill(
                    route,
                    {
                        "task": {
                            "id": "build",
                            "operation": "build_graph",
                            "status": "succeeded",
                        }
                    },
                )
        elif path == "/api/projects/p/graph-view":
            fulfill(
                route,
                {
                    "viewId": 0,
                    "breadcrumbs": [{"viewId": 0, "label": "Workbook"}],
                    "scopeNodeCount": 3,
                    "meta": {"generation": 1, "nodeCount": 3, "edgeCount": 1},
                    "nodes": [
                        {
                            "id": "left",
                            "kind": "group",
                            "viewId": 1,
                            "label": "First",
                            "nodeCount": 1,
                            "internalEdges": 0,
                            "position": {"x": 0, "y": 0},
                        },
                        {
                            "id": "right",
                            "kind": "group",
                            "viewId": 2,
                            "label": "Second",
                            "nodeCount": 2,
                            "internalEdges": 0,
                            "position": {"x": 10000, "y": 0},
                        },
                    ],
                    "edges": [{"source": "left", "target": "right", "weight": 1}],
                },
            )
        else:
            fulfill(route, {})

    page.route("**/*", handle)
    page.goto("http://localhost/")
    page.locator("[data-project='p']").click()
    # A click does not await openProject's fetch or its localStorage update.
    # Reload tests must start from the rendered, persisted initial project.
    playwright.expect(page.locator("#node-list [data-node='First!A1']")).to_be_visible()
    yield page, requests, errors, delayed, fulfill
    for route, data in delayed:
        try:
            fulfill(route, data)
        except playwright.Error:
            pass
    page.close()


def test_expanded_sidebar_fields_stay_inside_sidebar(app):
    page, _, errors, _, _ = app
    for label in ["Limites par opération", "Documentation IA à la demande"]:
        page.get_by_text(label, exact=True).click()
    assert page.evaluate("""() => {
        const sidebar=document.querySelector('aside').getBoundingClientRect();
        const fields='.budget input,.ai-settings input,.ai-settings select';
        return [...document.querySelectorAll(fields)].every(e=>{
            const b=e.getBoundingClientRect();
            return b.left>=sidebar.left+12&&b.right<=sidebar.right-12;
        });
    }""")
    assert not errors


def test_sheet_navigation_handles_blank_empty_and_stale_response(app):
    page, _, errors, delayed, fulfill = app
    assert page.locator("#manual-recalculation").count() == 0
    page.locator("#tab-graph").click()
    page.locator("#graph-sheet").select_option("Second")
    playwright.expect(
        page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("Second")
    assert "A2" in page.locator("#graph-preview .graph-reader-heading h3").inner_text()
    assert page.locator("#graph-canvas").evaluate("el=>el._cy.edges().length") == 1
    page.locator("#graph-sheet").select_option("")
    assert page.locator("#graph-sheet").input_value() == ""
    assert (
        "Second" in page.locator("#graph-preview .graph-reader-heading h3").inner_text()
    )
    page.locator("#graph-sheet").select_option("Empty")
    playwright.expect(page.locator("#graph-sheet-status")).to_contain_text(
        "Aucune cellule indexée"
    )
    assert "Empty" in page.locator("#graph-sheet-status").inner_text()
    assert (
        "Second" in page.locator("#graph-preview .graph-reader-heading h3").inner_text()
    )
    page.locator("#tab-nodes").click()
    assert page.locator("#graph-sheet-status").is_hidden()
    page.locator("#tab-graph").click()
    page.locator("#graph-sheet").select_option("Slow")
    deadline = time.monotonic() + 5
    while not delayed and time.monotonic() < deadline:
        page.wait_for_timeout(10)
    assert delayed
    page.locator("#graph-sheet").select_option("First")
    playwright.expect(
        page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("First")
    route, data = delayed.pop()
    fulfill(route, data)
    page.wait_for_timeout(100)
    assert page.locator("#graph-sheet").input_value() == "First"
    assert (
        "First" in page.locator("#graph-preview .graph-reader-heading h3").inner_text()
    )
    assert not errors


def test_full_workbook_overview_fits_and_reader_can_be_hidden(app):
    page, _, errors, _, _ = app
    page.locator("#tab-graph").click()
    open_graph_options(page)
    page.locator("#graph-build").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._cy?.nodes()"
        ".map(n=>n.id()).sort().join(',')==='left,right'"
    )
    assert page.evaluate("""() => {
        const cy=document.querySelector('#graph-canvas')._cy;
        return cy.nodes().every(n=>{const b=n.renderedBoundingBox();
            return b.x1>=0&&b.x2<=cy.width()&&b.y1>=0&&b.y2<=cy.height();});
    }""")
    canvas = page.locator("#graph-canvas")
    before = canvas.bounding_box()
    page.locator("#graph-reader-toggle").click()
    assert page.locator("#graph-preview").is_hidden()
    assert (
        page.locator("#graph-reader-toggle").get_attribute("aria-expanded") == "false"
    )
    after = canvas.bounding_box()
    assert after["width"] > before["width"] or after["height"] > before["height"]
    page.locator("#graph-fit").click()
    assert canvas.evaluate("el => el._cy.zoom() < .05") == (
        page.viewport_size["width"] == 390
    )
    page.locator("#graph-reader-toggle").click()
    assert page.locator("#graph-preview").is_visible()
    assert not errors


@pytest.mark.parametrize("has_input", [True, False])
def test_data_only_sheet_prefers_input_before_local_name(app, has_input):
    page, _, errors, _, fulfill = app
    name = {
        "id": "name:First:_xlnm._FilterDatabase",
        "kind": "name",
        "sheet": "First",
        "name": "_xlnm._FilterDatabase",
        "formula": "=First!A1:A2",
    }
    source = {
        "id": "First!A1",
        "kind": "input",
        "sheet": "First",
        "cell": "A1",
        "value": "Population estimates",
    }
    requested_kinds = []

    def first_sheet(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params.get("sheet") != ["First"]:
            route.fallback()
            return
        kind = params.get("kind", [""])[0]
        requested_kinds.append(kind)
        candidates = (
            []
            if kind == "cell"
            else [source]
            if kind == "input" and has_input
            else []
            if kind == "input"
            else [name, source]
            if has_input
            else [name]
        )
        fulfill(route, {"nodes": candidates, "pagination": {"total": len(candidates)}})

    def named_neighborhood(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params.get("nodeId") == [name["id"]]:
            fulfill(route, {"nodes": [name], "edges": []})
        else:
            route.fallback()

    page.route("**/api/projects/p/nodes?*", first_sheet)
    page.route("**/api/projects/p/neighborhood?*", named_neighborhood)
    page.locator("#tab-graph").click()
    page.locator("#graph-sheet").select_option("First")
    playwright.expect(
        page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("First · " + ("A1" if has_input else "_xlnm._FilterDatabase"))
    playwright.expect(page.locator("#node-detail h2")).to_contain_text(
        "First · " + ("A1" if has_input else "_xlnm._FilterDatabase")
    )
    assert requested_kinds == (
        ["cell", "input"] if has_input else ["cell", "input", ""]
    )
    assert page.locator("#graph-sheet").input_value() == "First"
    assert not errors


@pytest.mark.parametrize("next_sheet", ["Empty", ""])
def test_sheet_change_cancels_pending_neighborhood_after_optimistic_selection(
    app, next_sheet
):
    page, _, errors, _, fulfill = app
    page.locator("#tab-graph").click()
    page.locator("#graph-sheet").select_option("Second")
    heading = page.locator("#graph-preview .graph-reader-heading h3")
    playwright.expect(heading).to_contain_text("Second")
    # First/depth1 was loaded on the initial Graph access. This race uses an
    # uncached depth2 neighborhood, through the real depth control.
    page.locator("#graph-depth").fill("2")
    page.locator("#graph-depth").press("Tab")
    page.wait_for_function(
        "document.querySelector('#graph-notice').textContent.includes('profondeur 2')"
    )
    pending = []

    def delay_first(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params["nodeId"] == ["First!A1"]:
            pending.append(route)
        else:
            route.fallback()

    page.route("**/api/projects/p/neighborhood?*", delay_first)
    page.locator("#graph-sheet").select_option("First")
    page.wait_for_timeout(100)
    assert pending
    playwright.expect(heading).to_contain_text("First")
    page.locator("#graph-sheet").select_option(next_sheet)
    if next_sheet:
        playwright.expect(page.locator("#graph-sheet-status")).to_contain_text("Empty")
    fulfill(
        pending.pop(),
        {
            "nodes": [
                {
                    "id": "First!A1",
                    "sheet": "First",
                    "cell": "A1",
                    "kind": "input",
                    "value": 1,
                }
            ],
            "edges": [],
        },
    )
    page.wait_for_timeout(100)
    assert page.locator("#graph-sheet").input_value() == next_sheet
    assert "First" in heading.inner_text()
    if next_sheet:
        assert "Empty" in page.locator("#graph-sheet-status").inner_text()
    open_graph_options(page)
    page.locator("#graph-selected").click()
    assert "First" in page.locator("#node-detail h2").inner_text()
    assert not errors


@pytest.mark.parametrize("next_sheet", ["Empty", ""])
def test_sheet_change_cancels_pending_hierarchy_after_optimistic_selection(
    app, next_sheet
):
    page, _, errors, _, fulfill = app
    page.locator("#tab-graph").click()
    page.locator("#graph-sheet").select_option("Second")
    heading = page.locator("#graph-preview .graph-reader-heading h3")
    playwright.expect(heading).to_contain_text("Second")
    open_graph_options(page)
    page.locator("#graph-build").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._cy?.nodes()"
        ".map(n=>n.id()).sort().join(',')==='left,right'"
    )
    pending = []

    def delay_first(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params.get("nodeId") == ["First!A1"]:
            pending.append(route)
        else:
            route.fallback()

    page.route("**/api/projects/p/graph-view?*", delay_first)
    page.locator("#graph-sheet").select_option("First")
    page.wait_for_timeout(100)
    assert pending
    playwright.expect(heading).to_contain_text("First")
    page.locator("#graph-sheet").select_option(next_sheet)
    if next_sheet:
        playwright.expect(page.locator("#graph-sheet-status")).to_contain_text("Empty")
    try:
        fulfill(
            pending.pop(),
            {
                "viewId": -1,
                "breadcrumbs": [{"viewId": -1, "label": "First"}],
                "scopeNodeCount": 1,
                "meta": {"generation": 1, "nodeCount": 3, "edgeCount": 0},
                "nodes": [
                    {
                        "id": "first-leaf",
                        "nodeId": "First!A1",
                        "kind": "input",
                        "label": "First!A1",
                        "internalEdges": 0,
                        "position": {"x": 0, "y": 0},
                    }
                ],
                "edges": [],
            },
        )
    except playwright.Error:
        pass  # Aborting the obsolete fetch is an equally valid cancellation.
    page.wait_for_timeout(100)
    assert page.locator("#graph-sheet").input_value() == next_sheet
    assert "First" in heading.inner_text()
    assert page.locator("#graph-canvas").evaluate(
        "e=>e._cy.nodes().map(n=>n.id()).sort()"
    ) == ["left", "right"]
    assert not errors


def test_graph_tab_does_not_cancel_pending_node_selection(app):
    page, _, errors, _, fulfill = app
    pending = []

    def delay(route):
        pending.append(route)

    page.route("**/api/projects/p/neighborhood?*", delay)
    page.locator("#node-list [data-node='First!A1']").click()
    page.wait_for_timeout(100)
    assert len(pending) == 1
    page.locator("#tab-graph").click()
    page.wait_for_timeout(100)
    assert len(pending) == 1  # No refresh of the old center supersedes selection.
    fulfill(
        pending.pop(),
        {
            "nodes": [
                {
                    "id": "First!A1",
                    "sheet": "First",
                    "cell": "A1",
                    "kind": "input",
                    "value": 1,
                }
            ],
            "edges": [],
        },
    )
    playwright.expect(
        page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("First")
    page.locator("#tab-nodes").click()
    assert "First" in page.locator("#node-detail h2").inner_text()
    assert page.locator("#node-list .node.selected").count() == 1
    assert not errors


def test_depth_change_restarts_pending_target_and_ignores_old_depth(app):
    page, _, errors, _, fulfill = app
    page.locator("#tab-graph").click()
    page.wait_for_function("document.querySelector('#graph-canvas')._cy != null")
    pending = []

    def delay_second(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params["nodeId"] == ["Second!A2"]:
            pending.append((route, params))
        else:
            route.fallback()

    page.route("**/api/projects/p/neighborhood?*", delay_second)
    page.locator("#graph-sheet").select_option("Second")
    page.wait_for_timeout(100)
    assert len(pending) == 1
    page.locator("#graph-depth").fill("2")
    page.locator("#graph-depth").press("Tab")
    page.wait_for_timeout(100)
    assert len(pending) == 2 and pending[1][1]["depth"] == ["2"]
    route, _ = pending.pop()
    fulfill(
        route,
        {
            "nodes": [
                {
                    "id": "Second!A2",
                    "sheet": "Second",
                    "cell": "A2",
                    "kind": "cell",
                    "formula": "=A1*2",
                },
                {
                    "id": "Second!A1",
                    "sheet": "Second",
                    "cell": "A1",
                    "kind": "input",
                    "value": 2,
                },
            ],
            "edges": [{"source": "Second!A1", "target": "Second!A2"}],
        },
    )
    playwright.expect(
        page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("Second")
    route, _ = pending.pop()
    fulfill(route, {"nodes": [], "edges": []})
    page.wait_for_timeout(100)
    assert page.locator("#graph-depth").input_value() == "2"
    assert page.locator("#graph-canvas").evaluate("e=>e._cy.edges().length") == 1
    assert not errors


@pytest.mark.parametrize(
    "app",
    [(1440, 900, 0), (390, 844, 0), (1440, 900, 200), (390, 844, 200)],
    indirect=True,
    ids=["desktop-immediate", "mobile-immediate", "desktop-delayed", "mobile-delayed"],
)
def test_project_intent_cancels_old_selection_before_project_response(app):
    page, requests, errors, _, fulfill = app
    next_project = {"id": "next", "name": "Next workbook", "revision": 1}
    page.route(
        "**/api/projects",
        lambda route: fulfill(
            route,
            {
                "projects": [
                    {"id": "p", "name": "Navigation", "revision": 1},
                    next_project,
                ]
            },
        ),
    )
    page.reload()
    page.locator("#node-list [data-node='First!A1']").wait_for()
    page.locator("#search").fill("Second!A2")
    page.locator("#node-list [data-node='Second!A2']").wait_for()
    old_pending, project_pending = [], []
    page.route(
        "**/api/projects/p/neighborhood?*", lambda route: old_pending.append(route)
    )
    page.route("**/api/projects/next", lambda route: project_pending.append(route))
    page.locator("#node-list [data-node='Second!A2']").click()
    page.wait_for_timeout(100)
    page.locator("[data-project='next']").click()
    page.wait_for_timeout(100)
    assert old_pending and project_pending
    fulfill(
        old_pending.pop(),
        {
            "nodes": [
                {
                    "id": "Second!A2",
                    "sheet": "Second",
                    "cell": "A2",
                    "kind": "cell",
                    "formula": "=A1*2",
                },
            ],
            "edges": [],
        },
    )
    page.wait_for_timeout(100)
    assert "Second" in page.locator("#node-detail h2").inner_text()
    posted = [payload for path, payload in requests if path.endswith("/tasks")]
    assert len(posted) == 1
    assert posted[0]["operation"] == "evaluate"
    assert posted[0]["nodeId"] == "Second!A2" and posted[0]["force"] is True
    fulfill(
        project_pending.pop(),
        {
            "project": next_project,
            "graph": {
                "lazy": True,
                "nodes": [],
                "edges": [],
                "meta": {"sheets": ["Next"]},
            },
            "tasks": [],
        },
    )
    playwright.expect(page.locator("#project-title")).to_have_text("Next workbook")
    assert page.locator("#node-detail h2").count() == 0
    assert page.locator("#sheet-filter option").all_text_contents() == [
        "Toutes les feuilles",
        "Next",
    ]
    assert page.locator("#graph-sheet option").all_text_contents() == [
        "Choisir une feuille…",
        "Next",
    ]
    assert not errors


def test_failed_back_access_keeps_optimistic_target_and_local_retry(app):
    page, _, errors, _, _ = app
    page.locator("#node-list [data-node='First!A1']").click()
    page.locator("#node-detail h2").wait_for()
    page.locator("#tab-graph").click()
    page.locator("#graph-sheet").select_option("Second")
    heading = page.locator("#graph-preview .graph-reader-heading h3")
    playwright.expect(heading).to_contain_text("Second")
    page.locator("#graph-depth").fill("2")
    page.locator("#graph-depth").press("Tab")
    page.wait_for_function(
        "document.querySelector('#graph-notice').textContent.includes('profondeur 2')"
    )
    camera = page.locator("#graph-canvas").evaluate(
        "e=>({zoom:e._cy.zoom(),pan:e._cy.pan()})"
    )

    def fail_first(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params["nodeId"] == ["First!A1"]:
            route.abort("failed")
        else:
            route.fallback()

    pattern = "**/api/projects/p/neighborhood?*"
    page.route(pattern, fail_first)
    page.locator("#graph-back").click()
    playwright.expect(heading).to_contain_text("First")
    playwright.expect(
        page.locator("#graph-preview .node-access-error")
    ).not_to_be_empty()
    assert page.locator("#message").inner_text() == ""
    assert page.locator("#graph-back").is_disabled()
    assert (
        page.locator("#graph-canvas").evaluate(
            "e=>({zoom:e._cy.zoom(),pan:e._cy.pan()})"
        )
        == camera
    )
    page.unroute(pattern, fail_first)
    open_graph_options(page)
    page.locator("#graph-selected").click()
    playwright.expect(page.locator("#node-detail h2")).to_contain_text("First")
    playwright.expect(page.locator("#node-detail .node-access-error")).to_have_count(0)
    assert not errors


@pytest.mark.parametrize("stage", ["neighborhood", "graph-view"])
def test_explicit_group_navigation_cancels_staged_node_selection(app, stage):
    page, requests, errors, _, fulfill = app
    page.locator("#tab-graph").click()
    page.locator("#graph-sheet").select_option("Second")
    heading = page.locator("#graph-preview .graph-reader-heading h3")
    playwright.expect(heading).to_contain_text("Second")
    if stage == "neighborhood":
        page.locator("#graph-depth").fill("2")
        page.locator("#graph-depth").press("Tab")
        page.wait_for_function(
            "document.querySelector('#graph-notice').textContent"
            ".includes('profondeur 2')"
        )
    open_graph_options(page)
    page.locator("#graph-build").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._cy?.nodes()"
        ".map(n=>n.id()).sort().join(',')==='left,right'"
    )
    pending = []

    def delay_first(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params.get("nodeId") == ["First!A1"]:
            pending.append(route)
        else:
            route.fallback()

    page.route(f"**/api/projects/p/{stage}?*", delay_first)
    page.locator("#graph-sheet").select_option("First")
    page.wait_for_timeout(100)
    assert pending
    playwright.expect(heading).to_contain_text("First")
    page.locator("#graph-list").evaluate("e=>e.open=true")
    page.locator("[data-structure-node='0']").click()
    page.wait_for_timeout(100)
    assert any(
        path.endswith("/graph-view") and p.get("view") == ["1"] for path, p in requests
    )
    response = {
        "nodes": [
            {
                "id": "First!A1",
                "sheet": "First",
                "cell": "A1",
                "kind": "input",
                "value": 1,
            }
        ],
        "edges": [],
    }
    if stage == "graph-view":
        response.update(
            {
                "viewId": -1,
                "breadcrumbs": [{"viewId": -1, "label": "First"}],
                "scopeNodeCount": 1,
                "meta": {"generation": 1, "nodeCount": 3, "edgeCount": 0},
            }
        )
        response["nodes"] = [
            {
                "id": "first-leaf",
                "nodeId": "First!A1",
                "kind": "input",
                "label": "First!A1",
                "internalEdges": 0,
                "position": {"x": 0, "y": 0},
            }
        ]
    try:
        fulfill(pending.pop(), response)
    except playwright.Error:
        pass
    page.wait_for_timeout(100)
    assert "First" in heading.inner_text()
    assert page.locator("#graph-canvas").evaluate(
        "e=>e._cy.nodes().map(n=>n.id()).sort()"
    ) == ["left", "right"]
    open_graph_options(page)
    page.locator("#graph-selected").click()
    assert "First" in page.locator("#node-detail h2").inner_text()
    assert not errors
