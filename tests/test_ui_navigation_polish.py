"""Primary commands stay reachable, and changing cells starts a fresh reader."""

import pytest
import test_app_captures_browser as capture_tests
import test_app_viewer_regressions as serve_tests
import test_viewer_browser as export_tests
from graph_test_helpers import open_workspace_menu

from linexcel import viewer

app = serve_tests.app
browser = export_tests.browser
report = export_tests.report


def test_serve_compact_primary_commands_and_one_row_depth(app):
    page, _, errors, _, _ = app
    page.locator("#tab-graph").click()
    page.wait_for_function("document.querySelector('#graph-canvas')._cy != null")
    assert not page.locator("#graph-options").evaluate("e => e.open")
    bounds = page.evaluate("""() => {
      const rect = id => document.getElementById(id).getBoundingClientRect();
      return {canvasOffset: rect('graph-canvas').top - rect('view-graph').top,
        depth: ['graph-depth-minus','graph-depth','graph-depth-plus'].map(id => {
          const r=rect(id);return {y:r.y, height:r.height, left:r.left, right:r.right};
        }), overflow: [...document.querySelectorAll(
          '#view-graph button,#view-graph input,#view-graph select')]
          .filter(e=>e.getClientRects().length).filter(e=>{
            const r=e.getBoundingClientRect();
            return r.left < -1 || r.right > innerWidth+1;
          }).map(e=>e.id)};
    }""")
    export_tests.capture(page, f"serve-compact-{page.viewport_size['width']}")
    assert bounds["canvasOffset"] < 400
    assert (
        max(b["y"] for b in bounds["depth"]) - min(b["y"] for b in bounds["depth"]) < 3
    )
    assert not bounds["overflow"]
    assert page.locator("#graph-sheet").is_visible()
    if page.viewport_size["width"] <= 650:
        assert page.locator("#graph-read-cell").is_visible()
    else:
        assert page.locator("#graph-reader-toggle").is_visible()
    assert page.locator("#graph-reset").is_hidden()
    page.locator("#graph-options > summary").focus()
    page.keyboard.press("Enter")
    assert page.locator("#graph-reset").is_visible()
    page.locator("#graph-reset").focus()
    assert page.locator("#graph-reset").evaluate("e => e === document.activeElement")
    assert not errors


def test_serve_workbook_and_macro_guidance_are_distinct(app):
    page, requests, errors, _, fulfill = app
    page.locator("#project-title").get_by_text("Navigation", exact=True).wait_for()
    assert "mode choisi" not in page.locator("#node-detail").inner_text()
    assert "Chaque accès" in page.locator("#node-detail").inner_text()
    project = {"id": "p", "name": "Static VBA", "revision": 2}
    graph = {
        "nodes": [
            {
                "id": "vp:M.Run",
                "kind": "vba",
                "label": "M.Run",
                "code": "Sub Run()\nEnd Sub",
            }
        ],
        "edges": [],
        "meta": {"macroOnly": True, "sheets": []},
    }
    page.route(
        "**/api/projects/p",
        lambda route: fulfill(
            route,
            {
                "project": project,
                "graph": graph,
                "tasks": [],
            },
        ),
    )
    open_workspace_menu(page)
    page.locator('[data-project="p"]').click()
    page.locator("#project-title").get_by_text("Static VBA", exact=True).wait_for()
    assert "procédure" in page.locator("#node-detail").inner_text()
    assert "Aucune macro exécutée" in page.locator("#recalculation-policy").inner_text()
    assert "cellule recalculée" in page.locator("#recalculation-policy").inner_text()
    page.locator("#tab-graph").click()
    assert page.locator("#graph-sheet").is_hidden()
    assert (
        page.locator("#graph-search").get_attribute("placeholder")
        == "Procédure, module ou appel…"
    )
    assert (
        "procédure sélectionnée"
        in page.locator(".graph-reader-heading").inner_text().lower()
    )
    assert page.locator("#graph-tab-calculation").inner_text() == "Source VBA"
    assert "Sub Run()" in page.locator("#graph-panel-calculation").inner_text()
    assert "cellules de la feuille" not in page.locator("#graph-notice").inner_text()
    assert "Non disponible" not in page.locator("#graph-canvas").evaluate(
        "e => e._cy.nodes().map(n=>n.data('label')).join(' ')"
    )
    assert page.locator("#graph-recalculate").is_hidden()
    export_tests.capture(page, f"serve-vba-{page.viewport_size['width']}")
    assert not [body for endpoint, body in requests if endpoint.endswith("/tasks")]
    assert not errors


def search_one(page, label):
    page.locator("#lin-search").fill(label)
    page.locator("#lin-search").press("Enter")
    page.locator("#lin-search-results button.lin-result").filter(has_text=label).click()
    page.wait_for_function(
        "label => document.querySelector('#lin-panel h2')?.textContent === label",
        arg=label,
    )
    page.wait_for_function("!cy.animated()")


@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_export_new_cell_resets_reader_and_history_restores_it(report, size):
    graph = export_tests.filtered_graph()
    for node in graph["nodes"]:
        node["doc"] = "\n\n".join(
            f"Long documentation paragraph {i}." for i in range(100)
        )
    page, errors = report(graph, width=size[0], height=size[1])
    search_one(page, "Input")
    panel = page.locator("#lin-panel")
    panel.evaluate("e => { e.scrollTop=420; }")
    old_scroll = panel.evaluate("e => e.scrollTop")
    assert old_scroll > 300
    camera = page.evaluate("({zoom:cy.zoom(),pan:cy.pan()})")
    search_one(page, "Output")
    assert panel.evaluate("e => e.scrollTop") == 0
    assert panel.locator(".lin-nav").first.evaluate(
        "e => e.getBoundingClientRect().top"
    ) < panel.locator(".lin-ai-box").first.evaluate(
        "e => e.getBoundingClientRect().top"
    )
    page.locator("#lin-history-back").click()
    assert panel.evaluate("e => e.scrollTop") == pytest.approx(old_scroll, abs=1)
    assert page.evaluate("({zoom:cy.zoom(),pan:cy.pan()})") == camera
    assert not errors


@pytest.mark.parametrize("suffix", ["xlam", "xla"])
@pytest.mark.parametrize("language", ["fr", "en"])
def test_macro_options_keep_static_scope_in_both_languages(app, suffix, language):
    page, requests, errors, _, fulfill = app
    procedures = [
        {
            "id": "vp:M.Run",
            "kind": "vba",
            "label": "M.Run",
            "code": "Sub Run()\nCall Other\nEnd Sub",
        },
        {
            "id": "vp:M.Other",
            "kind": "vba",
            "label": "M.Other",
            "code": "Sub Other()\nEnd Sub",
        },
    ]
    edges = [
        {"source": "vp:M.Run", "target": "vp:M.Other", "kind": "call", "approx": True}
    ]
    graph = {
        "lazy": True,
        "nodes": procedures,
        "edges": edges,
        "meta": {"macroOnly": True, "sheets": [], "nodeCount": 2},
    }
    page.route(
        "**/api/projects/p",
        lambda route: fulfill(
            route,
            {
                "project": {"id": "p", "name": f"Static.{suffix}", "revision": 2},
                "graph": graph,
                "tasks": [],
            },
        ),
    )
    page.route(
        "**/api/projects/p/neighborhood?*",
        lambda route: fulfill(
            route,
            {
                "nodes": procedures,
                "edges": edges,
            },
        ),
    )
    open_workspace_menu(page)
    page.locator('[data-project="p"]').click()
    page.locator("#project-title").get_by_text(
        f"Static.{suffix}", exact=True
    ).wait_for()
    page.locator("#ui-language").select_option(language)
    page.locator("#tab-graph").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._cy?.nodes().length===2"
    )
    page.locator("#graph-options > summary").click()
    options = page.locator("#graph-options")
    assert page.locator("#graph-sheet-inventory").is_hidden()
    assert page.locator("#graph-range-key").is_hidden()
    text = options.inner_text().lower()
    assert "cellule" not in text and "feuille" not in text
    assert "cell" not in text and "sheet" not in text
    assert (
        "relation vba statique" if language == "fr" else "static vba relationship"
    ) in text
    assert ("sans exécution" if language == "fr" else "without execution") in text
    assert ("procédure" if language == "fr" else "procedure") in page.locator(
        "#graph-lazy-notice"
    ).inner_text().lower()
    assert page.evaluate("""() => [...document.querySelectorAll(
      '#graph-options button,#graph-options input')]
      .filter(e=>e.getClientRects().length).every(e=>{
        const r=e.getBoundingClientRect();return r.left>=0&&r.right<=innerWidth+1;
      })""")
    options.scroll_into_view_if_needed()
    export_tests.capture(
        page, f"macro-options-{suffix}-{language}-{page.viewport_size['width']}"
    )
    assert not [body for endpoint, body in requests if endpoint.endswith("/tasks")]
    assert not errors


@pytest.mark.parametrize("language", ["fr", "en"])
def test_macro_graph_dynamic_guidance_covers_literal_empty_and_hierarchy(app, language):
    page, requests, errors, _, fulfill = app
    literal = {
        "id": "vba-ref:literal",
        "kind": "opaque",
        "label": "implicit_active_workbook:?!A1",
        "referenceType": "vba_literal_unresolved_context",
    }
    graph = {
        "lazy": True,
        "nodes": [literal],
        "edges": [],
        "meta": {"macroOnly": True, "sheets": [], "nodeCount": 1},
    }
    page.route(
        "**/api/projects/p",
        lambda route: fulfill(
            route,
            {
                "project": {"id": "p", "name": "Literal VBA", "revision": 2},
                "graph": graph,
                "tasks": [],
            },
        ),
    )
    page.route(
        "**/api/projects/p/neighborhood?*",
        lambda route: fulfill(route, {"nodes": [literal], "edges": []}),
    )
    open_workspace_menu(page)
    page.locator('[data-project="p"]').click()
    page.locator("#project-title").get_by_text("Literal VBA", exact=True).wait_for()
    page.locator("#ui-language").select_option(language)
    page.locator("#tab-graph").click()
    page.wait_for_function("!document.querySelector('.graph-pending-overlay')")
    assert (
        "nœud sélectionné" if language == "fr" else "selected node"
    ) in page.locator(".graph-reader-heading").inner_text().lower()
    assert (
        "relation indexée pour cette cellule"
        not in page.locator("#graph-notice").inner_text()
    )
    page.locator("#graph-list > summary").click()
    assert (
        "relations statiques" if language == "fr" else "static source relationships"
    ) in page.locator("#graph .graph-direction").inner_text().lower()
    page.locator("#graph-options > summary").click()
    page.locator("#graph-build").click()
    page.wait_for_function(
        "document.querySelector('#graph-neighborhood').hidden === false"
    )
    text = page.locator("#graph-notice").inner_text().lower()
    assert ("graphe vba statique" if language == "fr" else "static vba graph") in text
    assert "plage" not in text and "range" not in text
    assert "cell" not in page.locator("#graph-build-status").inner_text().lower()
    graph["nodes"] = []
    open_workspace_menu(page)
    page.locator('[data-project="p"]').click()
    page.wait_for_function(
        "text => document.querySelector('#graph-notice').textContent === text",
        arg="Aucun nœud disponible." if language == "fr" else "No node available.",
    )
    assert page.locator("#graph-notice").inner_text() == (
        "Aucun nœud disponible." if language == "fr" else "No node available."
    )
    page.locator("#graph-build").click()
    page.wait_for_function(
        "document.querySelector('#graph-neighborhood').hidden === false"
    )
    assert ("hiérarchie vba" if language == "fr" else "vba hierarchy") in page.locator(
        "#graph-notice"
    ).inner_text().lower()
    assert "cell" not in page.locator("#graph-preview").inner_text().lower()
    assert all(
        body["operation"] == "build_graph"
        for endpoint, body in requests
        if endpoint.endswith("/tasks")
    )
    assert not errors


@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_export_dense_selection_has_a_direct_local_exploration(report, size):
    graph = {
        "nodes": [{"id": "hub", "kind": "cell", "label": "Hub", "sheet": "S"}]
        + [
            {"id": f"n{i:03}", "kind": "input", "label": f"Input {i:03}", "sheet": "S"}
            for i in range(95)
        ],
        "edges": [
            {"id": f"e{i}", "source": f"n{i:03}", "target": "hub", "kind": "ref"}
            for i in range(95)
        ],
    }
    page, errors = report(graph, width=size[0], height=size[1])
    assert page.evaluate("cy.nodes().length") == 96
    search_one(page, "Hub")
    page.evaluate("window.localExplorationFrame = null")
    page.locator("#lin-explore-selection").click()
    page.wait_for_function("""() => {
        const nodes = cy.nodes(':visible');
        const status = document.querySelector('#lin-explore-status');
        if (nodes.length <= 1 || nodes.length > 13 || cy.animated()
                || !status?.getClientRects().length
                || !cy.getElementById('hub').selected()) return false;
        const frame = JSON.stringify({ids: nodes.map(n=>n.id()).sort(),
            bounds: nodes.renderedBoundingBox(), zoom: cy.zoom(), pan: cy.pan()});
        const stable = window.localExplorationFrame === frame;
        window.localExplorationFrame = frame;
        return stable;
    }""")
    export_tests.capture(page, f"export-dense-local-{size[0]}")
    visible = page.evaluate("cy.nodes(':visible').length")
    assert 1 < visible <= 13
    assert page.locator("#lin-explore-status").is_visible()
    assert page.evaluate("cy.getElementById('hub').selected()")
    assert page.evaluate(
        "cy.nodes(':visible').every(n=>"
        "Number(n.style('font-size').replace('px',''))*cy.zoom() >= 9)"
    )
    assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
    assert not errors


def test_export_hover_clears_when_pointer_leaves_the_canvas(report):
    page, errors = report(export_tests.filtered_graph())
    coords = page.evaluate("""() => {
      const b=document.querySelector('#lin-cy').getBoundingClientRect(),
        p=cy.getElementById('c').renderedPosition();
      return {x:b.x+p.x,y:b.y+p.y};
    }""")
    page.mouse.move(coords["x"], coords["y"])
    page.wait_for_function("cy.getElementById('c').hasClass('hovered')")
    page.mouse.move(10, 10)
    page.wait_for_function("cy.nodes('.hovered').length === 0")
    assert not errors


@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_local_pixel_result_does_not_claim_model_generation(browser, size):
    page = browser.new_page(viewport={"width": size[0], "height": size[1]})
    capture = capture_tests._capture("image", result=capture_tests._images("Data"))
    observation = {
        "id": "observation",
        "operation": "describe_capture",
        "status": "succeeded",
        "options": {"captureTaskId": "image", "captureIndex": 0, "language": "fr"},
        "result": {
            "markdown": "Capture uniforme : aucune chaîne de texte lisible.",
            "observationOrigin": "deterministic_pixels",
            "model": None,
            "configuredModel": "MiniMax-M3",
            "language": "fr",
            "usage": {
                "inputTokens": 0,
                "outputTokens": 0,
                "totalTokens": 0,
                "estimated": False,
            },
            "sources": {"captureTaskId": "image", "captureIndex": 0},
        },
    }
    try:
        _, errors = capture_tests._route_app(
            page,
            [{"id": "p", "name": "Pixels"}],
            {"p": [capture, observation]},
            lambda *_: None,
            config={
                "ai": {
                    "available": True,
                    "model": "MiniMax-M3",
                    "visionModel": "MiniMax-M3",
                    "specializedReview": True,
                },
            },
        )
        page.locator("#tab-sheets").click()
        assert (
            "Deux appels texte"
            in page.locator("#workbook-ai .ai-review-policy").inner_text()
        )
        page.locator("#tab-captures").click()
        root = page.locator("#capture-ai")
        assert root.locator(".badge").first.inner_text() == "Contrôle pixel local"
        assert "aucun appel IA" in root.locator(".ai-result-origin").inner_text()
        assert "Modèle prévu : MiniMax-M3" in root.inner_text()
        assert "Comptage fourni par le modèle" not in root.inner_text()
        assert "Total : 0" in root.locator(".token-usage").inner_text()
        assert root.locator(".ai-review-policy").count() == 0
        root.scroll_into_view_if_needed()
        export_tests.capture(page, f"local-pixels-{size[0]}")
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize("missing_js", [False, True])
@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_export_empty_and_missing_renderer_keep_other_views_usable(
    browser, tmp_path, monkeypatch, missing_js, size
):
    if missing_js:
        monkeypatch.setattr(viewer, "_inline_scripts", lambda: "")
    graph = {
        "nodes": [],
        "edges": [],
        "meta": {
            "workbookDoc": "Workbook evidence remains available.",
            "stats": {"totalFormulas": 0, "totalNodes": 0, "totalEdges": 0},
        },
    }
    html = viewer.render_html(graph, language="en")
    path = tmp_path / "fallback.html"
    path.write_text(html, encoding="utf-8")
    page = browser.new_page(viewport={"width": size[0], "height": size[1]})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    try:
        page.goto(path.as_uri())
        page.locator(".lin-fallback" if missing_js else ".lin-emptygraph").wait_for()
        assert page.locator("#lin-tools").is_hidden()
        assert page.locator("#lin-search").is_disabled()
        assert page.locator("#lin-panel").is_hidden()
        page.locator("#lin-tab-overview").click()
        assert page.locator("#lin-overview-main").is_visible()
        assert (
            "Workbook evidence remains available."
            in page.locator("#lin-overview-main").inner_text()
        )
        page.locator("#lin-tab-graph").click()
        export_tests.capture(page, f"export-fallback-{missing_js}-{size[0]}")
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_export_query_with_zero_formulas_has_source_and_unopened_external_notice(
    report, size
):
    graph = {
        "nodes": [
            {
                "id": "q",
                "kind": "query",
                "label": "Import source",
                "code": 'let Source = File.Contents("data.csv") in Source',
                "sources": [
                    {"kind": "file", "target": "data.csv", "function": "File.Contents"}
                ],
            }
        ],
        "edges": [],
        "meta": {
            "stats": {
                "totalFormulas": 0,
                "totalNodes": 1,
                "totalEdges": 0,
                "queries": 1,
            }
        },
    }
    page, errors = report(graph, width=size[0], height=size[1])
    search_one(page, "Import source")
    text = page.locator("#lin-panel").inner_text()
    assert 'File.Contents("data.csv")' in text
    assert "Not loaded onto a sheet" in text
    assert "not read by linexcel" in text
    assert page.locator("#lin-tools").is_visible()
    export_tests.capture(page, f"export-query-{size[0]}")
    assert not errors


@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_missing_renderer_list_pages_searches_and_follows_source_links(
    browser, tmp_path, monkeypatch, size
):
    monkeypatch.setattr(viewer, "_inline_scripts", lambda: "")
    graph = {
        "nodes": [
            {"id": f"n{i:03}", "kind": "input", "label": f"Item {i:03}", "value": i}
            for i in range(100)
        ]
        + [
            {
                "id": "sum",
                "kind": "cell",
                "label": "Total",
                "formula": "=SUM(A1:A100)",
                "sourceFormula": "=SUM(A1:A100)",
                "value": 4950,
                "cachedValue": 4900,
                "valueSource": "engine",
                "cachedAgreement": "differ",
            }
        ],
        "edges": [
            {"id": f"e{i:03}", "source": f"n{i:03}", "target": "sum", "kind": "ref"}
            for i in range(100)
        ],
        "meta": {"stats": {"totalFormulas": 1, "totalNodes": 101, "totalEdges": 100}},
    }
    path = tmp_path / "missing-renderer.html"
    path.write_text(viewer.render_html(graph, language="en"), encoding="utf-8")
    page = browser.new_page(viewport={"width": size[0], "height": size[1]})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    try:
        page.goto(path.as_uri())
        assert page.evaluate("typeof cytoscape === 'undefined'")
        assert page.locator(".lin-fallback-list button").count() == 50
        page.locator(".lin-fallback-pages").get_by_role(
            "button", name="Next", exact=True
        ).click()
        assert (
            page.locator(".lin-fallback-list button").first.inner_text() == "Item 050"
        )
        assert page.locator("#lin-fallback-status").inner_text() == "51–100 / 101"
        page.locator(".lin-fallback-pages").get_by_role(
            "button", name="Next", exact=True
        ).click()
        assert page.locator(".lin-fallback-list button").count() == 1
        page.locator(".lin-fallback-list button").click()
        panel = page.locator("#lin-panel")
        assert panel.locator("h2").inner_text() == "Total"
        assert "=SUM(A1:A100)" in panel.inner_text()
        assert "4950" in panel.inner_text() and "4900" in panel.inner_text()
        assert panel.locator("#lin-fit-neighbors").is_hidden()
        assert panel.locator(".lin-nav button").count() == 50
        panel.get_by_role("button", name="Show more", exact=True).click()
        assert panel.locator(".lin-nav button").count() == 100
        panel.get_by_role("button", name="Item 099", exact=True).click()
        assert panel.locator("h2").inner_text() == "Item 099"
        panel.get_by_role("button", name="Total", exact=True).click()
        assert panel.locator("h2").inner_text() == "Total"
        panel.get_by_role("button", name="Close", exact=True).click()
        assert panel.is_hidden()
        search = page.locator("#lin-fallback-search")
        search.fill("Item 073")
        assert page.locator(".lin-fallback-list button").count() == 1
        search.press("Tab")
        page.locator(".lin-fallback-list button").focus()
        page.keyboard.press("Enter")
        assert panel.locator("h2").inner_text() == "Item 073"
        export_tests.capture(page, f"export-missing-renderer-reader-{size[0]}")
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
        assert not errors
    finally:
        page.close()
