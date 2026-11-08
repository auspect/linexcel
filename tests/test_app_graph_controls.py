"""Graph navigation follows the full sheet inventory and bounded API depth."""

import json
import os
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

import pytest
from graph_test_helpers import open_graph_depth, open_graph_options, open_workspace_menu

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
def graph_app(browser, request):
    page = browser.new_page(
        viewport={"width": request.param[0], "height": request.param[1]}
    )
    sheets = ["Z first", *[f"Sheet {i:02}" for i in range(1, 31)]]
    nodes = [
        {
            "id": f"Z first!A{i}",
            "sheet": sheets[0],
            "cell": f"A{i}",
            "kind": "input",
            "value": i,
        }
        for i in range(1, 26)
    ]
    # Directions alternate: A3 is a sibling source of A2, not a source of A1.
    edges = [
        {
            "source": f"Z first!A{i if i % 2 else i + 1}",
            "target": f"Z first!A{i + 1 if i % 2 else i}",
        }
        for i in range(1, 25)
    ]
    project = {"id": "p", "name": "Graph controls", "revision": 1}
    graph = {
        "lazy": True,
        "nodes": nodes[:1],
        "edges": [],
        "meta": {"sheets": sheets, "nodeCount": 25},
    }
    errors, requests = [], []
    state = {"truncated": False, "lazy": True}
    page.on("pageerror", lambda error: errors.append(str(error)))

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
        params = parse_qs(urlsplit(route.request.url).query)
        requests.append((path, params))
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
            fulfill(route, {})
        elif path == "/api/projects":
            fulfill(route, {"projects": [project]})
        elif path == "/api/projects/p":
            payload = (
                graph
                if state["lazy"]
                else {**graph, "lazy": False, "nodes": nodes, "edges": edges}
            )
            fulfill(route, {"project": project, "graph": payload, "tasks": []})
        elif path == "/api/projects/p/nodes":
            sheet = params.get("sheet", [sheets[0]])[0]
            kind = params.get("kind", [""])[0]
            selected = (
                []
                if kind == "cell"
                else [
                    {
                        "id": f"{sheet}!A1",
                        "sheet": sheet,
                        "cell": "A1",
                        "kind": "input",
                        "value": 27,
                    }
                ]
            )
            fulfill(route, {"nodes": selected, "pagination": {"total": len(selected)}})
        elif path == "/api/projects/p/neighborhood":
            center = params["nodeId"][0]
            if center.startswith("Z first!"):
                row = int(center.split("!A")[1])
                depth = int(params["depth"][0])
                selected = nodes[max(0, row - depth - 1) : min(25, row + depth)]
                ids = {n["id"] for n in selected}
                links = [e for e in edges if {e["source"], e["target"]} <= ids]
            else:
                selected = [
                    {
                        "id": center,
                        "sheet": center.split("!")[0],
                        "cell": "A1",
                        "kind": "input",
                        "value": 27,
                    }
                ]
                links = []
            fulfill(
                route,
                {
                    "nodes": selected,
                    "edges": links,
                    "pagination": {"truncated": state["truncated"]},
                },
            )
        elif path == "/api/projects/p/tasks":
            body = route.request.post_data_json
            assert body["operation"] == "evaluate"
            assert body["force"] is True
            fulfill(
                route,
                {
                    "task": {
                        "id": "targeted-evaluation",
                        "operation": "evaluate",
                        "nodeId": body["nodeId"],
                        "status": "succeeded",
                        "result": {"status": "unsupported"},
                    }
                },
            )
        else:
            fulfill(route, {})

    page.route("**/*", handle)
    page.goto("http://localhost/")
    open_workspace_menu(page)
    page.locator("[data-project='p']").click()
    yield page, sheets, requests, errors, state
    page.close()


def test_sheet_picker_retains_complete_workbook_order_and_option_identity(graph_app):
    page, sheets, _, errors, _ = graph_app
    page.locator("#tab-graph").click()
    options = page.locator("#graph-sheet option:not([value=''])")
    assert options.all_text_contents() == sheets
    option = options.last.element_handle()
    open_graph_options(page)
    page.locator("#graph-reset").click()
    assert option.evaluate("el=>el.isConnected")
    open_graph_options(page)
    assert "31" in page.locator("#graph-sheet-inventory").inner_text()
    page.locator("#graph-sheet").select_option(sheets[-1])
    playwright.expect(
        page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text(sheets[-1])
    playwright.expect(page.locator("#graph-sheet-status")).to_contain_text(sheets[-1])
    playwright.expect(page.locator("#graph-notice")).to_contain_text(
        "Aucune relation indexée"
    )
    assert options.all_text_contents() == sheets
    assert not errors


def test_search_dismissal_focus_overlay_and_bidirectional_routes(graph_app):
    page, _, _, errors, _ = graph_app
    page.locator("#tab-graph").click()
    page.locator("#graph-search").fill("A1")
    page.locator("#graph-results [data-related]").first.wait_for()
    page.locator("#graph-search").press("Escape")
    assert not page.locator("#graph-results").inner_text()
    page.locator("#graph-search").fill("A2")
    page.locator("#graph-results [data-related]").first.wait_for()
    page.locator("#project-title").click()
    assert not page.locator("#graph-results").inner_text()
    page.locator("#graph-fullscreen").click()
    bounds = page.locator("#view-graph").bounding_box()
    assert bounds["x"] == 0 and bounds["y"] == 0
    assert bounds["width"] == page.viewport_size["width"]
    assert bounds["height"] == page.viewport_size["height"]
    page.keyboard.press("Escape")
    assert page.locator("#view-graph").get_attribute("aria-modal") is None
    page.locator("#graph-canvas").evaluate("""e=>e._flow.render(
      [{id:'a',title:'A'},{id:'b',title:'B'}],
      [{source:'a',target:'b'},{source:'b',target:'a'}])""")
    page.wait_for_function(
        "()=>document.querySelectorAll('.react-flow__edge-path').length===2"
    )
    page.wait_for_timeout(100)
    crossed = page.locator("#graph-canvas").evaluate("""e=>{
      const boxes=e._flow.getNodes().map(n=>({x:n.position.x,y:n.position.y}));
      return [...e.querySelectorAll('.react-flow__edge-path')].some(path=>{
        const length=path.getTotalLength();
        for(let i=1;i<100;i++){const p=path.getPointAtLength(length*i/100);
          if(boxes.some(b=>p.x>b.x+2&&p.x<b.x+230&&p.y>b.y+2&&p.y<b.y+102))return true;}
        return false;});} """)
    assert not crossed
    assert not errors


def test_large_fanout_stays_readable_and_opens_bounded_indexed_cells(graph_app):
    page, _, _, errors, _ = graph_app
    center = {
        "id": "Z first!A1",
        "kind": "cell",
        "sheet": "Z first",
        "cell": "A1",
        "value": 1,
        "formula": "=1",
    }
    source = {
        "id": "Ledger!J2:J249676",
        "kind": "range",
        "sheet": "Ledger",
        "cell_count": 249675,
    }
    members = [
        {
            "id": f"Forecast!D{i}",
            "kind": "cell",
            "sheet": "Forecast",
            "cell": f"D{i}",
            "formula": f"=SUM(Ledger!J2:J249676)+A{i}",
            "patternId": "copied",
            "patternMemberCount": 12483,
        }
        for i in range(2, 200)
    ]
    nodes = [center, source, *members]
    edges = [{"source": source["id"], "target": n["id"]} for n in [center, *members]]
    page.route(
        "**/api/projects/p/neighborhood?*",
        lambda route: route.fulfill(
            json={
                "nodes": nodes,
                "edges": edges,
                "depth": 3,
                "pagination": {"returned": 200, "limit": 200, "truncated": True},
            }
        ),
    )
    page.route(
        "**/api/projects/p/graph-view?*",
        lambda route: route.fulfill(
            json={
                "viewId": -1,
                "scopeNodeCount": 32,
                "meta": {"nodeCount": 12483, "edgeCount": 198},
                "breadcrumbs": [
                    {"viewId": 0, "label": "Classeur"},
                    {"viewId": -1, "label": "Forecast pattern"},
                ],
                "nodes": [
                    {
                        "id": str(i),
                        "nodeId": n["id"],
                        "kind": "cell",
                        "label": n["id"],
                        "internalEdges": 0,
                        "nodeCount": 1,
                    }
                    for i, n in enumerate(members[:32])
                ],
                "edges": [],
            }
        ),
    )
    page.locator("#tab-graph").click()
    page.wait_for_function(
        "()=>document.querySelector('#graph-canvas')._flow?.getNodes().length>0"
    )
    open_graph_depth(page)
    page.locator("#graph-depth").fill("3")
    page.locator("#graph-depth").dispatch_event("change")
    page.locator("#graph-options").evaluate("e=>e.open=false")
    page.wait_for_function(
        "()=>document.querySelector('#graph-canvas')._flow?.getNodes().some(n=>n.data.original.aggregate)"
    )
    page.locator(".react-flow__controls-fitview").click()
    page.wait_for_timeout(150)
    data = page.locator("#graph-canvas").evaluate("""e=>{
      const flow=e._flow,group=flow.getNodes().find(n=>n.data.original.aggregate);
      return {cards:flow.getNodes().length,zoom:flow.getViewport().zoom,
        group:group.data.original,weights:flow.getEdges().map(e=>e.weight)};
    }""")
    assert data["cards"] == 3
    assert data["zoom"] >= 0.35
    assert len(data["group"]["members"]) == 198
    assert data["group"]["patternTotal"] == 12483
    assert sorted(data["weights"]) == [1, 198]
    assert "200" in page.locator("#graph-notice").text_content()
    page.locator("#graph-canvas").evaluate(
        "(e,id)=>e._flow.select(id)", data["group"]["id"]
    )
    page.wait_for_function(
        "()=>document.querySelector('#graph-canvas')._flow?.getNodes().length===32"
    )
    assert (
        page.locator("#graph-breadcrumbs").text_content().endswith("Forecast pattern")
    )
    assert not errors


def test_pan_keeps_text_nodes_and_translation_ignores_flow_owned_mutations(graph_app):
    page, _, _, errors, _ = graph_app
    page.locator("#ui-language").select_option("en")
    page.locator("#tab-graph").click()
    page.wait_for_function(
        "()=>document.querySelector('#graph-canvas')._flow?.getNodes().length===2"
    )
    page.wait_for_timeout(100)
    result = page.evaluate("""async()=>{
      const frame=()=>new Promise(r=>requestAnimationFrame(r));
      const flow=document.querySelector('#graph-canvas')._flow;
      await flow.setViewport({...flow.getViewport(),zoom:.2});
      for(let i=0;i<3;i++)await frame();
      const zoom=document.querySelector('#graph-zoom-level');
      const hint=document.querySelector('#graph-readability');
      const before=[zoom.firstChild,hint.firstChild];
      for(let i=0;i<6;i++){flow.panBy({x:i%2?2:-2,y:0});await frame();}
      const preserved=before[0]===zoom.firstChild&&before[1]===hint.firstChild;
      const original=document.createTreeWalker;let walks=0;
      document.createTreeWalker=function(...args){
        walks++;return original.apply(this,args)};
      const owned=document.createElement('span');owned.textContent='Recalculer';
      document.querySelector('.react-flow__viewport').append(owned);
      await frame();await frame();
      const ignored=walks===0&&owned.textContent==='Recalculer';owned.remove();
      const legacy=document.createElement('span');legacy.textContent='Recalculer';
      document.querySelector('.graph-docked-controls').append(legacy);
      await frame();await frame();
      const translated=walks>0&&legacy.textContent==='Recalculate';
      legacy.remove();document.createTreeWalker=original;
      return {preserved,ignored,translated};
    }""")
    assert result == {"preserved": True, "ignored": True, "translated": True}
    assert not errors


def test_search_escapes_canvas_clip_and_uniform_bus_labels_are_hidden(graph_app):
    page, _, _, errors, _ = graph_app
    page.locator("#tab-graph").click()
    matches = [
        {"id": f"Other!A{i}", "cell": f"A{i}", "sheet": "Other", "kind": "input"}
        for i in range(1, 21)
    ]
    page.route(
        "**/api/projects/p/nodes?*",
        lambda route: route.fulfill(
            json={"nodes": matches, "pagination": {"total": 20}}
        ),
    )
    page.locator("#graph-search").fill("Other")
    page.locator("#graph-results [data-related]").first.wait_for()
    assert page.locator("#graph-results").evaluate("e=>e.matches(':popover-open')")
    assert page.locator("#graph-results").evaluate("""e=>{
      const b=e.getBoundingClientRect();
      const hit=document.elementFromPoint(b.x+12,b.bottom-12);
      return hit?.closest('#graph-results')===e;
    }""")
    page.locator("#graph-search").press("Escape")
    assert not page.locator("#graph-results").evaluate("e=>e.matches(':popover-open')")
    page.locator("#graph-canvas").evaluate("""e=>{
      const cells=Array.from({length:32},(_,i)=>(
        {id:'c'+i,nodeId:'c'+i,title:'Cell '+i}));
      e._flow.render([...cells,{id:'rest',tone:'outside',title:'Rest'}],
        cells.flatMap(n=>[{source:n.id,target:'rest',weight:5},
          {source:'rest',target:n.id,weight:1}]));} """)
    assert page.locator("#graph-canvas").evaluate(
        "e=>e._flow.getEdges().every(edge=>[1,5].includes(edge.weight)&&!edge.label)"
    )
    assert not errors


@pytest.mark.parametrize("lazy", [True, False])
def test_depth_stepper_follows_mixed_direction_paths_beyond_two_levels(graph_app, lazy):
    page, _, requests, errors, state = graph_app
    state["lazy"] = lazy
    open_workspace_menu(page)
    page.locator("[data-project='p']").click()
    page.locator("#tab-graph").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._flow?.getNodes().length===2"
    )
    open_graph_depth(page)
    page.locator("#graph-depth-plus").click()
    open_graph_depth(page)
    page.locator("#graph-depth-plus").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._flow?.getNodes().length===4"
    )
    canvas = page.locator("#graph-canvas")
    assert canvas.evaluate("el=>el._flow.getNodes().map(n=>n.id)") == [
        f"Z first!A{i}" for i in range(1, 5)
    ]
    assert canvas.evaluate("el=>el._flow.getNode('Z first!A3').data.tone==='context'")
    open_graph_options(page)
    assert page.locator("#graph-neutral-key").get_attribute("hidden") is None
    for depth, count in [(5, 6), (20, 21)]:
        open_graph_depth(page)
        page.locator("#graph-depth").fill(str(depth))
        page.locator("#graph-depth").press("Tab")
        page.wait_for_function(
            "n=>document.querySelector('#graph-canvas')._flow?.getNodes().length===n",
            arg=count,
        )
        assert f"profondeur {depth}" in page.locator("#graph-notice").inner_text()
    assert page.locator("#graph-depth-plus").is_disabled()
    open_graph_depth(page)
    page.locator("#graph-depth-minus").click()
    assert page.locator("#graph-depth").input_value() == "19"
    if lazy:
        assert any(params.get("depth") == ["20"] for _, params in requests)
    assert not errors


def test_server_budget_truncation_remains_visible_when_missing_nodes_are_not_sent(
    graph_app,
):
    page, _, _, errors, state = graph_app
    state["truncated"] = True
    page.locator("#tab-graph").click()
    playwright.expect(page.locator("#graph-notice")).to_contain_text(
        "Limite d’affichage atteinte"
    )
    page.locator("#tab-captures").click()
    assert page.locator("#graph-notice").is_hidden()
    assert not errors


def test_pending_depth_keeps_the_previous_view_truthfully_labelled(graph_app):
    page, _, _, errors, _ = graph_app
    page.locator("#tab-graph").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._flow?.getNodes().length===2"
    )
    held = []

    def delay_third(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params["depth"] == ["3"]:
            held.append(route)
        else:
            route.fallback()

    page.route("**/api/projects/p/neighborhood?*", delay_third)
    open_graph_depth(page)
    page.locator("#graph-depth").fill("3")
    page.locator("#graph-depth").press("Tab")
    page.wait_for_timeout(100)
    assert held
    open_graph_options(page)
    page.locator("#graph-reset").click()
    assert "profondeur 1" in page.locator("#graph-notice").inner_text()
    assert "profondeur 3" in page.locator("#graph-depth-help").inner_text()
    held.pop().fulfill(
        content_type="application/json",
        body=json.dumps(
            {
                "nodes": [
                    {
                        "id": f"Z first!A{i}",
                        "sheet": "Z first",
                        "cell": f"A{i}",
                        "kind": "input",
                        "value": i,
                    }
                    for i in range(1, 5)
                ],
                "edges": [
                    {"source": "Z first!A1", "target": "Z first!A2"},
                    {"source": "Z first!A3", "target": "Z first!A2"},
                    {"source": "Z first!A3", "target": "Z first!A4"},
                ],
            }
        ),
    )
    playwright.expect(page.locator("#graph-notice")).to_contain_text("profondeur 3")
    assert "Chargement" not in page.locator("#graph-depth-help").inner_text()
    assert not errors


def test_depth_failure_stops_loading_and_retains_previous_graph(graph_app):
    page, _, _, errors, _ = graph_app
    page.locator("#tab-graph").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._flow?.getNodes().length===2"
    )

    def reject_third(route):
        params = parse_qs(urlsplit(route.request.url).query)
        if params["depth"] == ["3"]:
            route.fulfill(
                status=503,
                content_type="application/json",
                body=json.dumps({"error": {"message": "Index indisponible"}}),
            )
        else:
            route.fallback()

    page.route("**/api/projects/p/neighborhood?*", reject_third)
    open_graph_depth(page)
    page.locator("#graph-depth").fill("3")
    page.locator("#graph-depth").press("Tab")
    playwright.expect(page.locator("#graph-depth-help")).to_contain_text(
        "Index indisponible"
    )
    assert "Chargement" not in page.locator("#graph-depth-help").inner_text()
    assert "profondeur 1" in page.locator("#graph-notice").inner_text()
    assert page.locator("#message").inner_text() == ""
    page.locator("#tab-captures").click()
    assert page.locator("#graph-depth-help").is_hidden()
    page.locator("#tab-graph").click()
    open_graph_depth(page)
    page.locator("#graph-depth").fill("1")
    page.locator("#graph-depth").press("Tab")
    playwright.expect(page.locator("#graph-depth-help")).not_to_contain_text(
        "Index indisponible"
    )
    assert not errors


def test_lazy_range_in_another_node_page_cannot_invent_neighborhood_edges(graph_app):
    page, _, _, errors, _ = graph_app
    input_node = {
        "id": "QRA!H2",
        "sheet": "QRA",
        "cell": "H2",
        "row": 2,
        "column": 8,
        "kind": "input",
        "value": "[qra.xlsm]QRA",
    }
    formula_node = {
        "id": "QRA!H3",
        "sheet": "QRA",
        "cell": "H3",
        "row": 3,
        "column": 8,
        "kind": "cell",
        "formula": '=MID(H2,FIND("]",H2)+1,100)',
    }
    off_page_range = {
        "id": "QRA!A1:XFD1048576",
        "sheet": "QRA",
        "kind": "range",
        "bounds": [1, 1, 1048576, 16384],
        "reference": "A1:XFD1048576",
    }
    foreign_input = {
        "id": "QRA!Z1",
        "sheet": "QRA",
        "cell": "Z1",
        "row": 1,
        "column": 26,
        "kind": "input",
        "value": 99,
    }
    link = {"source": "QRA!H2", "target": "QRA!H3"}
    graph = {
        "lazy": True,
        "nodes": [input_node, formula_node, off_page_range, foreign_input],
        "edges": [{"source": "QRA!Z1", "target": "QRA!H3"}],
        "meta": {"sheets": ["QRA"], "nodeCount": 4},
    }
    neighborhood = {
        "nodes": [input_node, formula_node],
        "edges": [link],
        "pagination": {"limit": 200, "returned": 2, "truncated": False},
    }
    page.route(
        "**/api/projects/p",
        lambda route: route.fulfill(
            json={
                "project": {"id": "p", "name": "QRA range", "revision": 1},
                "graph": graph,
                "tasks": [],
            }
        ),
    )
    held = []

    def load_neighborhood(route):
        if not held:
            held.append(route)
        else:
            route.fulfill(json=neighborhood)

    page.route("**/api/projects/p/neighborhood?*", load_neighborhood)
    open_workspace_menu(page)
    page.locator("[data-project='p']").click()
    page.locator("#tab-graph").click()
    canvas = page.locator("#graph-canvas")
    page.wait_for_timeout(50)
    assert len(held) == 1
    pending_notice = page.locator("#graph-notice").inner_text()
    held[0].fulfill(json=neighborhood)
    assert "source(s)" not in pending_notice
    assert "usage(s)" not in pending_notice
    assert "Aucune relation indexée" not in pending_notice
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._flow?.getNodes().length===2"
    )

    def assert_exact_neighborhood():
        # Selecting a new center keeps the previous canvas while its request loads.
        playwright.expect(page.locator("#graph-notice")).to_contain_text("2 nœuds")
        assert set(canvas.evaluate("el=>el._flow.getNodes().map(n=>n.id)")) == {
            "QRA!H2",
            "QRA!H3",
        }
        assert canvas.evaluate(
            "el=>el._flow.getEdges().map(e=>({source:e.source,target:e.target}))"
        ) == [link]
        assert (
            canvas.evaluate(
                "el=>el._flow.getEdges().filter(e=>e.className==='membership').length"
            )
            == 0
        )
        assert "2 nœuds" in page.locator("#graph-notice").inner_text()
        assert off_page_range["id"] not in page.locator("#graph").inner_text()
        assert foreign_input["id"] not in page.locator("#graph").inner_text()

    assert_exact_neighborhood()
    assert "1 source(s), 0 usage(s)" in page.locator("#graph-notice").inner_text()
    page.locator("#graph-list summary").click()
    page.locator("#graph [data-graph-node='QRA!H2']").click()
    playwright.expect(
        page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("H2")
    assert_exact_neighborhood()
    assert "0 source(s), 1 usage(s)" in page.locator("#graph-notice").inner_text()
    open_graph_depth(page)
    page.locator("#graph-depth").fill("20")
    page.locator("#graph-depth").press("Tab")
    playwright.expect(page.locator("#graph-notice")).to_contain_text("profondeur 20")
    assert_exact_neighborhood()
    assert not errors
