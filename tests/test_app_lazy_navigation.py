"""Lazy browser navigation keeps the reader and canvas usable during I/O."""

import os
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

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


@dataclass
class Navigation:
    page: object
    nodes: list[dict]
    requests: list = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    tasks: dict = field(default_factory=dict)
    evaluations: list[dict] = field(default_factory=list)
    held: list = field(default_factory=list)
    delayed_nodes: set = field(default_factory=set)
    rejected_nodes: set = field(default_factory=set)
    held_evaluations: set = field(default_factory=set)
    saved_evaluation: dict | None = None
    hold_saved_result: bool = False
    reject_saved_result: bool = False
    finish_calculations: bool = False

    def neighborhood(self, node_id):
        return {
            "nodes": self.nodes,
            "edges": [
                {"source": "Source!A1", "target": "Calc!B1"},
                {"source": "Source!A1", "target": "Calc!B2"},
                {"source": "Source!A1", "target": "Calc!B3"},
                {"source": "Calc!B1", "target": "Calc!B2"},
                {"source": "Calc!B1", "target": "Calc!B3"},
                {"source": "Calc!B2", "target": "Calc!B3"},
            ],
            "pagination": {"limit": 200, "returned": 4, "truncated": False},
            "coverage": {"center": node_id, "depth": 1},
        }

    def route(self, route):
        url = urlsplit(route.request.url)
        path = url.path.removeprefix("/mounted/")
        query = parse_qs(url.query, keep_blank_values=True)
        self.requests.append((route.request.method, path, query))
        if not path:
            route.fulfill(body=APP.read_text("utf-8"), content_type="text/html")
            return
        if path.endswith("cytoscape.min.js"):
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text("utf-8"),
                content_type="application/javascript",
            )
            return
        project = {"id": "p", "name": "Lazy navigation", "revision": 1}
        if path == "api/config":
            data = {"limits": {"memoryMb": 2048, "uploadMb": 64}}
        elif path == "api/projects":
            data = {"projects": [project]}
        elif path == "api/projects/p":
            data = {
                "project": project,
                "graph": {
                    "lazy": True,
                    "nodes": self.nodes,
                    "edges": [],
                    "formulaPatterns": [],
                    "meta": {"sheets": ["Source", "Calc"], "nodeCount": 1000000},
                    "pagination": {"offset": 0, "limit": 200, "total": 1000000},
                },
                "tasks": [
                    {
                        "id": "saved-doc",
                        "operation": "document",
                        "nodeId": "Calc!B1",
                        "status": "succeeded",
                        "resultAvailable": True,
                    },
                    *([self.saved_evaluation] if self.saved_evaluation else []),
                ],
            }
        elif path == "api/projects/p/neighborhood":
            node_id = query["nodeId"][0]
            data = self.neighborhood(node_id)
            if node_id in self.delayed_nodes:
                self.held.append((route, data, "neighborhood", node_id))
                return
            if node_id in self.rejected_nodes:
                route.fulfill(
                    status=503, json={"error": {"message": "Links unavailable"}}
                )
                return
        elif path == "api/projects/p/nodes":
            text = query.get("q", [""])[0].casefold()
            sheet = query.get("sheet", [""])[0]
            matches = [
                n
                for n in self.nodes
                if (not sheet or n["sheet"] == sheet)
                and (not text or text in n["id"].casefold())
            ]
            data = {
                "nodes": matches,
                "pagination": {"offset": 0, "limit": 200, "total": len(matches)},
            }
        elif path == "api/projects/p/tasks":
            body = route.request.post_data_json
            assert body["operation"] == "evaluate", (
                "Navigation must not start global/AI work"
            )
            assert body["force"] is True, "Access requests a new targeted calculation"
            self.evaluations.append(body)
            task = {
                "id": f"evaluate-{len(self.evaluations)}",
                "operation": "evaluate",
                "nodeId": body["nodeId"],
                "status": "queued",
                "createdAt": f"2026-10-12T10:00:{len(self.evaluations):02}Z",
                "resultAvailable": False,
            }
            self.tasks[task["id"]] = task
            data = {"task": task}
            if body["nodeId"] in self.held_evaluations:
                self.held.append((route, data, "evaluate", body["nodeId"]))
                return
        elif path == "api/tasks/saved-doc":
            data = {
                "task": {
                    "id": "saved-doc",
                    "operation": "document",
                    "nodeId": "Calc!B1",
                    "status": "succeeded",
                    "resultAvailable": True,
                    "result": {"markdown": "Stored node explanation", "language": "fr"},
                }
            }
            if self.hold_saved_result:
                self.held.append((route, data, "saved_result", "Calc!B1"))
                return
            if self.reject_saved_result:
                route.fulfill(
                    status=503, json={"error": {"message": "Saved result unavailable"}}
                )
                return
        elif path.startswith("api/tasks/evaluate-"):
            task = self.tasks[path.rsplit("/", 1)[1]]
            if self.finish_calculations:
                task = {
                    **task,
                    "status": "succeeded",
                    "resultAvailable": True,
                    "result": {
                        "status": "completed",
                        "nodeId": task["nodeId"],
                        "value": 12,
                        "steps": [
                            {
                                "nodeId": task["nodeId"],
                                "evaluationStatus": "calculated",
                                "calculatedValue": 12,
                                "formula": "=A1+1",
                            }
                        ],
                    },
                }
                self.tasks[task["id"]] = task
            data = {"task": task}
        else:
            raise AssertionError(path)
        route.fulfill(json=data)

    def release(self, kind, node_id):
        index = next(
            i
            for i, (_, _, k, node) in enumerate(self.held)
            if k == kind and node == node_id
        )
        route, data, _, _ = self.held.pop(index)
        route.fulfill(json=data)

    def reject(self, kind, node_id, message):
        index = next(
            i
            for i, (_, _, k, node) in enumerate(self.held)
            if k == kind and node == node_id
        )
        route, _, _, _ = self.held.pop(index)
        route.fulfill(status=503, json={"error": {"message": message}})

    def tap(self, node_id):
        canvas = self.page.locator("#graph-canvas")
        canvas.scroll_into_view_if_needed()
        self.page.wait_for_function(
            "!document.querySelector('#graph-canvas')._cy.animated()"
        )
        # ResizeObserver and Cytoscape paint run after tab/scroll layout. A
        # physical click targets the painted card, not pre-paint coordinates.
        self.page.evaluate(
            "() => new Promise(resolve => requestAnimationFrame(() => "
            "requestAnimationFrame(resolve)))"
        )
        point = canvas.evaluate(
            """(element, id) => {
                const p=element._cy.getElementById(id).renderedPosition();
                const b=element.getBoundingClientRect();
                return {x:b.left+element.clientLeft+p.x,
                    y:b.top+element.clientTop+p.y};
            }""",
            node_id,
        )
        assert self.page.evaluate(
            """p=>document.elementFromPoint(p.x,p.y)?.closest('#graph-canvas')
                ===document.querySelector('#graph-canvas')""",
            point,
        ), "The painted card must be inside the visible canvas before clicking"
        self.page.mouse.click(point["x"], point["y"])

    def remember_canvas(self):
        self.page.locator("#graph-zoom-in").click()
        self.page.locator("#graph-canvas").focus()
        self.page.keyboard.press("ArrowRight")
        return self.page.locator("#graph-canvas").evaluate(
            """element => {
                const cy=element._cy; window.navigationCanvas=cy;
                return {zoom:cy.zoom(),pan:cy.pan(),
                    positions:cy.nodes().map(n=>({id:n.id(),position:n.position()}))};
            }"""
        )

    def assert_canvas(self, remembered):
        canvas = self.page.locator("#graph-canvas")
        assert canvas.evaluate("e=>e._cy===window.navigationCanvas")
        assert (
            canvas.evaluate(
                """e=>({zoom:e._cy.zoom(),pan:e._cy.pan(),
                positions:e._cy.nodes().map(n=>({id:n.id(),position:n.position()}))})"""
            )
            == remembered
        )

    def assert_bounded(self):
        for _, path, query in self.requests:
            assert not path.endswith(("/export", "/graph-view", "/graph"))
            if path.endswith(("/neighborhood", "/nodes")):
                assert int(query["limit"][0]) <= 200
        assert not self.errors


@pytest.fixture(params=[(1440, 900), (390, 844)])
def navigation(browser, request):
    page = browser.new_page(
        viewport={"width": request.param[0], "height": request.param[1]}
    )
    page.set_default_timeout(5000)
    # A persisted preference from the old interface cannot disable access-triggered
    # calculation after migration to the always-lazy policy.
    page.add_init_script("localStorage.setItem('linexcel-manual-recalculation','true')")
    nodes = [
        {
            "id": "Source!A1",
            "sheet": "Source",
            "cell": "A1",
            "kind": "input",
            "value": 10,
            "valueSource": "file",
        },
        *[
            {
                "id": f"Calc!B{i}",
                "sheet": "Calc",
                "cell": f"B{i}",
                "kind": "cell",
                "formula": "='Source'!A1+1" if i == 1 else f"=B{i - 1}+1",
                "cachedValue": 10 + i,
                "valueSource": "file",
            }
            for i in range(1, 4)
        ],
    ]
    app = Navigation(page, nodes)
    page.on("pageerror", lambda error: app.errors.append(str(error)))
    page.route("**/*", app.route)
    page.goto("http://localhost/mounted/")
    page.locator("[data-project='p']").click()
    page.locator("#node-list [data-node='Source!A1']").click()
    playwright.expect(page.locator("#node-detail h2")).to_contain_text("A1")
    page.locator("#tab-graph").click()
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._cy?.nodes().length===4"
    )
    yield app
    for route, data, _, _ in app.held:
        try:
            route.fulfill(json=data)
        except playwright.Error:
            pass
    page.close()


def test_old_manual_preference_cannot_block_calculation_behind_saved_reply(navigation):
    app = navigation
    app.hold_saved_result = True
    app.tap("Calc!B1")
    playwright.expect(
        app.page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("B1")
    # Pump browser events until the existing reply request is actually held.
    app.page.wait_for_function(
        "document.querySelector('#graph-preview .graph-reader-heading h3')"
        ".textContent.includes('B1')"
    )
    assert any(k == "saved_result" for _, _, k, _ in app.held)
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert [x["nodeId"] for x in app.evaluations] == ["Calc!B1"]
    app.release("saved_result", "Calc!B1")
    app.assert_bounded()


def test_first_graph_visit_calculates_its_initial_formula_with_old_preference(
    navigation,
):
    app = navigation
    app.page.locator("#tab-nodes").click()
    app.page.locator("[data-project='p']").click()
    assert app.evaluations == []
    app.hold_saved_result = True
    app.page.locator("#tab-graph").click()
    playwright.expect(
        app.page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("B1")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert [x["nodeId"] for x in app.evaluations] == ["Calc!B1"]
    # No direct cell action preceded this first Graph visit for the reopened
    # project. The stale preference and saved document fetch cannot block it.
    app.release("saved_result", "Calc!B1")
    app.assert_bounded()


def test_saved_reply_error_does_not_block_calculation_or_next_node(navigation):
    app = navigation
    app.reject_saved_result = True
    app.tap("Calc!B1")
    playwright.expect(
        app.page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("B1")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert app.evaluations[-1]["nodeId"] == "Calc!B1"
    playwright.expect(
        app.page.locator("#graph-preview .node-access-error")
    ).to_contain_text("Saved result unavailable")
    assert app.page.locator("#message").inner_text() == ""
    app.tap("Calc!B2")
    playwright.expect(
        app.page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("B2")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert [x["nodeId"] for x in app.evaluations] == ["Calc!B1", "Calc!B2"]
    app.assert_bounded()


def test_poll_recalculation_and_return_to_graph_keep_camera(navigation):
    app = navigation
    app.tap("Calc!B2")
    playwright.expect(
        app.page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("B2")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    remembered = app.remember_canvas()
    app.finish_calculations = True
    playwright.expect(app.page.locator("#graph-calculation-steps")).to_contain_text(
        "12", timeout=7000
    )
    app.assert_canvas(remembered)
    app.page.locator("#tab-tasks").click()
    app.page.locator("#tab-graph").click()
    app.assert_canvas(remembered)
    app.assert_bounded()


def test_known_node_opens_while_its_neighbors_wait_and_latest_click_wins(navigation):
    app = navigation
    app.delayed_nodes.update({"Calc!B1", "Calc!B2"})
    app.tap("Calc!B1")
    heading = app.page.locator("#graph-preview .graph-reader-heading h3")
    playwright.expect(heading).to_contain_text("B1")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert app.evaluations[-1]["nodeId"] == "Calc!B1"
    assert any(
        node == "Calc!B1" for _, _, kind, node in app.held if kind == "neighborhood"
    )
    app.tap("Calc!B2")
    playwright.expect(heading).to_contain_text("B2")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert app.evaluations[-1]["nodeId"] == "Calc!B2"
    app.tap("Calc!B3")
    playwright.expect(heading).to_contain_text("B3")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    app.release("neighborhood", "Calc!B2")
    app.release("neighborhood", "Calc!B1")
    playwright.expect(heading).to_contain_text("B3")
    assert [x["nodeId"] for x in app.evaluations] == ["Calc!B1", "Calc!B2", "Calc!B3"]
    app.assert_bounded()


def test_neighborhood_error_keeps_known_cell_and_navigation_available(navigation):
    app = navigation
    app.rejected_nodes.add("Calc!B2")
    remembered = app.remember_canvas()
    app.tap("Calc!B2")
    heading = app.page.locator("#graph-preview .graph-reader-heading h3")
    playwright.expect(heading).to_contain_text("B2")
    # The local graph notice accounts for the failed fetch, while the known
    # source/cache and calculation remain available in the new cell's reader.
    playwright.expect(app.page.locator("#view-graph")).to_contain_text(
        "Links unavailable"
    )
    assert "12" in app.page.locator("#graph-panel-calculation").inner_text()
    app.assert_canvas(remembered)
    app.tap("Calc!B3")
    playwright.expect(heading).to_contain_text("B3")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert app.evaluations[-1]["nodeId"] == "Calc!B3"
    app.assert_bounded()


def test_history_restores_previous_camera_and_recalculates_access(navigation):
    app = navigation
    app.tap("Calc!B2")
    heading = app.page.locator("#graph-preview .graph-reader-heading h3")
    playwright.expect(heading).to_contain_text("B2")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    remembered = app.remember_canvas()
    app.tap("Calc!B3")
    playwright.expect(heading).to_contain_text("B3")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    app.page.locator("#graph-zoom-out").click()
    app.page.locator("#graph-back").click()
    playwright.expect(heading).to_contain_text("B2")
    app.assert_canvas(remembered)
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert [x["nodeId"] for x in app.evaluations] == ["Calc!B2", "Calc!B3", "Calc!B2"]
    app.assert_bounded()


def test_slow_calculation_submission_does_not_reuse_old_result_or_block_navigation(
    navigation,
):
    app = navigation
    source = "='Source'!A1+1"
    app.nodes[1].update(sourceFormula=source)
    app.saved_evaluation = {
        "id": "old-completed-calculation",
        "operation": "evaluate",
        "nodeId": "Calc!B1",
        "status": "succeeded",
        "resultAvailable": True,
        "createdAt": "2026-10-10T10:00:00Z",
        "finishedAt": "2026-10-10T10:00:01Z",
        "result": {
            "status": "completed",
            "value": 999,
            "comparison": "different",
            "steps": [
                {
                    "nodeId": "Calc!B1",
                    "formula": source,
                    "evaluationStatus": "calculated",
                    "calculatedValue": 999,
                }
            ],
        },
    }
    # Reopen through the UI so the saved completed task is a genuine project
    # payload, not an injected browser variable or a synthetic render call.
    app.page.locator("#tab-nodes").click()
    app.page.locator("[data-project='p']").click()
    app.page.locator("#node-list [data-node='Source!A1']").click()
    app.page.locator("#tab-graph").click()
    app.page.wait_for_function(
        "document.querySelector('#graph-canvas')._cy?.nodes().length===4"
    )
    app.held_evaluations.add("Calc!B1")
    app.rejected_nodes.add("Calc!B1")
    app.tap("Calc!B1")
    reader = app.page.locator("#graph-panel-calculation")
    result_card = reader.locator(".value-card").nth(1)
    result = result_card.locator("strong")
    playwright.expect(
        app.page.locator("#graph-preview .graph-reader-heading .formula")
    ).to_have_text(source)
    playwright.expect(
        reader.locator(".value-card").first.locator("strong")
    ).to_have_text("11")
    playwright.expect(result).to_have_text("Envoi du recalcul…")
    playwright.expect(result_card).not_to_contain_text("Différent du cache")
    playwright.expect(app.page.locator("#graph-calculation-steps")).not_to_contain_text(
        "999"
    )
    # A failed relationship fetch has a different stage from the calculation
    # POST, which is still awaiting the server and must retain its pending label.
    playwright.expect(reader.locator(".node-access-error")).to_contain_text(
        "Links unavailable"
    )
    playwright.expect(result).to_have_text("Envoi du recalcul…")
    assert [x["nodeId"] for x in app.evaluations] == ["Calc!B1"]
    assert app.evaluations[0]["force"] is True
    assert any(kind == "evaluate" for _, _, kind, _ in app.held)
    app.reject("evaluate", "Calc!B1", "Compute submission unavailable")
    playwright.expect(result).to_have_text("Recalcul non abouti")
    playwright.expect(result_card).not_to_contain_text("Différent du cache")
    playwright.expect(reader).to_contain_text("Compute submission unavailable")
    playwright.expect(reader).to_contain_text("Links unavailable")
    playwright.expect(app.page.locator("#graph-calculation-steps")).not_to_contain_text(
        "999"
    )
    assert app.page.locator("#message").inner_text() == ""
    app.tap("Calc!B2")
    playwright.expect(
        app.page.locator("#graph-preview .graph-reader-heading h3")
    ).to_contain_text("B2")
    app.page.wait_for_function("document.querySelector('#graph-recalculate').disabled")
    assert [x["nodeId"] for x in app.evaluations] == ["Calc!B1", "Calc!B2"]
    app.assert_bounded()
