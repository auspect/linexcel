"""Browser contract checks for the mounted, lazy web application."""

import json
import os
from pathlib import Path

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


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_lazy_selection_mounted_api_and_tab_boundaries(browser, viewport):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    project = {"id": "p1", "name": "Budget <2026>", "revision": 1}
    graph = {
        "nodes": [
            {
                "id": "Budget!A1",
                "sheet": "Budget",
                "cell": "A1",
                "formula": "=1+2",
                "cachedValue": 3,
                "valueSource": "excel-cache",
                "kind": "formula",
            }
        ],
        "edges": [{"source": "Budget!Z9", "target": "Budget!A1"}],
        "meta": {"sheets": ["Budget", "Feuille vide"]},
    }
    graph["nodes"] = (
        [
            {"id": f"Budget!B{i}", "sheet": "Budget", "cell": f"B{i}", "value": "Titre"}
            for i in range(20)
        ]
        + graph["nodes"]
        + [{"id": "Budget!Z9", "sheet": "Budget", "cell": "Z9", "value": 1}]
    )
    tasks = []
    operations = []
    requests = []

    def handle(route):
        if route.request.url.endswith("assets/cytoscape.min.js"):
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
            return
        path = route.request.url.split("/mounted/", 1)[-1]
        requests.append(path)
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/config":
            body = {"limits": {"memoryMb": 1024, "uploadMb": 64, "retentionDays": 7}}
        elif path == "api/projects":
            body = {"projects": [project]}
        elif path == "api/projects/p1":
            body = {"project": project, "graph": graph, "tasks": tasks}
        elif path.startswith("api/tasks/"):
            task = next(t for t in tasks if t["id"] == path.rsplit("/", 1)[-1])
            task["status"] = "succeeded"
            task["result"] = {
                "status": "unsupported",
                "diagnostics": ["Référence externe non résolue"],
            }
            body = {"task": task}
        elif path == "api/projects/p1/tasks":
            request = route.request.post_data_json
            operations.append(request)
            task = {
                "id": str(len(tasks)),
                "operation": request["operation"],
                "nodeId": request.get("nodeId"),
                "status": "succeeded",
                "result": {"notice": "Rendu LibreOffice : capture déterministe"}
                if request["operation"] == "capture"
                else {
                    "status": "unsupported",
                    "diagnostics": ["Référence externe non résolue"],
                },
            }
            if not tasks and request["operation"] == "evaluate":
                task["status"] = "running"
                task.pop("result")
            tasks.append(task)
            body = {"task": task}
        else:
            route.fulfill(status=404, body='{"message":"Route inconnue"}')
            return
        route.fulfill(body=json.dumps(body), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        page.locator("[data-project='p1']").click()
        page.locator("[data-node]").first.wait_for()
        assert operations == []
        assert page.locator("#memory").get_attribute("max") == "1024"
        page.locator("[data-node='Budget!A1']").click()
        page.locator("#node-detail").get_by_text(
            "Non pris en charge", exact=False
        ).wait_for()
        assert len(operations) == 1
        # Poll only the active task, not the potentially large project graph.
        assert "api/tasks/0" in requests
        assert requests.count("api/projects/p1") == 1
        assert operations[0]["budget"] == {"seconds": 120, "memoryMb": 512}
        page.locator("[data-node='Budget!A1']").click()
        page.wait_for_function(
            "document.querySelector('#recalculate').disabled === false"
        )
        # The server owns version-aware cache reuse, even for a known node.
        assert len(operations) == 2
        assert operations[-1]["force"] is False
        assert operations[0]["force"] is False
        assert "3" in page.locator("#node-detail").inner_text()
        for name in ["nodes", "graph", "sheets", "captures", "tasks", "diagnostics"]:
            page.locator(f"[data-tab='{name}']").click()
            assert page.locator(f"#view-{name}").is_visible()
            assert (
                page.locator("#workspace-content > [role=tabpanel]:visible").count()
                == 1
            )
            assert page.evaluate(
                "document.documentElement.scrollWidth <= window.innerWidth"
            )
        assert len(operations) == 2  # Visiting captures and diagnostics is inert.
        assert "Feuille vide" in page.locator("#sheet-list").inner_text()
        page.locator("[data-tab='graph']").click()
        page.locator("#graph-list").evaluate("element => element.open = true")
        assert page.locator("[data-graph-node='Budget!A1']").is_visible()
        assert page.locator("[data-graph-node='Budget!Z9']").is_visible()
        assert page.locator("[data-graph-node]").count() == 2
        page.locator("[data-graph-node='Budget!Z9']").click()
        assert page.locator("#view-graph").is_visible()
        assert "Budget!Z9" in page.locator("#graph-notice").inner_text()
        page.locator("#graph-back").click()
        assert "Budget!A1" in page.locator("#graph-notice").inner_text()
        page.locator("#graph-search").fill("Z9")
        page.locator("#graph-results [data-related='Budget!Z9']").click()
        assert page.locator("#view-graph").is_visible()
        page.locator("#graph-back").click()
        page.locator("[data-tab='nodes']").click()
        page.locator("#recalculate").click()
        page.wait_for_function(
            "document.querySelector('#recalculate').disabled === false"
        )
        assert operations[-1]["force"] is True
        page.locator("[data-tab='captures']").click()
        page.locator("#capture").click()
        notice = page.locator("#capture-results .notice")
        notice.wait_for()
        assert "Rendu LibreOffice" in notice.inner_text()
        page.locator("[data-tab='nodes']").click()
        assert not notice.is_visible()
        assert errors == []
    finally:
        page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_upload_file_validation_and_accessible_removal(browser, viewport):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})

    def handle(route):
        if route.request.url.endswith("assets/cytoscape.min.js"):
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
            return
        if route.request.url.endswith("/mounted/"):
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
        else:
            route.fulfill(
                body=json.dumps({"projects": [], "limits": {"uploadMb": 64}}),
                content_type="application/json",
            )

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        page.locator("#workbook").set_input_files(
            {"name": "notes.txt", "mimeType": "text/plain", "buffer": b"hello"}
        )
        assert "Format non pris" in page.locator("#message").inner_text()
        page.locator("#workbook").set_input_files(
            {
                "name": "budget.xlsx",
                "mimeType": "application/octet-stream",
                "buffer": b"workbook",
            }
        )
        assert "budget.xlsx" in page.locator("#workbook-files").inner_text()
        assert not page.locator("#message").inner_text()
        page.locator("#references").set_input_files(
            {
                "name": "BUDGET.xlsx",
                "mimeType": "application/octet-stream",
                "buffer": b"duplicate",
            }
        )
        assert "double" in page.locator("#message").inner_text()
        assert not page.locator("#references-files").inner_text()
        page.locator("#references").set_input_files(
            {
                "name": "source.xlsx",
                "mimeType": "application/octet-stream",
                "buffer": b"source",
            }
        )
        assert "source.xlsx" in page.locator("#references-files").inner_text()
        remove = page.get_by_role("button", name="Retirer source.xlsx")
        remove.focus()
        page.keyboard.press("Enter")
        assert not page.locator("#references-files").inner_text()
        # Native input is keyboard accessible; drop uses the same validation.
        page.locator('[data-drop="references"]').evaluate("""element => {
            const data = new DataTransfer();
            data.items.add(new File(['external'], 'external.xlsm'));
            element.dispatchEvent(new DragEvent('drop', {
                bubbles: true, dataTransfer: data
            }));
        }""")
        assert "external.xlsm" in page.locator("#references-files").inner_text()
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
    finally:
        page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_compact_range_membership_navigation(browser, viewport):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    project = {"id": "ranges", "name": "Plages compactes"}
    graph = {
        "nodes": [
            {
                "id": "Data!E4",
                "sheet": "Data",
                "cell": "E4",
                "kind": "input",
                "row": 4,
                "column": 5,
                "value": 8,
            },
            {
                "id": "Data!E1:E1000000",
                "sheet": "Data",
                "reference": "Data!E1:E1000000",
                "kind": "range",
                "bounds": [1, 5, 1000000, 5],
            },
            {
                "id": "Summary!C5",
                "sheet": "Summary",
                "cell": "C5",
                "kind": "cell",
                "formula": "=SUM(Data!E1:E1000000)",
            },
        ],
        "edges": [{"source": "Data!E1:E1000000", "target": "Summary!C5"}],
        "meta": {"sheets": ["Data", "Summary"]},
    }

    def handle(route):
        if route.request.url.endswith("assets/cytoscape.min.js"):
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
            return
        url = route.request.url
        if url.endswith("/mounted/"):
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if url.endswith("/api/projects"):
            data = {"projects": [project]}
        elif url.endswith("/api/projects/ranges"):
            data = {"project": project, "graph": graph, "tasks": []}
        elif url.endswith("/tasks"):
            data = {
                "task": {
                    "id": "t",
                    "operation": "evaluate",
                    "nodeId": "Summary!C5",
                    "status": "succeeded",
                    "result": {
                        "status": "completed",
                        "value": 8,
                        "steps": [
                            {
                                "node_id": "Summary!C5" if i == 1998 else f"Data!E{i}",
                                "formula": "=1+2" if i >= 1998 else None,
                                "evaluationStatus": "calculated"
                                if i >= 1998
                                else "input",
                                "calculatedValue": 3 if i >= 1998 else None,
                                "calculatedValueKind": "scalar",
                                "cachedValue": 1,
                                "comparison": "different" if i >= 1998 else "unknown",
                                "dependencies": [],
                            }
                            for i in range(2000)
                        ],
                    },
                }
            }
        else:
            data = {}
        route.fulfill(body=json.dumps(data), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        page.locator('[data-project="ranges"]').click()
        page.locator('[data-node="Data!E4"]').click()
        assert "incluse dans cette plage" in page.locator("#node-detail").inner_text()
        page.locator('[data-tab="graph"]').click()
        page.wait_for_function("document.querySelector('#graph-canvas')._cy != null")
        canvas = page.locator("#graph-canvas")
        canvas.scroll_into_view_if_needed()
        initial_zoom = canvas.evaluate("element => element._cy.zoom()")
        page.locator("#graph-zoom-in").click()
        assert canvas.evaluate("element => element._cy.zoom()") > initial_zoom
        page.locator("#graph-zoom-out").click()
        before_pan = canvas.evaluate("element => element._cy.pan()")
        canvas.focus()
        page.keyboard.press("ArrowRight")
        assert canvas.evaluate("element => element._cy.pan().x") < before_pan["x"]
        canvas.scroll_into_view_if_needed()
        box = canvas.bounding_box()
        before_pan = canvas.evaluate("element => element._cy.pan()")
        page.mouse.move(box["x"] + 10, box["y"] + 10)
        page.mouse.down()
        page.mouse.move(box["x"] + 45, box["y"] + 35, steps=5)
        page.mouse.up()
        assert canvas.evaluate("element => element._cy.pan()") != before_pan
        page.locator("#graph-fit").click()
        page.locator("#graph-depth").select_option("2")
        assert canvas.evaluate("element => element._cy.nodes().length") == 3
        page.locator("#graph-sheet").select_option("Data")
        assert canvas.evaluate("element => element._cy.nodes().length") == 3
        assert "Data · E4" in page.locator("#graph-preview").inner_text()
        page.locator("#graph-sheet").select_option("")
        page.locator("#graph-depth").select_option("1")
        page.locator("#graph-reset").click()
        canvas.scroll_into_view_if_needed()
        box = canvas.bounding_box()
        position = canvas.evaluate(
            "element => element._cy.getElementById('Data!E1:E1000000')"
            ".renderedPosition()"
        )
        page.mouse.click(box["x"] + position["x"], box["y"] + position["y"])
        page.wait_for_function(
            "document.querySelector('#graph-preview').textContent.includes('E1:E1000000')"
        )
        page.locator("#graph-list").evaluate("element => element.open = true")
        assert page.locator("#view-graph").is_visible()
        assert (
            "E1:E1000000"
            in page.locator('[data-graph-node="Data!E1:E1000000"]').inner_text()
        )
        assert page.locator('[data-graph-node="Data!E4"]').is_visible()
        assert page.locator('[data-graph-node="Summary!C5"]').is_visible()
        assert page.locator("[data-graph-node]").count() == 3
        note = page.locator("#graph .range-membership").inner_text()
        assert "1 cellules représentées" in note
        assert "1\u202f000\u202f000 positions" in note
        assert "1 relations" in page.locator("#stats").inner_text()
        assert (
            "Summary · C5"
            in page.locator('[data-graph-node="Summary!C5"]').inner_text()
        )
        page.locator('[data-graph-node="Summary!C5"]').click()
        page.wait_for_function(
            "document.querySelectorAll('#calculation-steps .step-card').length === 100"
        )
        assert "100 / 2000" in page.locator("#calculation-steps").inner_text()
        assert "Valeur moteur" in page.locator("#calculation-steps").inner_text()
        assert "Différent du cache" in page.locator("#calculation-steps").inner_text()
        assert page.locator("#graph-recalculate").is_enabled()
        assert "Valeur de dépendance" not in page.locator("#graph-preview").inner_text()
        assert (
            "Summary!C5"
            in page.locator("#calculation-steps .step-card").first.inner_text()
        )
        assert (
            "Data!E1999"
            in page.locator("#calculation-steps .step-card").nth(1).inner_text()
        )
        assert "omis" in page.locator("#node-detail .diagnostic pre").inner_text()
        assert page.locator("#download-node-result").count() == 1
        assert page.locator("#view-graph").is_visible()
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
    finally:
        page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_volatile_snapshot_scoping_and_empty_project_graph(browser, viewport):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    projects = [{"id": key, "name": f"Project {key}"} for key in ("a", "b", "c")]
    graph = {
        "nodes": [
            {
                "id": "Volatile!A1",
                "sheet": "Volatile",
                "cell": "A1",
                "kind": "cell",
                "formula": "=RAND()",
                "cachedValue": 0.1,
            }
        ],
        "edges": [],
        "meta": {"sheets": ["Volatile"]},
    }
    graph["nodes"].extend(
        {
            "id": f"Volatile!C{i}",
            "sheet": "Volatile",
            "cell": f"C{i}",
            "kind": "input",
            "value": i,
        }
        for i in range(24)
    )
    graph["edges"] = [
        {"source": f"Volatile!C{i}", "target": "Volatile!A1"} for i in range(24)
    ]
    task = {
        "id": "volatile",
        "operation": "evaluate",
        "nodeId": "Volatile!A1",
        "status": "succeeded",
        "finishedAt": 1780000000,
        "result": {
            "status": "completed",
            "value": 0.2,
            "comparison": "different",
            "volatile": True,
            "steps": [
                {
                    "nodeId": "Volatile!A1",
                    "formula": "=RAND()",
                    "evaluationStatus": "calculated",
                    "calculatedValue": 0.2,
                    "cachedValue": 0.1,
                    "comparison": "different",
                }
            ],
        },
    }

    def handle(route):
        path = route.request.url.split("/mounted/", 1)[-1]
        if path == "assets/cytoscape.min.js":
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
            return
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/projects":
            body = {"projects": projects}
        elif path.endswith("/tasks"):
            body = {"task": task}
        elif path.startswith("api/projects/"):
            key = path.rsplit("/", 1)[-1]
            body = {
                "project": next(p for p in projects if p["id"] == key),
                "graph": graph
                if key == "a"
                else None
                if key == "b"
                else {"nodes": [], "edges": []},
                "tasks": [{**task, "id": "parent", "nodeId": "Volatile!B1"}, task]
                if key == "a"
                else [],
            }
        else:
            body = {}
        route.fulfill(body=json.dumps(body), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        for destination in ("b", "c"):
            page.locator('[data-project="a"]').click()
            page.locator('[data-tab="nodes"]').click()
            page.locator('[data-node="Volatile!A1"]').click()
            page.locator("#calculation-steps .volatile-operation").wait_for()
            steps = page.locator("#calculation-steps").inner_text()
            assert "comparaisons au cache ne sont pas interprétées" in steps
            assert "Différent du cache" not in steps
            assert "Identique au cache" not in steps
            assert "0.2" in steps
            page.locator('[data-tab="graph"]').click()
            assert page.locator("#graph-preview .volatile-operation").is_visible()
            assert "volatile!a1" in page.locator("#graph-preview").inner_text().lower()
            assert page.locator("#graph-canvas").evaluate("element => !!element._cy")
            assert page.locator("#graph-canvas").evaluate(
                """element => {
                    const boxes = element._cy.nodes().map(node => node.boundingBox());
                    return boxes.every((a, i) => boxes.slice(i + 1).every(b =>
                        a.x2 < b.x1 || b.x2 < a.x1 || a.y2 < b.y1 || b.y2 < a.y1));
                }"""
            )
            assert (
                page.locator("#graph-canvas").evaluate("element => element._cy.zoom()")
                >= 0.65
            )
            page.locator("#graph-fit").click()
            assert (
                page.locator("#graph-canvas").evaluate("element => element._cy.zoom()")
                < 0.65
            )
            page.locator("#graph-reset").click()
            assert (
                page.locator("#graph-canvas").evaluate("element => element._cy.zoom()")
                >= 0.65
            )
            page.locator(f'[data-project="{destination}"]').click()
            page.wait_for_function(
                "name => document.querySelector('#project-title').textContent === name",
                arg=f"Project {destination}",
            )
            assert page.locator("#view-graph").is_visible()
            assert page.locator("#graph-preview").inner_text() == ""
            assert not page.locator("#graph-canvas").evaluate(
                "element => !!element._cy"
            )
            assert page.locator("#graph-canvas canvas").count() == 0
            assert page.locator("#graph-fit").is_disabled()
            assert "Volatile!A1" not in page.locator("#view-graph").inner_text()
    finally:
        page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_gallery_opt_in_ai_safe_previews_patterns_and_sheet_focus(browser, viewport):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    project = {"id": "modern", "name": "Modern workflows"}
    png = (
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
        "AAAAC0lEQVR42mP8/x8AAwMCAO+aE1sAAAAASUVORK5CYII="
    )
    capture = {
        "id": "capture-one",
        "operation": "capture",
        "status": "succeeded",
        "finishedAt": 1780000000,
        "result": {
            "notice": "Rendu LibreOffice",
            "screenshots": [
                {"sheet": "Deals", "data": png},
                {"sheet": "Fuel", "data": png},
            ],
        },
    }
    graph = {
        "nodes": [
            {
                "id": "Deals!A1",
                "sheet": "Deals",
                "cell": "A1",
                "kind": "cell",
                "formula": "=1+1",
                "patternId": "p",
                "patternMemberCount": 2,
            },
            {
                "id": "Deals!A2",
                "sheet": "Deals",
                "cell": "A2",
                "kind": "cell",
                "formula": "=1+1",
                "patternId": "p",
                "patternMemberCount": 2,
            },
            {
                "id": "Fuel!B1",
                "sheet": "Fuel",
                "cell": "B1",
                "kind": "cell",
                "formula": "=2+2",
            },
        ],
        "edges": [],
        "formulaPatterns": [
            {
                "id": "p",
                "sheet": "Deals",
                "memberIds": ["Deals!A1", "Deals!A2"],
                "memberCount": 2,
                "ranges": ["A1:A2"],
                "signature": "=1+1",
            }
        ],
        "meta": {
            "sheets": ["Deals", "Fuel"],
            "untrusted": "<img src=x onerror=window.bad=1>",
            "data": "data:image/png;base64," + "A" * 200000,
        },
    }
    # Restored task files can arrive in UUID order, not chronological order.
    tasks = [
        capture,
        {
            **capture,
            "id": "older-capture",
            "finishedAt": 1770000000,
            "result": {"screenshots": [{"sheet": "Obsolete", "data": png}]},
        },
    ]
    requests = []
    markdown = (
        "# Explication\n\n**Valeur** et `=1+1`\n\n"
        "| Source | Valeur |\n|---|---|\n| A1 | 2 |\n\n"
        "<img src=x onerror=window.bad=1>\n<script>window.bad=1</script>"
    )

    def handle(route):
        path = route.request.url.split("/mounted/", 1)[-1]
        if path == "assets/cytoscape.min.js":
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
            return
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"),
                content_type="text/html",
                headers={
                    "Content-Security-Policy": (
                        "default-src 'self'; script-src 'self' 'unsafe-inline'; "
                        "style-src 'self' 'unsafe-inline'; "
                        "img-src 'self' data: blob:; connect-src 'self'"
                    )
                },
            )
            return
        if path == "api/config":
            data = {
                "ai": {
                    "available": True,
                    "model": "local-doc",
                    "visionModel": "local-vision",
                    "languages": ["fr", "en"],
                    "tokenBudget": 1000,
                }
            }
        elif path == "api/projects":
            data = {"projects": [project]}
        elif path == "api/projects/modern":
            data = {"project": project, "graph": graph, "tasks": tasks}
        elif path == "api/projects/modern/tasks":
            body = route.request.post_data_json
            requests.append(body)
            task = {
                "id": f"task-{len(requests)}",
                "operation": body["operation"],
                "nodeId": body.get("nodeId"),
                "options": body.get("options", {}),
                "status": "succeeded",
            }
            if body["operation"] == "evaluate":
                task["result"] = {"status": "completed", "value": 2}
            else:
                task["result"] = {
                    "markdown": markdown,
                    "model": "local-doc",
                    "language": body["options"]["language"],
                    "usage": {"totalTokens": 80},
                    "generatedAt": 1780000000,
                    "sources": body.get("options", {})
                    if body["operation"] == "describe_capture"
                    else {
                        "calculationTaskId": next(
                            (
                                t["id"]
                                for t in reversed(tasks)
                                if t["operation"] == "evaluate"
                                and t.get("nodeId") == body.get("nodeId")
                            ),
                            None,
                        )
                    },
                }
                if body["force"]:
                    task["status"] = "running"
                    task.pop("result")
            tasks.append(task)
            data = {"task": task}
        elif path.endswith("/cancel"):
            task = next(t for t in tasks if t["id"] == path.split("/")[-2])
            task["status"] = "cancelled"
            data = {"task": task}
        elif path.startswith("api/tasks/"):
            data = {
                "task": next(t for t in tasks if t["id"] == path.rsplit("/", 1)[-1])
            }
        else:
            data = {}
        route.fulfill(body=json.dumps(data), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        picker = page.locator('[data-file-pick="workbook"]')
        with page.expect_file_chooser() as event:
            picker.press("Enter")
        event.value.set_files(
            {
                "name": "keyboard.xlsx",
                "mimeType": "application/octet-stream",
                "buffer": b"file",
            }
        )
        assert "keyboard.xlsx" in page.locator("#workbook-files").inner_text()
        page.locator('[data-project="modern"]').click()
        page.locator('[data-node="Deals!A1"]').click()
        page.locator("#node-ai [data-ai-start]").wait_for()
        assert not [r for r in requests if r["operation"] != "evaluate"]
        assert "2 cellules" in page.locator(".pattern-card").inner_text()
        page.locator(".pattern-card summary").click()
        page.locator('.pattern-card [data-related="Deals!A2"]').click()
        page.locator("#node-ai [data-ai-start]").click()
        page.locator("#node-ai .ai-result").wait_for()
        assert requests[-1]["operation"] == "document"
        assert requests[-1]["nodeId"] == "Deals!A2"
        assert requests[-1]["budget"]["tokens"] == 1000
        # Documents that predate any recalculation retain that source limitation.
        doc_task = next(t for t in reversed(tasks) if t["operation"] == "document")
        saved_calculation = doc_task["result"]["sources"]["calculationTaskId"]
        doc_task["result"]["sources"]["calculationTaskId"] = None
        page.locator('[data-project="modern"]').click()
        page.locator('[data-node="Deals!A2"]').click()
        page.locator("#node-ai").get_by_text(
            "Sources : formules et valeurs enregistrées, sans instantané de recalcul.",
            exact=True,
        ).wait_for()
        page.locator("#node-ai").get_by_text(
            "Un recalcul est maintenant disponible.", exact=False
        ).wait_for()
        doc_task["result"]["sources"]["calculationTaskId"] = saved_calculation
        page.locator('[data-project="modern"]').click()
        page.locator('[data-node="Deals!A2"]').click()
        assert page.locator("#node-ai .ai-result table").count() == 1
        assert (
            page.locator("#node-ai .ai-result img, #node-ai .ai-result script").count()
            == 0
        )
        assert page.evaluate("window.bad") is None
        ai_count = len([r for r in requests if r["operation"] == "document"])
        page.get_by_text("Documentation IA à la demande", exact=True).click()
        page.locator("#ai-language").select_option("en")
        assert page.locator("#node-ai .ai-result").count() == 0
        assert len([r for r in requests if r["operation"] == "document"]) == ai_count
        page.locator("#ai-language").select_option("fr")
        assert page.locator("#node-ai .ai-result").count() == 1
        page.locator("#recalculate").click()
        page.locator("#node-ai").get_by_text(
            "Un recalcul plus récent est disponible.", exact=False
        ).wait_for()
        page.locator("#node-ai [data-ai-start]").click()
        page.locator("#node-ai [data-ai-cancel]").click()
        page.locator("#node-ai").get_by_text(
            "Génération annulée.", exact=True
        ).wait_for()
        page.locator('[data-tab="graph"]').click()
        page.locator("#graph-sheet").select_option("Fuel")
        page.wait_for_function(
            "document.querySelector('#graph-preview').textContent.includes('Fuel · B1')"
        )
        assert page.locator("#graph-canvas").evaluate("el => el._cy.zoom()") <= 1.25
        page.locator("#graph-search").fill("Deals!A1")
        page.locator('#graph-results [data-related="Deals!A1"]').click()
        assert page.locator("#graph-sheet").input_value() == "Deals"
        page.locator('[data-tab="captures"]').click()
        assert page.locator(".capture-thumb").count() == 2
        assert page.locator("#capture-selected figure img").count() == 1
        page.wait_for_function(
            "document.querySelector('#capture-selected figure img').naturalWidth > 0"
        )
        page.locator("#capture-next").click()
        assert "Fuel" in page.locator("#capture-selected").inner_text()
        assert not [r for r in requests if r["operation"] == "describe_capture"]
        page.add_style_tag(
            content=(
                "#capture-selected figure{max-height:100px}"
                "#capture-selected figure img{min-height:1000px}"
                ".capture-gallery{max-height:80px}"
            )
        )
        page.locator("#capture-selected figure").evaluate("el => el.scrollTop = 250")
        page.locator(".capture-gallery").evaluate("el => el.scrollTop = 30")
        page.locator("#capture-ai [data-ai-start]").click()
        page.locator("#capture-ai .ai-result").wait_for()
        page.wait_for_function(
            "document.querySelector('#capture-selected figure').scrollTop >= 240"
        )
        assert page.locator(".capture-gallery").evaluate("el => el.scrollTop") == 30
        assert requests[-1]["operation"] == "describe_capture"
        assert requests[-1]["options"]["captureIndex"] == 1
        assert requests[-1]["options"]["captureTaskId"] == "capture-one"
        page.locator('[data-tab="sheets"]').click()
        page.locator("#workbook-ai [data-ai-start]").click()
        page.locator("#workbook-ai .ai-result").wait_for()
        assert requests[-1]["nodeId"] is None
        assert (
            "Sources : formules, valeurs enregistrées"
            in page.locator("#workbook-ai").inner_text()
        )
        assert (
            "Des calculs ciblés plus récents"
            in page.locator("#workbook-ai").inner_text()
        )
        page.locator('[data-tab="tasks"]').click()
        assert page.locator("#task-list img").count() == 0
        assert "data:image/png;base64" not in page.locator("#task-list").inner_text()
        assert page.locator("#task-list .json-key").count() > 0
        page.locator('[data-tab="diagnostics"]').click()
        assert page.locator("#diagnostics img").count() == 0
        assert len(page.locator("#diagnostics").inner_text()) < 10000
        assert page.evaluate("window.bad") is None
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
        assert errors == []
    finally:
        page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_graph_reader_calculation_first_and_long_document_state(browser, viewport):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    errors, requests = [], []
    page.on("pageerror", lambda error: errors.append(str(error)))
    project = {"id": "reader", "name": "Reader regression"}
    graph = {
        "nodes": [
            {
                "id": "Sheet!B1",
                "sheet": "Sheet",
                "cell": "B1",
                "kind": "cell",
                "formula": "=A1+1",
                "cachedValue": 99,
                "valueSource": "saved_cache",
            },
            {
                "id": "Sheet!A1",
                "sheet": "Sheet",
                "cell": "A1",
                "kind": "input",
                "value": 2,
                "valueSource": "input",
            },
        ],
        "edges": [{"source": "Sheet!A1", "target": "Sheet!B1"}],
        "meta": {"sheets": ["Sheet"]},
    }
    result = {
        "status": "completed",
        "value": 3,
        "comparison": "different",
        "steps": [
            {
                "nodeId": "Sheet!B1",
                "formula": "=A1+1",
                "evaluationStatus": "calculated",
                "calculatedValue": 3,
                "cachedValue": 99,
                "comparison": "different",
                "dependencies": ["Sheet!A1"],
            },
            {
                "nodeId": "Sheet!A1",
                "evaluationStatus": "input",
                "cachedValue": 2,
                "valueSource": "input",
                "dependencies": [],
            },
        ],
    }
    tasks = [
        {
            "id": "calculated",
            "operation": "evaluate",
            "nodeId": "Sheet!B1",
            "status": "succeeded",
            "result": result,
        },
        {
            "id": "documented",
            "operation": "document",
            "nodeId": "Sheet!B1",
            "status": "succeeded",
            "options": {"language": "fr"},
            "result": {
                "markdown": "\n\n".join(
                    f"## Section {i}\n"
                    "A long explanation of the calculation and its evidence."
                    for i in range(80)
                ),
                "language": "fr",
                "model": "test",
                "usage": {"inputTokens": 20, "outputTokens": 40},
            },
        },
    ]

    def handle(route):
        path = route.request.url.split("/mounted/", 1)[-1].split("?", 1)[0]
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "assets/cytoscape.min.js":
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
            return
        if path == "api/projects":
            data = {"projects": [project]}
        elif path == "api/projects/reader":
            data = {"project": project, "graph": graph, "tasks": tasks}
        elif path == "api/projects/reader/tasks":
            requests.append(route.request.post_data_json)
            data = {"task": tasks[0]}
        else:
            data = {}
        route.fulfill(body=json.dumps(data), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/?lang=en")
        page.locator('[data-project="reader"]').click()
        page.locator("#tab-graph").click()
        calculation = page.locator("#graph-panel-calculation")
        assert calculation.is_visible()
        assert "AI" not in calculation.inner_text()
        assert "99" in calculation.inner_text() and "3" in calculation.inner_text()
        assert "Differs from the cache" in calculation.inner_text()
        assert "Sheet!A1" in page.locator("#graph-calculation-steps").inner_text()
        assert page.locator("#graph-calculation-steps .step-card").count() == 2
        assert "Input data" in page.locator("#graph-calculation-steps").inner_text()
        # UI labels translate, but an engine string that happens to match one
        # must remain byte-for-byte identical to the workbook result.
        result["steps"][0]["calculatedValue"] = "Formule"
        page.locator("#graph-recalculate").click()
        page.locator("#graph-calculation-steps [data-no-i18n]").get_by_text(
            "Formule", exact=True
        ).wait_for()
        camera = page.locator("#graph-canvas").evaluate(
            "el => ({zoom:el._cy.zoom(), pan:el._cy.pan()})"
        )
        size = page.locator("#graph-preview").bounding_box()
        page.locator("#graph-tab-calculation").focus()
        page.keyboard.press("ArrowRight")
        assert page.locator("#graph-tab-ai").get_attribute("aria-selected") == "true"
        reader = page.locator("#graph-panel-ai")
        assert reader.is_visible() and calculation.is_hidden()
        assert page.locator("#graph-preview").bounding_box()["height"] == size["height"]
        assert reader.evaluate("el => el.scrollHeight > el.clientHeight * 3")
        reader.evaluate("el => el.scrollTop = 500")
        page.locator("#graph-tab-ai").press("Home")
        page.locator("#graph-tab-calculation").press("End")
        assert reader.evaluate("el => el.scrollTop") == 500
        # A calculation task completes while the user is reading another panel.
        # Its normal render cycle must preserve the active tab, scroll and focus.
        page.locator("#graph-recalculate").evaluate("el => el.click()")
        page.wait_for_function(
            "document.querySelector('#graph-recalculate').disabled === false"
        )
        assert page.locator("#graph-tab-ai").get_attribute("aria-selected") == "true"
        assert reader.evaluate("el => el.scrollTop") == 500
        assert page.locator("#graph-tab-ai").evaluate(
            "el => el === document.activeElement"
        )
        assert (
            page.locator("#graph-canvas").evaluate(
                "el => ({zoom:el._cy.zoom(), pan:el._cy.pan()})"
            )
            == camera
        )
        if viewport[0] > 1100:
            before = page.locator("#graph-preview").bounding_box()["width"]
            page.locator("#graph-reader-width").focus()
            page.keyboard.press("ArrowRight")
            assert page.locator("#graph-preview").bounding_box()["width"] > before
            assert (
                page.locator("#graph-canvas").evaluate("el => el._cy.zoom()")
                == camera["zoom"]
            )
        assert requests and all(r["operation"] == "evaluate" for r in requests)
        page.locator("#graph-search").fill("Sheet!A1")
        page.locator('#graph-results [data-related="Sheet!A1"]').click()
        assert calculation.is_visible() and reader.is_hidden()
        assert (
            page.locator("#graph-tab-calculation").get_attribute("aria-selected")
            == "true"
        )
        assert calculation.evaluate("el => el.scrollTop") == 0
        assert "99" not in calculation.inner_text()
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
        assert errors == []
    finally:
        page.close()


@pytest.mark.parametrize("viewport", [(1440, 900), (390, 844)])
def test_english_ui_shared_graph_ai_usage_and_explicit_generation(browser, viewport):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    project = {"id": "english", "name": "Public demo"}
    graph = {
        "nodes": [
            {
                "id": "Fuel!B2",
                "sheet": "Fuel",
                "cell": "B2",
                "kind": "cell",
                "formula": "=2+2",
                "cachedValue": 4,
            }
        ],
        "edges": [],
        "meta": {"sheets": ["Fuel"]},
        "formulaPatterns": [],
    }
    graph["nodes"].append(
        {
            "id": "Feuilles!A1",
            "sheet": "Feuilles",
            "cell": "A1",
            "kind": "input",
            "value": "Formule",
        }
    )
    saved = {
        "id": "saved-fr",
        "operation": "document",
        "nodeId": "Fuel!B2",
        "status": "succeeded",
        "createdAt": 1780000000,
        "options": {"language": "fr"},
        "result": {
            "markdown": "**Documentation française conservée.**",
            "model": "local-doc",
            "language": "fr",
            "generatedAt": 1780000000,
            "usage": {
                "inputTokens": 10,
                "outputTokens": 5,
                "totalTokens": 15,
                "estimated": False,
            },
            "sources": {"calculationTaskId": None},
        },
    }
    tasks = [
        saved,
        {
            "id": "imported",
            "operation": "import",
            "nodeId": None,
            "status": "succeeded",
            "progress": {"phase": "succeeded"},
        },
    ]
    requests = []

    def handle(route):
        path = route.request.url.split("/mounted/", 1)[-1].split("?", 1)[0]
        if path == "assets/cytoscape.min.js":
            route.fulfill(
                body=(APP.parent / "cytoscape.min.js").read_text(encoding="utf-8"),
                content_type="application/javascript",
            )
            return
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/config":
            data = {
                "ai": {
                    "available": True,
                    "model": "local-doc",
                    "visionModel": "local-vision",
                    "languages": ["fr", "en"],
                    "tokenBudget": 1000,
                }
            }
        elif path == "api/projects":
            data = {"projects": [project]}
        elif path == "api/projects/english":
            data = {"project": project, "graph": graph, "tasks": tasks}
        elif path == "api/projects/english/tasks":
            body = route.request.post_data_json
            requests.append(body)
            task = {
                "id": f"new-{len(requests)}",
                "operation": body["operation"],
                "nodeId": body.get("nodeId"),
                "options": body.get("options", {}),
                "createdAt": 1790000000 + len(requests),
                "status": "succeeded" if body["operation"] == "evaluate" else "running",
            }
            if body["operation"] == "evaluate":
                task["result"] = {"status": "completed", "value": 4}
            tasks.append(task)
            data = {"task": task}
        elif path.endswith("/cancel"):
            task = next(t for t in tasks if t["id"] == path.split("/")[-2])
            task["status"] = "cancelled"
            data = {"task": task}
        elif path.startswith("api/tasks/"):
            task = next(t for t in tasks if t["id"] == path.rsplit("/", 1)[-1])
            data = {"task": task}
        else:
            data = {}
        route.fulfill(body=json.dumps(data), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/?lang=en")
        page.get_by_text("From file to formula.", exact=True).wait_for()
        assert page.locator("html").get_attribute("lang") == "en"
        assert page.locator("#ui-language").input_value() == "en"
        assert page.locator("#ai-language").input_value() == "fr"
        assert page.locator("#workbook").get_attribute("type") == "file"
        assert page.get_by_role("button", name="＋ Choose workbook").is_visible()
        page.locator('[data-project="english"]').click()
        page.get_by_role("tab", name="Operations", exact=True).click()
        page.locator("#task-list").get_by_text("Whole workbook", exact=True).wait_for()
        assert "Classeur entier" not in page.locator("#task-list").inner_text()
        page.get_by_role("tab", name="Graph", exact=True).click()
        assert (
            page.get_by_role("tab", name="Linexcel calculation").get_attribute(
                "aria-selected"
            )
            == "true"
        )
        assert page.locator("#graph-panel-calculation").is_visible()
        assert "AI" not in page.locator("#graph-panel-calculation").inner_text()
        assert page.locator("#graph-ai").is_hidden()
        page.get_by_role("tab", name="AI documentation", exact=True).click()
        page.locator("#graph-ai .ai-result").wait_for()
        assert (
            page.locator('#graph-sheet option[value="Feuilles"]').inner_text()
            == "Feuilles"
        )
        assert (
            "Documentation française conservée."
            in page.locator("#graph-ai").inner_text()
        )
        usage = page.locator("#graph-ai .token-usage").inner_text()
        assert "Input: 10" in usage and "Output: 5" in usage and "Total: 15" in usage
        assert "Model-reported counts" in usage
        saved["result"]["usage"] = {"outputTokens": 7, "estimated": True}
        page.locator('[data-project="english"]').click()
        page.get_by_role("tab", name="AI documentation", exact=True).click()
        page.locator("#graph-ai .token-usage .badge").get_by_text(
            "Estimated"
        ).wait_for()
        estimated = page.locator("#graph-ai .token-usage").inner_text()
        assert "Input: —" in estimated and "Output: 7" in estimated
        assert "Total: —" in estimated
        assert not [r for r in requests if r["operation"] == "document"]
        page.locator("#graph-full-detail").click()
        assert (
            "Documentation française conservée."
            in page.locator("#node-ai").inner_text()
        )
        assert not [r for r in requests if r["operation"] == "document"]
        page.get_by_text("On-demand AI documentation", exact=True).click()
        page.locator("#ai-language").select_option("en")
        assert page.locator("#node-ai .ai-result").count() == 0
        page.get_by_role("tab", name="Graph", exact=True).click()
        assert page.locator("#graph-ai .ai-result").count() == 0
        # Two immediate activations of the same control still create one request.
        page.locator("#graph-ai [data-ai-start]").evaluate(
            "button => { button.click(); button.click(); }"
        )
        page.locator("#graph-ai [data-ai-cancel]").wait_for()
        assert len([r for r in requests if r["operation"] == "document"]) == 1
        assert requests[-1]["options"]["language"] == "en"
        page.locator("#graph-full-detail").click()
        assert page.locator("#node-ai [data-ai-start]").is_disabled()
        page.locator("#node-ai [data-ai-cancel]").click()
        page.locator("#node-ai").get_by_text(
            "Generation cancelled.", exact=True
        ).wait_for()
        page.get_by_role("tab", name="Graph", exact=True).click()
        assert page.locator("#graph-ai [data-ai-cancel]").count() == 0
        assert page.locator("#graph-ai [data-ai-start]").is_enabled()
        page.locator("#ai-language").select_option("fr")
        assert page.locator("#graph-ai .ai-result").count() == 1
        assert page.locator("#node-ai .ai-result").count() == 1
        # Changing UI language does not change documentation language or request AI.
        page.locator("#ui-language").select_option("fr")
        page.get_by_role("tab", name="Graphe", exact=True).wait_for()
        assert page.locator("#ai-language").input_value() == "fr"
        page.locator("#ui-language").select_option("en")
        assert "lang=en" in page.url
        assert len([r for r in requests if r["operation"] == "document"]) == 1
        for tab, title in [
            ("sheets", "Workbook sheets"),
            ("captures", "Sheet captures"),
            ("tasks", "Operations and results"),
            ("diagnostics", "Analysis status"),
        ]:
            page.locator(f"#tab-{tab}").click()
            page.get_by_role("heading", name=title, exact=True).wait_for()
            assert (
                page.locator("#workspace-content > [role=tabpanel]:visible").count()
                == 1
            )
            assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
        assert errors == []
    finally:
        page.close()
