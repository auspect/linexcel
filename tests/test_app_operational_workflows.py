"""Reviewable user journeys across local sessions, data and portable reports."""

import test_app_viewer_regressions as serve_tests
from graph_test_helpers import open_graph_options, open_workspace_menu

app = serve_tests.app
browser = serve_tests.browser


def test_mobile_menu_escape_and_first_real_content(app):
    page, _, errors, _, _ = app
    first = page.locator("#node-list [data-node]").first.bounding_box()
    assert first["y"] < page.viewport_size["height"] * 0.6
    if page.viewport_size["width"] == 390:
        assert page.locator("#workspace-sidebar").is_hidden()
        page.locator("#workspace-menu").click()
        assert page.locator("main").evaluate("e=>e.inert")
        assert page.locator("#workspace-close").evaluate(
            "e=>e===document.activeElement"
        )
        page.keyboard.press("Escape")
        assert not page.locator("main").evaluate("e=>e.inert")
        assert page.locator("#workspace-menu").evaluate("e=>e===document.activeElement")
    assert not errors


def test_readable_focal_graph_and_mobile_value_comparison(app):
    page, _, errors, _, _ = app
    page.locator("#tab-graph").click()
    page.locator("#graph-sheet").select_option("Second")
    page.wait_for_function(
        "document.querySelector('#graph-canvas')._cy?.nodes().length===2"
    )
    open_graph_options(page)
    page.locator("#graph-focus").click()
    page.wait_for_function("document.querySelector('#graph-canvas')._cy.zoom()*13>=12")
    metrics = page.locator("#graph-canvas").evaluate("""e=>{
      const cy=e._cy,n=cy.getElementById('Second!A2'),p=n.renderedPosition();
      return {font:n.numericStyle('font-size')*cy.zoom(),x:p.x,y:p.y,
        width:cy.width(),height:cy.height(),hostWidth:e.clientWidth,hostHeight:e.clientHeight};
    }""")
    assert metrics["font"] >= 12
    assert 0 < metrics["x"] < metrics["width"]
    assert 0 < metrics["y"] < metrics["height"]
    assert abs(metrics["width"] - metrics["hostWidth"]) <= 2
    assert abs(metrics["height"] - metrics["hostHeight"]) <= 2
    if page.viewport_size["width"] == 390:
        page.locator("#graph-read-cell").click()
        assert page.locator("#graph-canvas").is_hidden()
        cards = page.locator("#graph-panel-calculation .value-card")
        assert (
            cards.nth(1).bounding_box()["y"]
            >= cards.nth(0).bounding_box()["y"] + cards.nth(0).bounding_box()["height"]
        )
        page.keyboard.press("Escape")
        assert page.locator("#graph-canvas").is_visible()
    assert not errors


def test_report_explicit_selection_budget_and_download(app):
    page, requests, errors, _, fulfill = app
    submissions = []

    def submit(route):
        submissions.append(route.request.post_data_json)
        fulfill(
            route,
            {
                "task": {
                    "id": "report-1",
                    "operation": "export_html",
                    "status": "succeeded",
                }
            },
        )

    page.route("**/api/projects/p/tasks", submit)
    page.locator(".project-utilities > summary").click()
    page.locator("#report-open").click()
    page.locator("#report-nodes").fill("\n".join(f"First!A{i}" for i in range(1, 22)))
    page.locator("#report-generate").click()
    assert not submissions
    assert "20" in page.locator("#report-status").inner_text()
    page.locator("#report-nodes").fill("First!A1\nFirst!A1\nSecond!A2")
    page.locator("#report-language").select_option("en")
    page.locator("#report-generate").click()
    page.locator("#report-download").wait_for()
    assert submissions == [
        {
            "operation": "export_html",
            "options": {
                "nodeIds": ["First!A1", "Second!A2"],
                "captureIds": [],
                "language": "en",
            },
            "budget": {"seconds": 600, "memoryMb": 1024},
        }
    ]
    assert (
        page.locator("#report-download").get_attribute("href")
        == "api/tasks/report-1/artifact"
    )
    page.keyboard.press("Escape")
    assert not page.locator("#report-dialog").evaluate("e=>e.open")
    assert not errors


def test_data_understand_and_return_preserves_cell_and_scroll(app):
    page, _, errors, _, fulfill = app
    page.route(
        "**/api/projects/p/data",
        lambda route: fulfill(
            route,
            {
                "available": True,
                "presentationAvailable": True,
                "sourceSha256": "a" * 64,
                "sheets": [
                    {
                        "sheet": "First",
                        "url": "/api/projects/p/datasette/source/first",
                        "rowCount": 80,
                        "cellCount": 80,
                    }
                ],
            },
        ),
    )
    page.route(
        "**/api/projects/p/data/presentation?*",
        lambda route: fulfill(
            route,
            {
                "sheet": "First",
                "sourceSha256": "a" * 64,
                "window": {"row": 1, "column": 1, "rows": 60, "columns": 20},
                "extent": {"rows": 80, "columns": 20},
                "defaults": {"columnWidth": 64, "rowHeight": 20},
                "cells": [
                    {
                        "address": "A1",
                        "row": 1,
                        "column": 1,
                        "value": 1,
                        "display": "1",
                        "numberFormat": "General",
                    }
                ],
            },
        ),
    )
    page.locator("#tab-data").click()
    page.locator("[data-cell='A1']").click()
    page.locator("#excel-window").evaluate("e=>{e.scrollLeft=170;e.scrollTop=100}")
    saved = page.locator("#excel-window").evaluate("e=>[e.scrollLeft,e.scrollTop]")
    page.locator("#excel-understand").click()
    page.locator("#view-graph").wait_for()
    assert "First · A1" in page.locator("#graph-preview").inner_text()
    page.locator("#data-return").click()
    page.locator("#view-data").wait_for()
    assert (
        page.locator("#excel-window").evaluate("e=>[e.scrollLeft,e.scrollTop]") == saved
    )
    assert page.locator("[data-cell='A1']").evaluate("e=>e===document.activeElement")
    assert page.locator("#data-sheet").input_value() == "0"
    assert not errors


def test_session_fragment_consumed_and_health_never_generates(app):
    page, requests, errors, _, fulfill = app
    session = []
    page.route(
        "**/api/session",
        lambda route: (
            session.append(route.request.post_data_json),
            fulfill(route, {"authenticated": True}),
        ),
    )
    page.goto("http://localhost/#session=review-only-bootstrap")
    page.locator("[data-node='First!A1']").wait_for()
    assert session == [{"token": "review-only-bootstrap"}]
    assert "#" not in page.url
    page.route(
        "**/api/health",
        lambda route: fulfill(
            route,
            {
                "checks": {
                    "storage": {
                        "status": "ready",
                        "usedBytes": 1024,
                        "quotaBytes": 1048576,
                    },
                    "ai": {"status": "disabled", "visionCapability": "not_probed"},
                    "libreoffice": {
                        "status": "unavailable",
                        "code": "executable_missing",
                    },
                },
                "generationRequested": False,
            },
        ),
    )
    open_workspace_menu(page)
    page.locator("#server-health > summary").click()
    page.locator("#health-check").click()
    page.locator("#health-status").get_by_text("Stockage", exact=True).wait_for()
    assert "PATH" in page.locator("#health-status").inner_text()
    assert "non testée" in page.locator("#health-status").inner_text()
    assert not any(path.endswith("/tasks") for path, _ in requests)
    assert not errors


def test_current_capability_registry_is_localized_without_legacy_refusals(app):
    from linexcel.capabilities import capability_registry

    page, _, errors, _, fulfill = app
    registry = capability_registry()
    registry["features"][0]["limitations"].append("future_unknown_capability_limit")
    page.route(
        "**/api/projects/p",
        lambda route: fulfill(
            route,
            {
                "project": {"id": "p", "name": "Current capabilities", "revision": 2},
                "graph": {"nodes": [], "edges": [], "meta": {"capabilities": registry}},
                "tasks": [],
            },
        ),
    )
    open_workspace_menu(page)
    page.locator("[data-project='p']").click()
    page.locator("#tab-diagnostics").click()
    text = page.locator("#diagnostics").inner_text()
    assert "définitions calculées statiques" in text
    assert "plages, noms, tableaux et formules" in text
    assert "INDIRECT et OFFSET résolus" in text
    assert "équivalence complète" in text
    assert "formules, plages et dates externes non calculées" not in text
    assert "future_unknown_capability_limit" not in text
    assert (
        "future_unknown_capability_limit" in page.locator("#diagnostics").text_content()
    )
    page.locator("#diagnostics details").filter(
        has=page.locator("summary", has_text="Limites documentées")
    ).first.locator(":scope > summary").click()
    assert "Régions de données vides" in page.locator("#diagnostics").inner_text()
    page.locator("#ui-language").select_option("en")
    assert "static computed definitions" in page.locator("#diagnostics").inner_text()
    assert "Recorded UTC clock" in page.locator("#diagnostics").inner_text()
    assert (
        "parity with native Excel is not established"
        in page.locator("#diagnostics").inner_text()
    )
    assert not errors


def test_import_completion_poll_reloads_completed_capture_gallery(app):
    """Import refresh must not replace visible capture evidence with summaries."""
    import base64

    from vision_helpers import image_bytes

    page, _, errors, _, fulfill = app
    completed = False
    capture_reads = []
    encoded = base64.b64encode(image_bytes()).decode()
    task_id = "c" * 32
    project = {"id": "p", "name": "Import and capture", "revision": 2}
    graph = {
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
        "meta": {"sheets": ["First"]},
    }

    def project_reply(route):
        fulfill(
            route,
            {
                "project": project,
                "graph": graph,
                "tasks": [
                    {
                        "id": "import-poll",
                        "operation": "import",
                        "status": "succeeded" if completed else "running",
                        "createdAt": 1,
                    },
                    {
                        "id": task_id,
                        "operation": "capture",
                        "status": "succeeded" if completed else "running",
                        "createdAt": 2,
                        "resultAvailable": completed,
                    },
                ],
            },
        )

    def task_reply(route):
        nonlocal completed
        completed = True
        if route.request.url.endswith(task_id):
            capture_reads.append(route.request.url)
            fulfill(
                route,
                {
                    "task": {
                        "id": task_id,
                        "operation": "capture",
                        "status": "succeeded",
                        "createdAt": 2,
                        "resultAvailable": True,
                        "result": {
                            "screenshots": [
                                {
                                    "sheet": "First",
                                    "data": "data:image/png;base64," + encoded,
                                }
                            ]
                        },
                    }
                },
            )
        else:
            fulfill(
                route,
                {
                    "task": {
                        "id": "import-poll",
                        "operation": "import",
                        "status": "succeeded",
                        "createdAt": 1,
                    }
                },
            )

    page.route("**/api/projects/p", project_reply)
    page.route("**/api/tasks/import-poll", task_reply)
    page.route("**/api/tasks/" + task_id, task_reply)
    open_workspace_menu(page)
    page.locator("[data-project='p']").click()
    page.locator("#tab-captures").click()
    page.locator("#capture-selected figure img").wait_for(timeout=10000)
    page.wait_for_function(
        "document.querySelector('#capture-selected figure img').naturalWidth>0"
    )
    assert len(capture_reads) >= 2
    assert not errors
