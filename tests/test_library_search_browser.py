"""API-driven browser contracts for the no-AI library search view."""

import json
import os
from pathlib import Path

import pytest
from test_app_browser import APP

playwright = pytest.importorskip("playwright.sync_api")
CAPTURE_DIR = (
    Path(os.environ["M3_CAPTURE_DIR"]) if os.environ.get("M3_CAPTURE_DIR") else None
)
VIEWPORTS = [(1440, 900), (390, 844)]
STATES = ["found", "empty", "capped", "boundary", "ambiguous", "pinned"]


@pytest.fixture(scope="module")
def browser():
    with playwright.sync_playwright() as driver:
        instance = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        yield instance
        instance.close()


@pytest.mark.parametrize("state", STATES)
@pytest.mark.parametrize("viewport", VIEWPORTS)
def test_library_search_states_are_traced_and_accessible(browser, state, viewport):
    page = browser.new_page(viewport={"width": viewport[0], "height": viewport[1]})
    page.set_default_timeout(6000)
    errors = []
    requests = []
    page.on("pageerror", lambda error: errors.append(error.stack or str(error)))
    project_id = "a" * 32
    revision = {
        "projectId": project_id,
        "sha256": "1" * 64,
        "id": f"{project_id}:{'1' * 64}",
    }
    second_revision = {
        "projectId": "b" * 32,
        "sha256": "2" * 64,
        "id": f"{'b' * 32}:{'2' * 64}",
    }
    hits = (
        []
        if state == "empty"
        else [
            {
                "kind": "cell",
                "filename": "same.xlsx",
                "revision": revision,
                "sheet": "S",
                "address": "A1",
                "value": '<img src=x onerror="window.injected=true">',
                "valueProvenance": "saved_input",
                "labels": [
                    {
                        "text": "Declared",
                        "role": "table_column",
                        "method": "table_column",
                        "sheet": "S",
                        "cell": "A1",
                    },
                    {
                        "text": "Inferred",
                        "role": "column",
                        "method": "header_band",
                        "certainty": "heuristic",
                        "sheet": "S",
                        "cell": "A2",
                    },
                ],
                "formula": "=1+2",
                "handle": "lx1.test",
            }
        ]
    )
    if state == "found":
        hits.append({**hits[0], "revision": second_revision, "handle": "lx1.other"})
    boundary_reason = {
        "found": "input_constant",
        "boundary": "macro",
        "ambiguous": "ambiguous_external",
        "pinned": "unresolved_external",
    }.get(state)
    caps = ["entries_per_workbook"] if state == "capped" else []
    search_trace = {
        "schemaVersion": "linexcel.search.v1",
        "tool": "search_catalog",
        "terms": ["cafe"],
        "normalization": "NFD fold; word tokens; AND",
        "scope": {"project": None, "revision": None, "kinds": ["cell"]},
        "revisions": [revision],
        "generation": "generation-1",
        "backend": "fts5",
        "counts": {"eligibleEntries": 10, "matched": len(hits), "returned": len(hits)},
        "capsHit": caps,
        "durationMs": 12,
        "complete": not caps,
    }
    status = {
        "buildState": {"status": "ready"},
        "stale": False,
        "revisions": [
            {"projectId": project_id, "filename": "same.xlsx", "revision": revision},
            {
                "projectId": second_revision["projectId"],
                "filename": "same.xlsx",
                "revision": second_revision,
            },
        ],
    }
    link = {
        "id": "link-1",
        "filename": "reader.xlsx",
        "relationshipPath": "reader.xlsx",
        "reference": "'[reader.xlsx]S'!A1",
        "sourceRevision": revision,
        "sourceNode": "S!A1",
        "status": "ambiguous" if state == "ambiguous" else "pinned",
        "pin": {"projectId": project_id, "sha256": revision["sha256"]}
        if state == "pinned"
        else None,
        "stalePin": state == "pinned",
        "targetRevision": None,
        "candidates": [
            {
                "id": revision["id"],
                "projectId": project_id,
                "sha256": revision["sha256"],
            }
        ],
    }
    link_payload = {"links": [link] if state in {"ambiguous", "pinned"} else []}
    long_node_value = "long-node-value-" * 20
    lineage = {
        "nodes": [
            {
                "id": "node-1",
                "filename": "same.xlsx",
                "projectId": project_id,
                "revision": revision,
                "sheet": "S",
                "address": "A1",
                "kind": "input",
                "value": "source-value",
            },
            {
                "id": "node-2",
                "filename": "reader.xlsx",
                "projectId": second_revision["projectId"],
                "revision": second_revision,
                "sheet": "T",
                "address": "B2",
                "kind": "formula",
                "value": long_node_value if state == "found" else "derived-value",
            },
        ],
        "edges": [{"source": "node-2", "target": "node-1", "kind": "cell_reference"}],
        "boundaries": [{"nodeId": "node-1", "reason": boundary_reason}]
        if boundary_reason
        else [],
        "coverage": {
            "complete": boundary_reason == "input_constant",
            "exhaustive": True,
            "boundaries": [boundary_reason] if boundary_reason else [],
            "capsHit": [],
            "nodesReturned": 2,
            "edgesReturned": 1,
        },
        "trace": {
            "schemaVersion": "linexcel.lineage.v1",
            "direction": "upstream",
            "generation": "generation-1",
            "revisions": [revision],
            "counts": {"nodes": 2, "edges": 1},
            "durationMs": 8,
            "capsHit": [],
            "complete": boundary_reason == "input_constant",
        },
    }

    def handle(route):
        url = route.request.url
        if url.endswith(("app-graph.js", "app-graph.css")):
            asset = APP.parent / url.rsplit("/", 1)[-1]
            route.fulfill(
                body=asset.read_text(encoding="utf-8"),
                content_type="text/css"
                if asset.suffix == ".css"
                else "application/javascript",
            )
            return
        path = url.split("/mounted/", 1)[-1]
        requests.append((route.request.method, path))
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/config":
            body = {
                "limits": {"memoryMb": 1024, "uploadMb": 64, "retentionDays": 7},
                "uploads": {"chunkBytes": 8388608},
            }
        elif path == "api/projects":
            body = {
                "projects": [
                    {
                        "id": project_id,
                        "name": "same.xlsx",
                        "filename": "same.xlsx",
                        "workbookRevision": revision,
                    }
                ]
            }
        elif path == "api/library/attachments":
            body = {"attachments": []}
        elif path.startswith("api/library/status"):
            body = status
        elif path.startswith("api/library/links"):
            body = link_payload
        elif path.startswith("api/library/search"):
            body = {
                "results": hits,
                "generation": "generation-1",
                "pagination": {
                    "offset": 0,
                    "limit": 50,
                    "matched": len(hits),
                    "hasMore": False,
                },
                "coverage": {"complete": not caps, "capsHit": caps},
                "trace": search_trace,
            }
        elif path.startswith("api/library/lineage"):
            direction = path.partition("?")[2]
            if "direction=downstream" in direction:
                lineage["trace"]["direction"] = "downstream"
            else:
                lineage["trace"]["direction"] = "upstream"
            body = lineage
        elif path.startswith("api/library/pins") and route.request.method == "POST":
            assert route.request.post_data_json["targetProject"] == project_id
            link.update(
                status="pinned",
                pin={"projectId": project_id, "sha256": revision["sha256"]},
                stalePin=False,
            )
            body = {"link": link}
        else:
            route.fulfill(status=404, json={"message": "unknown route"})
            return
        route.fulfill(body=json.dumps(body), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        try:
            page.wait_for_function(
                "document.querySelector('#connection')?.dataset.state === 'ok'",
                timeout=6000,
            )
        except Exception as cause:
            raise AssertionError(
                f"mock server status did not become ready; requests={requests}; "
                f"errors={errors}; "
                f"connection={page.locator('#connection').inner_text()}"
            ) from cause
        page.locator(".library-search-trigger").click()
        page.locator("#library-search-query").fill("café")
        page.get_by_role("button", name="Rechercher", exact=True).last.click()
        if hits:
            page.get_by_role(
                "button", name="Afficher la piste pour same.xlsx S!A1"
            ).first.click()
            try:
                page.locator(".library-verdict").wait_for(timeout=2500)
            except Exception as cause:
                raise AssertionError(
                    f"lineage did not render; requests={requests}; errors={errors}; "
                    f"page={page.locator('#library-search-view').inner_text()}"
                ) from cause
        else:
            page.locator(".library-empty").first.wait_for()
        if state in {"ambiguous", "pinned"}:
            page.locator(".library-link").get_by_text(
                "reader.xlsx", exact=False
            ).wait_for()
        if state == "ambiguous":
            target = page.get_by_label("Révision cible pour link-1")
            target.select_option(f"{project_id}:{revision['sha256']}")
            page.get_by_role("button", name="Épingler cette cible", exact=True).click()
            try:
                page.locator(".library-link > div").get_by_text(
                    "cible choisie", exact=True
                ).wait_for()
            except Exception as cause:
                raise AssertionError(
                    f"pin did not render; requests={requests}; errors={errors}; "
                    f"links={page.locator('.library-links').inner_text()}"
                ) from cause
        if state == "pinned":
            page.get_by_text("épingle périmée", exact=False).wait_for()
        if state == "capped":
            page.get_by_text("Résultats partiels", exact=False).wait_for()
            page.get_by_text(
                "limite d’entrées du classeur", exact=False
            ).first.wait_for()
        if state == "empty":
            page.locator(".library-trace-summary a").click()
            page.locator(".library-trace-last").get_by_text(
                "ne prouve pas l’absence"
            ).wait_for()
        if state == "boundary":
            page.locator(".library-graph-canvas .flow-card.boundary").get_by_text(
                "macro non exécutée", exact=False
            ).wait_for()
        if state == "found":
            page.get_by_text("Inféré · heuristique").first.wait_for()
            page.get_by_text("Origine de la valeur", exact=False).first.wait_for()
            page.locator(".library-revision-group").get_by_text(
                "Formule", exact=False
            ).wait_for()
            assert page.locator(".library-trace-last.is-incomplete").count() == 0
            page.locator(".library-accessible-graph").get_by_text(
                "source-value", exact=False
            ).wait_for()
            page.locator(".library-revision-navigation button").filter(
                has_text="reader.xlsx"
            ).click()
            long_value_card = page.locator(
                '.library-graph-canvas .flow-card[title*="long-node-value-"]'
            )
            assert long_node_value in long_value_card.get_attribute("title")
            value = long_value_card.locator(".flow-value")
            assert (
                value.evaluate("element => getComputedStyle(element).textOverflow")
                == "ellipsis"
            )
            assert long_value_card.evaluate(
                "element => element.scrollWidth <= element.clientWidth"
            )
            assert page.locator(".library-result").count() == 2
            assert page.locator("#library-search-view img").count() == 0
            assert (
                page.locator("#library-search-view")
                .get_by_text('<img src=x onerror="window.injected=true">')
                .count()
                == 2
            )
            page.get_by_label("Sens").select_option("downstream")
            page.wait_for_function(
                """document.querySelector('.library-trace-summary')
                  .innerText.includes('En aval')"""
            )
            assert any(
                "direction=downstream" in path
                for _, path in requests
                if "library/lineage" in path
            )
            assert (
                page.locator(".library-result").nth(0).inner_text()
                != page.locator(".library-result").nth(1).inner_text()
            )
        assert page.evaluate(
            "document.documentElement.scrollWidth <= window.innerWidth"
        ), page.evaluate(
            """() => [...document.querySelectorAll("body *")]
              .filter((element) =>
                element.getBoundingClientRect().right > innerWidth + 1)
              .slice(0, 10)
              .map((element) => ({
                tag: element.tagName,
                className: element.className,
                right: element.getBoundingClientRect().right,
                text: element.innerText?.slice(0, 80),
              }))"""
        )
        if viewport[0] == 390:
            height = page.locator(".library-search-form-row button").evaluate(
                "(el) => el.getBoundingClientRect().height"
            )
            assert round(height) >= 44
        assert not any("/ai/" in path for _, path in requests)
        assert errors == []
        if CAPTURE_DIR:
            CAPTURE_DIR.mkdir(parents=True, exist_ok=True)
            prefix = f"{state}-{viewport[0]}x{viewport[1]}"
            page.locator(".workspace > main").evaluate("(el) => el.scrollTo(0, 0)")
            page.screenshot(path=str(CAPTURE_DIR / f"{prefix}-overview.png"))
            for section in (".library-lineage", ".library-links", ".library-trace"):
                page.evaluate(
                    """(selector) => {
                      const main = document.querySelector(".workspace > main");
                      const target = document.querySelector(selector);
                      main.scrollTop += target.getBoundingClientRect().top
                        - main.getBoundingClientRect().top - 60;
                    }""",
                    section,
                )
                if section == ".library-lineage":
                    heading_top = page.locator("#library-lineage-heading").evaluate(
                        "(el) => el.getBoundingClientRect().top"
                    )
                    assert heading_top >= 48
                page.screenshot(
                    path=str(
                        CAPTURE_DIR
                        / f"{prefix}-{section.removeprefix('.library-')}.png"
                    )
                )
    finally:
        page.close()


def test_pin_unpin_and_direction_changes_replace_the_lineage_verdict(browser):
    page = browser.new_page(viewport={"width": 390, "height": 844})
    page.set_default_timeout(6000)
    project_id = "a" * 32
    revision = {
        "projectId": project_id,
        "sha256": "1" * 64,
        "id": f"{project_id}:{'1' * 64}",
    }
    pinned = {"value": True}
    requests = []
    link = {
        "id": "link-1",
        "bookKey": "target.xlsx",
        "filename": "target.xlsx",
        "relationshipPath": "target.xlsx",
        "reference": "'[target.xlsx]Sheet'!A1",
        "sourceRevision": revision,
        "sourceNode": "Sheet!A1",
        "status": "pinned",
        "pin": {"projectId": project_id, "sha256": revision["sha256"]},
        "targetRevision": revision,
        "candidates": [
            {
                "id": revision["id"],
                "projectId": project_id,
                "sha256": revision["sha256"],
            }
        ],
        "stalePin": False,
    }

    def handle(route):
        url = route.request.url
        if url.endswith(("app-graph.js", "app-graph.css")):
            asset = APP.parent / url.rsplit("/", 1)[-1]
            route.fulfill(
                body=asset.read_text(encoding="utf-8"),
                content_type="text/css"
                if asset.suffix == ".css"
                else "application/javascript",
            )
            return
        path = url.split("/mounted/", 1)[-1]
        requests.append((route.request.method, path))
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/config":
            body = {
                "limits": {"memoryMb": 1024, "uploadMb": 64, "retentionDays": 7},
                "uploads": {"chunkBytes": 8388608},
            }
        elif path == "api/projects":
            body = {
                "projects": [
                    {"id": project_id, "name": "demo.xlsx", "filename": "demo.xlsx"}
                ]
            }
        elif path == "api/library/attachments":
            body = {"attachments": []}
        elif path.startswith("api/library/status"):
            body = {
                "buildState": {"status": "ready"},
                "stale": False,
                "revisions": [
                    {
                        "projectId": project_id,
                        "filename": "demo.xlsx",
                        "revision": revision,
                    }
                ],
            }
        elif path.startswith("api/library/links"):
            body = {
                "generation": "generation-1",
                "links": [
                    {
                        **link,
                        "status": "pinned" if pinned["value"] else "ambiguous",
                        "pin": link["pin"] if pinned["value"] else None,
                    }
                ],
                "pagination": {
                    "offset": 0,
                    "limit": 50,
                    "matched": 1,
                    "hasMore": False,
                },
                "coverage": {"complete": True, "capsHit": []},
            }
        elif path.startswith("api/library/search"):
            body = {
                "generation": "generation-1",
                "results": [
                    {
                        "kind": "cell",
                        "filename": "demo.xlsx",
                        "revision": revision,
                        "sheet": "Sheet",
                        "address": "A1",
                        "value": 7,
                        "valueProvenance": "saved_input",
                        "handle": "lx1.root",
                    }
                ],
                "pagination": {
                    "offset": 0,
                    "limit": 50,
                    "matched": 1,
                    "hasMore": False,
                },
                "coverage": {"complete": True, "capsHit": []},
                "trace": {
                    "schemaVersion": "linexcel.search.v1",
                    "tool": "search_catalog",
                    "terms": ["7"],
                    "normalization": (
                        "NFD, casefold, remove combining marks; word tokens; AND"
                    ),
                    "scope": {"kinds": ["cell"]},
                    "revisions": [revision],
                    "generation": "generation-1",
                    "backend": "fts5",
                    "counts": {"eligibleEntries": 1, "matched": 1, "returned": 1},
                    "durationMs": 1,
                    "complete": True,
                    "capsHit": [],
                },
            }
        elif path.startswith("api/library/lineage"):
            direction = "downstream" if "direction=downstream" in path else "upstream"
            is_complete = pinned["value"]
            body = {
                "nodes": [
                    {
                        "id": "n1",
                        "filename": "demo.xlsx",
                        "projectId": project_id,
                        "revision": revision,
                        "sheet": "Sheet",
                        "address": "A1",
                        "kind": "input",
                        "value": 7,
                        "valueProvenance": "saved_input",
                    }
                ],
                "edges": [],
                "boundaries": []
                if is_complete
                else [
                    {
                        "node": "n1",
                        "reason": "ambiguous_external",
                        "relationshipPath": "target.xlsx",
                    }
                ],
                "coverage": {
                    "complete": is_complete,
                    "exhaustive": True,
                    "boundaries": [] if is_complete else ["ambiguous_external"],
                    "capsHit": [],
                },
                "trace": {
                    "schemaVersion": "linexcel.lineage.v1",
                    "tool": "directed_lineage",
                    "direction": direction,
                    "generation": "generation-1",
                    "revisions": [revision],
                    "limits": {"depth": 20},
                    "counts": {"nodes": 1, "edges": 0},
                    "complete": is_complete,
                    "exhaustive": True,
                    "capsHit": [],
                },
            }
        elif path.startswith("api/library/pins/") and route.request.method == "DELETE":
            pinned["value"] = False
            body = {"link": {**link, "status": "ambiguous", "pin": None}}
        else:
            route.fulfill(status=404, json={"message": "unknown route"})
            return
        route.fulfill(body=json.dumps(body), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        page.wait_for_function(
            "document.querySelector('#connection')?.dataset.state === 'ok'"
        )
        page.locator(".library-search-trigger").click()
        page.locator("#library-search-query").fill("7")
        page.get_by_role("button", name="Rechercher", exact=True).last.click()
        page.get_by_role(
            "button", name="Afficher la piste pour demo.xlsx Sheet!A1"
        ).click()
        page.locator(".library-verdict.is-complete").wait_for()
        page.get_by_role("button", name="Retirer l’épingle").click()
        page.locator(".library-verdict.is-incomplete").wait_for()
        assert page.locator(".library-verdict.is-complete").count() == 0
        page.get_by_label("Sens").select_option("downstream")
        page.wait_for_timeout(100)
        summary = page.locator(".library-trace-summary").inner_text()
        assert "En aval" in summary, (
            f"summary={summary!r}; requests={requests}; "
            f"trace={page.locator('.library-trace li').all_inner_texts()}"
        )
        assert any(
            "direction=downstream" in path
            for _, path in requests
            if "library/lineage" in path
        )
    finally:
        page.close()


def test_failed_search_clears_previous_result_and_records_error(browser):
    page = browser.new_page(viewport={"width": 390, "height": 844})
    page.set_default_timeout(6000)
    project_id = "a" * 32
    revision = {
        "projectId": project_id,
        "sha256": "1" * 64,
        "id": f"{project_id}:{'1' * 64}",
    }

    def handle(route):
        url = route.request.url
        if url.endswith(("app-graph.js", "app-graph.css")):
            asset = APP.parent / url.rsplit("/", 1)[-1]
            route.fulfill(
                body=asset.read_text(encoding="utf-8"),
                content_type="text/css"
                if asset.suffix == ".css"
                else "application/javascript",
            )
            return
        path = url.split("/mounted/", 1)[-1]
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/config":
            body = {
                "limits": {"memoryMb": 1024, "uploadMb": 64, "retentionDays": 7},
                "uploads": {"chunkBytes": 8388608},
            }
        elif path == "api/projects":
            body = {"projects": []}
        elif path == "api/library/attachments":
            body = {"attachments": []}
        elif path.startswith("api/library/status"):
            body = {"buildState": {"status": "ready"}, "stale": False, "revisions": []}
        elif path.startswith("api/library/links"):
            body = {
                "links": [],
                "pagination": {
                    "offset": 0,
                    "limit": 50,
                    "matched": 0,
                    "hasMore": False,
                },
                "coverage": {"complete": True, "capsHit": []},
            }
        elif path.startswith("api/library/search") and "q=fail" in path:
            route.fulfill(
                status=400, json={"error": {"message": "Query exceeds query_chars"}}
            )
            return
        elif path.startswith("api/library/search"):
            body = {
                "generation": "generation-1",
                "results": [
                    {
                        "kind": "cell",
                        "filename": "demo.xlsx",
                        "revision": revision,
                        "sheet": "Sheet",
                        "address": "A1",
                        "value": 7,
                        "valueProvenance": "saved_input",
                        "handle": "lx1.root",
                    }
                ],
                "pagination": {
                    "offset": 0,
                    "limit": 50,
                    "matched": 1,
                    "hasMore": False,
                },
                "coverage": {"complete": True, "capsHit": []},
                "trace": {
                    "schemaVersion": "linexcel.search.v1",
                    "tool": "search_catalog",
                    "terms": ["ok"],
                    "normalization": (
                        "NFD, casefold, remove combining marks; word tokens; AND"
                    ),
                    "scope": {"kinds": ["cell"]},
                    "revisions": [revision],
                    "generation": "generation-1",
                    "backend": "fts5",
                    "counts": {"eligibleEntries": 1, "matched": 1, "returned": 1},
                    "durationMs": 1,
                    "complete": True,
                    "capsHit": [],
                },
            }
        else:
            route.fulfill(status=404, json={"message": "unknown route"})
            return
        route.fulfill(body=json.dumps(body), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        page.wait_for_function(
            "document.querySelector('#connection')?.dataset.state === 'ok'"
        )
        page.locator(".library-search-trigger").click()
        query = page.locator("#library-search-query")
        query.fill("ok")
        page.get_by_role("button", name="Rechercher", exact=True).last.click()
        page.get_by_text("demo.xlsx", exact=False).wait_for()
        query.fill("fail")
        assert page.locator(".library-result").count() == 0
        page.get_by_role("button", name="Rechercher", exact=True).last.click()
        page.get_by_role("alert").get_by_text(
            "requête est trop longue", exact=False
        ).wait_for()
        assert page.locator(".library-result").count() == 0
        assert "Échec" in page.locator(".library-trace-summary").inner_text()
    finally:
        page.close()


def test_failed_scoped_import_is_cleaned_through_attachment_route(browser):
    page = browser.new_page(viewport={"width": 390, "height": 844})
    page.set_default_timeout(6000)
    removed = []

    def handle(route):
        url = route.request.url
        if url.endswith(("app-graph.js", "app-graph.css")):
            asset = APP.parent / url.rsplit("/", 1)[-1]
            route.fulfill(
                body=asset.read_text(encoding="utf-8"),
                content_type="text/css"
                if asset.suffix == ".css"
                else "application/javascript",
            )
            return
        path = url.split("/mounted/", 1)[-1]
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/config":
            body = {
                "limits": {"memoryMb": 1024, "uploadMb": 64, "retentionDays": 7},
                "uploads": {"chunkBytes": 8388608},
            }
        elif path == "api/projects":
            body = {"projects": []}
        elif path == "api/library/attachments":
            body = {"attachments": []}
        elif path.startswith("api/library/status"):
            body = {"buildState": {"status": "ready"}, "stale": False, "revisions": []}
        elif path.startswith("api/library/links"):
            body = {
                "links": [],
                "pagination": {
                    "offset": 0,
                    "limit": 50,
                    "matched": 0,
                    "hasMore": False,
                },
                "coverage": {"complete": True, "capsHit": []},
            }
        elif path == "api/uploads" and route.request.method == "POST":
            body = {"upload": {"id": "upload-1", "chunkBytes": 8388608}}
        elif path.startswith("api/uploads/upload-1/files/"):
            route.fulfill(status=200, json={"upload": {"id": "upload-1"}})
            return
        elif path == "api/uploads/upload-1/complete":
            body = {
                "project": {
                    "id": "c" * 32,
                    "filename": "broken.xlsx",
                    "searchScopeId": "upload-1",
                },
                "task": {
                    "id": "task-1",
                    "status": "failed",
                    "error": {"message": "File is not a zip file"},
                },
            }
        elif (
            path == "api/library/attachments/upload-1"
            and route.request.method == "DELETE"
        ):
            removed.append(path)
            body = {}
        else:
            route.fulfill(status=404, json={"message": "unknown route"})
            return
        route.fulfill(body=json.dumps(body), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        page.wait_for_function(
            "document.querySelector('#connection')?.dataset.state === 'ok'"
        )
        page.locator(".library-search-trigger").click()
        page.locator("#library-attach-file").set_input_files(
            {
                "name": "broken.xlsx",
                "mimeType": (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
                "buffer": b"PK\x03\x04broken",
            }
        )
        page.get_by_role("alert").get_by_text("endommagé", exact=False).wait_for()
        assert removed == ["api/library/attachments/upload-1"]
        assert "Excel" in page.locator(".library-trace-summary").inner_text()
    finally:
        page.close()


def test_scoped_attachment_waits_for_catalog_before_loading_links(browser):
    page = browser.new_page(viewport={"width": 390, "height": 844})
    page.set_default_timeout(6000)
    project_id = "a" * 32
    revision = {
        "projectId": project_id,
        "sha256": "1" * 64,
        "id": f"{project_id}:{'1' * 64}",
    }
    scoped_status_calls = {"count": 0}
    scoped_link_calls = {"pending": 0, "ready": 0}
    requests = []

    def handle(route):
        url = route.request.url
        if url.endswith(("app-graph.js", "app-graph.css")):
            asset = APP.parent / url.rsplit("/", 1)[-1]
            route.fulfill(
                body=asset.read_text(encoding="utf-8"),
                content_type="text/css"
                if asset.suffix == ".css"
                else "application/javascript",
            )
            return
        path = url.split("/mounted/", 1)[-1]
        requests.append((route.request.method, path))
        if not path:
            route.fulfill(
                body=APP.read_text(encoding="utf-8"), content_type="text/html"
            )
            return
        if path == "api/config":
            body = {
                "limits": {"memoryMb": 1024, "uploadMb": 64, "retentionDays": 7},
                "uploads": {"chunkBytes": 8388608},
            }
        elif path == "api/projects":
            body = {"projects": []}
        elif path == "api/library/attachments":
            body = {"attachments": []}
        elif path.startswith("api/library/status"):
            if "scope=scope-1" not in path:
                body = {"buildState": {"status": "ready"}, "revisions": []}
            else:
                scoped_status_calls["count"] += 1
                state = "building" if scoped_status_calls["count"] < 5 else "ready"
                body = {
                    "buildState": {"status": state, "scope": "scope-1"},
                    "stale": state == "building",
                    "revisions": [
                        {
                            "projectId": project_id,
                            "filename": "attached.xlsx",
                            "revision": revision,
                        }
                    ],
                    "coverage": {
                        "complete": state == "ready",
                        "capsHit": []
                        if state == "ready"
                        else ["catalog_refresh_pending"],
                    },
                }
        elif path.startswith("api/library/links"):
            if "scope=scope-1" in path and scoped_status_calls["count"] < 5:
                scoped_link_calls["pending"] += 1
                route.fulfill(
                    status=409,
                    json={"error": {"message": "Catalog building; retry lookup"}},
                )
                return
            if "scope=scope-1" in path:
                scoped_link_calls["ready"] += 1
            body = {
                "links": [],
                "pagination": {"offset": 0, "limit": 50, "hasMore": False},
                "coverage": {"complete": True, "capsHit": []},
            }
        elif path == "api/uploads" and route.request.method == "POST":
            body = {"upload": {"id": "upload-1", "chunkBytes": 8388608}}
        elif path.startswith("api/uploads/upload-1/files/"):
            route.fulfill(status=200, json={"upload": {"id": "upload-1"}})
            return
        elif path == "api/uploads/upload-1/complete":
            body = {
                "project": {
                    "id": project_id,
                    "filename": "attached.xlsx",
                    "searchScopeId": "scope-1",
                    "searchScopeExpiresAt": 2000000000,
                },
                "task": {"id": "task-1", "status": "succeeded"},
            }
        else:
            route.fulfill(status=404, json={"message": "unknown route"})
            return
        route.fulfill(body=json.dumps(body), content_type="application/json")

    page.route("**/*", handle)
    try:
        page.goto("http://localhost/mounted/")
        page.wait_for_function(
            "document.querySelector('#connection')?.dataset.state === 'ok'"
        )
        page.locator(".library-search-trigger").click()
        page.locator("#library-attach-file").set_input_files(
            {
                "name": "attached.xlsx",
                "mimeType": (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
                "buffer": b"PK\x03\x04test",
            }
        )
        page.locator(".library-attachment-actions button").first.wait_for()
        for _ in range(60):
            if scoped_link_calls["ready"]:
                break
            page.wait_for_timeout(100)
        assert scoped_status_calls["count"] >= 5
        assert scoped_link_calls == {"pending": 0, "ready": 1}
        assert not any(
            method == "DELETE" and "api/library/attachments/scope-1" in path
            for method, path in requests
        )
        assert page.locator(".library-attachment-actions button").count() == 2
        assert page.get_by_role("alert").count() == 0
    finally:
        page.close()
