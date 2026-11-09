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
        if state in {"empty", "capped"}
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
                    },
                    {
                        "text": "Inferred",
                        "role": "column",
                        "method": "header_band",
                        "certainty": "heuristic",
                    },
                ],
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
        "path": "[1]S!A1",
        "status": "ambiguous" if state == "ambiguous" else "pinned",
        "pin": {"projectId": project_id, "sha256": revision["sha256"]}
        if state == "pinned"
        else None,
        "stalePin": state == "pinned",
        "candidates": [],
    }
    link_payload = {"links": [link] if state in {"ambiguous", "pinned"} else []}
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
                "value": "derived-value",
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
            page.get_by_role("button", name="Épingler", exact=True).click()
            page.locator(".library-link > div").get_by_text(
                "pinned", exact=True
            ).wait_for()
        if state == "pinned":
            page.get_by_text("épingle périmée", exact=False).wait_for()
        if state == "capped":
            page.get_by_text("recherche tronquée", exact=False).wait_for()
            assert page.locator(".library-trace-last.is-incomplete").count() == 1
        if state == "empty":
            page.locator(".library-trace-last").get_by_text(
                "ne prouve pas l’absence"
            ).wait_for()
        if state == "boundary":
            page.get_by_test_id("rf__node-boundary:0").get_by_text(
                "macro non exécutée", exact=False
            ).wait_for()
        if state == "found":
            page.get_by_text("Inféré · heuristique").first.wait_for()
            assert page.locator(".library-trace-last.is-incomplete").count() == 0
            page.locator(".library-accessible-graph").get_by_text(
                "reader.xlsx · 22222222 · T!B2 · formula → same.xlsx · "
                "11111111 · S!A1 · input",
                exact=False,
            ).wait_for()
            page.get_by_text("éléments synthétiques", exact=False).wait_for()
            assert page.locator(".library-result").count() == 2
            assert page.locator("#library-search-view img").count() == 0
            assert (
                page.locator("#library-search-view")
                .get_by_text('<img src=x onerror="window.injected=true">')
                .count()
                == 2
            )
            assert (
                page.locator(".library-result").nth(0).inner_text()
                != page.locator(".library-result").nth(1).inner_text()
            )
        assert page.evaluate(
            "document.documentElement.scrollWidth <= window.innerWidth"
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
