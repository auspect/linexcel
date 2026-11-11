"""Real WSGI/API envelopes for request identity, revision views and trace access."""

import io
import json
import os
import threading
import time
from socketserver import ThreadingMixIn
from wsgiref.simple_server import WSGIRequestHandler, WSGIServer, make_server

import pytest
from openpyxl import Workbook
from test_external import _link, linked_workbook
from test_web import request, wait

from linexcel.catalog import CatalogLimits
from linexcel.web import create_app

playwright = pytest.importorskip("playwright.sync_api")


@pytest.fixture(scope="module")
def browser():
    with playwright.sync_playwright() as driver:
        instance = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        yield instance
        instance.close()


def book_bytes(label="ALPHA", value=7):
    book = Workbook()
    book.active.append(["Key", "Value"])
    book.active.append([label, value])
    for number in range(70):
        book.active.append([f"capneedle {number}", number])
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


@pytest.fixture
def live_library(tmp_path):
    app = create_app(
        tmp_path / "data",
        resolve_user=lambda env: "alice",
        catalog_limits=CatalogLimits(hits=50),
    )

    class Server(ThreadingMixIn, WSGIServer):
        daemon_threads = True

    class Handler(WSGIRequestHandler):
        def log_message(self, *args):
            pass

    gate = {"active": False, "started": threading.Event(), "release": threading.Event()}
    app.test_search_gate = gate

    def gated_app(environ, start_response):
        if gate["active"] and environ["PATH_INFO"] == "/api/library/search":
            gate["active"] = False
            gate["started"].set()
            gate["release"].wait(10)
        return app(environ, start_response)

    server = make_server(
        "127.0.0.1", 0, gated_app, server_class=Server, handler_class=Handler
    )
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    def upload(name, raw, scoped=False):
        code, body, _ = request(
            app,
            "POST",
            "/api/uploads",
            {"workbook": {"name": name, "size": len(raw)}, "searchScoped": scoped},
        )
        assert code == 200, body
        scope = json.loads(body)["upload"]["id"]
        code, body, _ = request(
            app,
            "PUT",
            f"/api/uploads/{scope}/files/0",
            CONTENT_TYPE="application/octet-stream",
            CONTENT_LENGTH=str(len(raw)),
            HTTP_X_UPLOAD_OFFSET="0",
            **{"wsgi.input": io.BytesIO(raw)},
        )
        assert code == 200, body
        code, body, _ = request(app, "POST", f"/api/uploads/{scope}/complete", {})
        assert code == 200, body
        result = json.loads(body)
        assert wait(app.store, result["task"]["id"])["status"] == "succeeded"
        return result["project"]

    yield app, f"http://127.0.0.1:{server.server_port}", upload
    gate["release"].set()
    server.shutdown()
    thread.join(5)
    server.server_close()
    app.close()


def open_search(browser, url, width):
    page = browser.new_page(
        viewport={"width": width, "height": 844 if width == 390 else 900}
    )
    page.set_default_timeout(10000)
    page.goto(url)
    page.locator("#connection[data-state=ok]").wait_for(state="attached")
    page.locator(".library-search-trigger").click()
    return page


def scope_ready(page, url, scope=None):
    for _ in range(200):
        response = page.request.get(
            url + "/api/library/status" + (f"?scope={scope}" if scope else "")
        )
        assert response.ok, response.text()
        status = response.json()
        if status["buildState"]["status"] == "ready" and not status["stale"]:
            return status
        time.sleep(0.02)
    raise AssertionError("Catalog did not reach ready/fresh")


def search(page, query):
    page.locator("#library-search-query").fill(query)
    with page.expect_response(
        lambda response: "/api/library/search?" in response.url
    ) as response:
        page.locator(".library-search-form-row button[type=submit]").click()
    assert response.value.ok, response.value.text()
    return response.value.json()


def select(page, address):
    with page.expect_response(
        lambda response: "/api/library/lineage?" in response.url
    ) as response:
        page.locator(".library-occurrences button").filter(
            has_text="!" + address
        ).first.click()
    assert response.value.ok, response.value.text()
    page.locator(".library-verdict").wait_for()
    return response.value.json()


@pytest.mark.parametrize("width", [1440, 390])
def test_superseding_real_search_unlocks_controls_and_settles_trace(
    browser, live_library, width
):
    app, url, upload = live_library
    upload("alpha.xlsx", book_bytes())
    page = open_search(browser, url, width)
    try:
        scope_ready(page, url)
        gate = app.test_search_gate
        gate["active"] = True
        page.locator("#library-search-query").fill("ALPHA")
        page.locator(".library-search-form-row button[type=submit]").click()
        assert gate["started"].wait(5)
        page.locator(".library-trace-summary").filter(
            has_text="Recherche en cours"
        ).wait_for()
        assert (
            "Recherche terminée"
            not in page.locator(".library-trace-summary").inner_text()
        )
        page.locator("#library-search-query").fill("capneedle")
        playwright.expect(
            page.locator(".library-search-form-row button[type=submit]")
        ).to_be_enabled()
        playwright.expect(page.locator("#library-attach-file")).to_be_enabled()
        assert "remplacée" in page.locator(".library-trace-summary").inner_text()
        gate["release"].set()
        result = search(page, "capneedle")
        assert result["results"] and not result["coverage"]["complete"]
        assert page.locator(".library-result").count() <= 10
        page.locator(".library-trace-summary a").click()
        assert "Recherche · remplacé" in page.locator(".library-trace").inner_text()
        assert "Recherche · terminé" in page.locator(".library-trace").inner_text()
    finally:
        page.close()


@pytest.mark.parametrize("width", [1440, 390])
@pytest.mark.parametrize("inflight", [False, True])
def test_recovery_changes_scope_and_invalidates_completed_or_pending_answer(
    browser, live_library, width, inflight
):
    app, url, upload = live_library
    first = upload("alpha.xlsx", book_bytes(), True)
    second = upload("beta.xlsx", book_bytes("BETA", 9), True)
    page = open_search(browser, url, width)
    try:
        page.locator(".library-advanced > summary").click()
        recover = page.locator(".library-recoverable-attachments button").filter(
            has_text="Récupérer"
        )
        recover.filter(has_text=first["id"][:8]).click()
        scope_ready(page, url, first["searchScopeId"])
        assert search(page, "ALPHA")["results"]
        select(page, "A2")
        if inflight:
            gate = app.test_search_gate
            gate["active"] = True
            page.locator(".library-search-form-row button[type=submit]").click()
            assert gate["started"].wait(5)
            page.locator(".library-trace-summary").filter(
                has_text="Recherche en cours"
            ).wait_for()
        recover.filter(has_text=second["id"][:8]).click()
        assert page.locator(".library-result").count() == 0
        assert page.locator(".library-verdict").count() == 0
        playwright.expect(
            page.locator(".library-search-form-row button[type=submit]")
        ).to_be_enabled()
        if inflight:
            gate["release"].set()
        scope_ready(page, url, second["searchScopeId"])
        result = search(page, "BETA")
        assert result["results"]
        assert {hit["revision"]["projectId"] for hit in result["results"]} == {
            second["id"]
        }
    finally:
        page.close()


@pytest.mark.parametrize("width", [1440, 390])
def test_group_membership_counts_null_evidence_and_current_pin_confirmation(
    browser, live_library, width
):
    _, url, upload = live_library
    target = upload("target.xlsx", linked_workbook({"A1": "Seed", "B1": 7}, [], "S"))
    reader = upload(
        "reader.xlsx",
        linked_workbook(
            {"A1": "ROOT", "B1": "='[1]S'!B1"},
            [_link("target.xlsx", "S", {"B1": 7})],
            "S",
        ),
    )
    page = open_search(browser, url, width)
    try:
        scope_ready(page, url)
        links = page.request.get(url + "/api/library/links").json()["links"]
        link = links[0]
        response = page.request.post(
            url + "/api/library/pins",
            data={
                "linkId": link["id"],
                "targetProject": target["id"],
                "targetSha256": target["workbookRevision"]["sha256"],
            },
        )
        assert response.ok, response.text()
        page.reload()
        page.locator(".library-search-trigger").click()
        card = page.locator(".library-link").filter(has=page.locator("select"))
        card.get_by_role("button", name="Confirmer cette cible", exact=True).wait_for()
        assert card.locator("select").input_value() == target["workbookRevision"]["id"]
        with page.expect_request(
            lambda request: (
                request.method == "POST" and request.url.endswith("/api/library/pins")
            )
        ) as pin:
            card.get_by_role("button", name="Confirmer cette cible", exact=True).click()
        assert pin.value.post_data_json["targetProject"] == target["id"]
        assert page.locator("#library-search-view [role=alert]").count() == 0
        search(page, "ROOT")
        lineage = select(page, "B1")
        assert len(lineage["nodes"]) == 2
        page.locator(".library-graph-canvas .react-flow__node").first.wait_for()
        model = page.locator(".library-graph-canvas").evaluate(
            "e => ({nodes:e._flow.getNodes(),edges:e._flow.getEdges()})"
        )
        assert len(model["nodes"]) == 2
        assert {
            member["id"]
            for node in model["nodes"]
            for member in node["data"]["members"]
        } == {node["id"] for node in lineage["nodes"]}
        assert len(model["edges"]) == 1 and not model["edges"][0].get("hidden")
        assert (
            "0 lien synthétique" in page.locator(".library-graph-counts").inner_text()
        )
        page.locator(".library-graph-canvas").scroll_into_view_if_needed()
        page.wait_for_timeout(100)
        bounds = page.locator(".library-graph-canvas").evaluate("""e => {
              const c=e.getBoundingClientRect();
              return [...e.querySelectorAll('.react-flow__node')].map(n=>{
                const b=n.getBoundingClientRect();
                return {left:b.left-c.left,right:b.right-c.left,
                  top:b.top-c.top,bottom:b.bottom-c.top,width:b.width,
                  cw:c.width,ch:c.height};
              });
        }""")
        assert all(
            0 <= b["left"]
            and b["right"] <= b["cw"]
            and 0 <= b["top"]
            and b["bottom"] <= b["ch"]
            and b["width"] >= 130
            for b in bounds
        )
        label_bounds = page.locator(".library-graph-canvas").evaluate("""e => {
              const c=e.getBoundingClientRect();
              return [...e.querySelectorAll('.react-flow__edge-textbg')].map(n=>{
                const b=n.getBoundingClientRect();
                return {left:b.left-c.left,right:b.right-c.left,
                  top:b.top-c.top,bottom:b.bottom-c.top,cw:c.width,ch:c.height};
              });
        }""")
        assert label_bounds and all(
            0 <= b["left"] <= b["right"] <= b["cw"]
            and 0 <= b["top"] <= b["bottom"] <= b["ch"]
            for b in label_bounds
        )
        page.locator(".library-revision-navigation button").filter(
            has_text="reader.xlsx"
        ).click()
        playwright.expect(
            page.locator(".library-revision-navigation button[aria-pressed=true]")
        ).to_contain_text("reader.xlsx")
        assert (
            "aucune valeur enregistrée"
            in page.locator(".library-accessible-graph").inner_text()
        )
        assert (
            "cache enregistré" in page.locator(".library-accessible-graph").inner_text()
        )
        assert (
            "Formule\u00a0:" in page.locator(".library-accessible-graph").inner_text()
        )
        assert (
            reader["workbookRevision"]["sha256"][:8]
            in page.locator(".library-accessible-graph").inner_text()
        )
    finally:
        page.close()


@pytest.mark.parametrize("width", [1440, 390])
def test_initial_trace_access_keyboard_anchors_and_ready_fresh_poll(
    browser, live_library, width
):
    _, url, upload = live_library
    upload("alpha.xlsx", book_bytes())
    page = open_search(browser, url, width)
    try:
        scope_ready(page, url)
        for query in ["ALPHA", "absentneedle", "capneedle"]:
            search(page, query)
            page.locator(".workspace > main").evaluate("e => e.scrollTop=0")
            box = page.locator(".library-trace-summary a").bounding_box()
            assert (
                box
                and box["y"] >= 48
                and box["y"] + box["height"] <= (844 if width == 390 else 900)
            )
            page.locator(".library-trace-summary a").focus()
            page.keyboard.press("Enter")
            playwright.expect(page.locator("#library-trace-heading")).to_be_focused()
            assert (
                page.locator("#library-trace-heading").bounding_box()["y"]
                >= page.locator(".library-trace-summary").bounding_box()["y"]
                + page.locator(".library-trace-summary").bounding_box()["height"]
            )
        calls = []

        def status_sequence(route):
            response = route.fetch()
            data = response.json()
            calls.append(data)
            if len(calls) < 3:
                data["buildState"] = {"status": "ready"}
                data["stale"] = True
            route.fulfill(json=data)

        page.route("**/api/library/status", status_sequence)
        page.reload()
        page.locator(".library-search-trigger").click()
        playwright.expect(page.locator(".library-catalog-state")).to_contain_text(
            "Mise à jour en attente"
        )
        playwright.expect(page.locator(".library-catalog-state")).not_to_contain_text(
            "À jour"
        )
        page.wait_for_function(
            """() => {
              const text=document.querySelector('.library-catalog-state').innerText;
              return text.includes('À jour') && !text.includes('pendant la mise');
            }"""
        )
        assert len(calls) >= 3
        playwright.expect(page.locator(".library-catalog-state")).not_to_contain_text(
            "Mise à jour en attente"
        )
        settled = len(calls)
        page.wait_for_timeout(1100)
        assert len(calls) == settled
    finally:
        page.close()
