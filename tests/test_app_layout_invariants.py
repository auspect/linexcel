"""Layout invariants of the application, on a real imported workbook.

At 1440x900, 824x768 and 390x844, for the sidebar, the open import form, each
settings section and each main tab, every visible button, link, field, select,
summary and tab can be hit at its centre: the element found there is the
control or one of its descendants. The document never scrolls sideways. On a
phone every such control is at least 44 px wide and 44 px tall.
"""

import base64
import json
import os
import threading
from socketserver import ThreadingMixIn
from wsgiref.simple_server import WSGIRequestHandler, WSGIServer, make_server

import pytest
from conftest import build_lineage_workbook
from graph_test_helpers import open_workspace_menu
from test_web import request, wait

from linexcel.web import create_app

playwright = pytest.importorskip("playwright.sync_api")

VIEWPORTS = {"desktop": (1440, 900), "tablet": (824, 768), "phone": (390, 844)}
TABS = ["nodes", "graph", "sheets", "data", "captures", "tasks", "diagnostics"]
SECTIONS = ["limits", "ai", "backup", "health"]

# (viewport, state, control) -> reason. One line each.
EXCEPTIONS = {
    ("phone", "tab graph", 'a "React Flow attribution"'): "present before this change",
}

# Visible controls of a scope, each scrolled into view and probed at its centre.
PROBE = """([scope, phone]) => {
  const root = document.querySelector(scope);
  const selector = 'button,a[href],input:not([type=hidden]):not([type=file]),'
    + 'select,textarea,summary,[role=tab]';
  const name = e => {
    if (e.id) return e.tagName.toLowerCase() + '#' + e.id;
    const label = (e.getAttribute('aria-label') || e.textContent || e.value || '')
      .trim().replace(/\\s+/g, ' ').slice(0, 40);
    const kind = [...e.classList].slice(0, 2).map(c => '.' + c).join('');
    return e.tagName.toLowerCase() + kind + ' "' + label + '"';
  };
  const problems = [];
  let checked = 0;
  for (const control of root.querySelectorAll(selector)) {
    if (!control.checkVisibility({opacityProperty: true, visibilityProperty: true}))
      continue;
    let box = control.getBoundingClientRect();
    // Visually hidden on purpose (skip links, screen-reader labels).
    if (box.width <= 2 || box.height <= 2) continue;
    checked += 1;
    control.scrollIntoView({block: 'center', inline: 'center'});
    box = control.getBoundingClientRect();
    const x = box.left + box.width / 2, y = box.top + box.height / 2;
    const found = document.elementFromPoint(x, y);
    if (!found || !(found === control || control.contains(found)))
      problems.push([name(control), 'covered by ' + (found ? name(found) : 'nothing')]);
    if (phone && (box.width < 44 || box.height < 44))
      problems.push([name(control),
        'smaller than 44 px: ' + Math.round(box.width) + 'x' + Math.round(box.height)]);
  }
  const overflow = document.documentElement.scrollWidth - innerWidth;
  if (overflow > 0)
    problems.push(['document', 'overflows sideways by ' + overflow + ' px']);
  return {problems, checked};
}"""


@pytest.fixture(scope="module")
def browser():
    with playwright.sync_playwright() as driver:
        instance = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        yield instance
        instance.close()


@pytest.fixture(scope="module")
def live_app(tmp_path_factory):
    app = create_app(
        tmp_path_factory.mktemp("layout") / "data", resolve_user=lambda env: "alice"
    )

    class Server(ThreadingMixIn, WSGIServer):
        daemon_threads = True

    class Handler(WSGIRequestHandler):
        def log_message(self, *args):
            pass

    server = make_server(
        "127.0.0.1", 0, app, server_class=Server, handler_class=Handler
    )
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    workbook = {
        "name": "lineage.xlsx",
        "data": base64.b64encode(build_lineage_workbook()).decode(),
    }
    code, body, _ = request(app, "POST", "/api/projects", {"workbook": workbook})
    assert code == 200, body
    imported = json.loads(body)
    assert wait(app.store, imported["task"]["id"])["status"] == "succeeded"
    yield f"http://127.0.0.1:{server.server_port}", imported["project"]["id"]
    server.shutdown()
    thread.join(5)
    server.server_close()
    app.close()


def _page(browser, live_app, viewport):
    url, project = live_app
    width, height = VIEWPORTS[viewport]
    page = browser.new_page(viewport={"width": width, "height": height})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    page.goto(f"{url}/?project={project}")
    page.locator("#project-title").get_by_text("lineage").wait_for()
    page.locator("#connection[data-state='ok']").wait_for(state="attached")
    return page, errors


def _settle(page):
    page.evaluate(
        "() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))"
    )
    page.wait_for_timeout(250)  # Transitions of 150-240 ms end before probing.


def _problems(page, viewport, state, scope):
    _settle(page)
    found = page.evaluate(PROBE, [scope, viewport == "phone"])
    assert found["checked"] >= 3, (viewport, state)  # The probe saw the state.
    return {
        (viewport, state, control): problem
        for control, problem in found["problems"]
        if (viewport, state, control) not in EXCEPTIONS
    }


def _close_menu(page):
    close = page.locator("#workspace-close")
    if close.is_visible():
        close.click()


@pytest.mark.parametrize("viewport", list(VIEWPORTS))
def test_sidebar_import_form_and_settings_controls_can_be_hit(
    browser, live_app, viewport
):
    page, errors = _page(browser, live_app, viewport)
    try:
        problems = {}
        open_workspace_menu(page)
        problems |= _problems(page, viewport, "sidebar", "#workspace-sidebar")
        page.locator("#import-panel > summary").click()
        page.locator("#limits-edit").wait_for()
        problems |= _problems(page, viewport, "import form", "#workspace-sidebar")
        page.locator("#settings-open").click()
        for section in SECTIONS:
            page.locator(f"#settings [data-settings='{section}']").click()
            page.locator(f"#settings-{section}").wait_for()
            if section == "health":
                page.locator("#health-check").click()
                page.locator("#health-status p").first.wait_for()
            problems |= _problems(page, viewport, f"settings {section}", "#settings")
        assert problems == {}
        assert errors == []
    finally:
        page.close()


@pytest.mark.parametrize("viewport", list(VIEWPORTS))
def test_every_main_tab_of_a_workbook_keeps_its_controls_reachable(
    browser, live_app, viewport
):
    page, errors = _page(browser, live_app, viewport)
    try:
        problems = {}
        _close_menu(page)
        for tab in TABS:
            page.locator(f"#tab-{tab}").click()
            page.locator(f"#view-{tab}").wait_for()
            if tab == "graph":
                page.locator("#graph-canvas .react-flow__node").first.wait_for()
            problems |= _problems(page, viewport, f"tab {tab}", "body")
        assert problems == {}
        assert errors == []
    finally:
        page.close()
