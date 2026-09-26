"""Runtime regressions for graph interaction and responsive report rendering.

Run with ``uv sync --extra screenshots``, ``uv run playwright install chromium``
and ``uv run pytest tests/test_viewer_browser.py``. Set
``LINEXCEL_BROWSER_CHANNEL=msedge`` to use an installed Edge on Windows.
"""

import os
from pathlib import Path

import pytest

from linexcel.viewer import render_html

playwright = pytest.importorskip("playwright.sync_api")


@pytest.fixture(scope="module")
def browser():
    with playwright.sync_playwright() as driver:
        browser = driver.chromium.launch(
            channel=os.environ.get("LINEXCEL_BROWSER_CHANNEL") or None
        )
        yield browser
        browser.close()


@pytest.fixture
def report(browser, tmp_path):
    pages = []

    def open_report(graph, *, width=1440, height=900, language="en"):
        graph = {
            **graph,
            "meta": {
                "stats": {
                    "totalFormulas": 0,
                    "totalNodes": len(graph["nodes"]),
                    "totalEdges": len(graph["edges"]),
                },
                **graph.get("meta", {}),
            },
        }
        page = browser.new_page(viewport={"width": width, "height": height})
        pages.append(page)
        errors = []
        page.on("pageerror", lambda error: errors.append(str(error)))
        path = tmp_path / f"report-{len(pages)}.html"
        path.write_text(render_html(graph, language=language), encoding="utf-8")
        page.goto(path.as_uri())
        page.wait_for_function("document.querySelector('#lin-cy')._cyreg?.cy != null")
        page.evaluate(
            "() => { window.cy = document.querySelector('#lin-cy')._cyreg.cy; }"
        )
        page.wait_for_timeout(100)
        return page, errors

    yield open_report
    for page in pages:
        page.close()


def filtered_graph():
    return {
        "nodes": [
            {"id": "i", "kind": "input", "sheet": "In", "label": "Input"},
            {
                "id": "c",
                "kind": "cell",
                "sheet": "Out",
                "label": "Output",
                "formula": "=In!A1",
                "value": 2,
                "cachedValue": 1,
                "valueSource": "engine",
                "cachedAgreement": "differ",
            },
            {"id": "u", "kind": "input", "sheet": "Other", "label": "Unrelated"},
        ],
        "edges": [{"id": "e", "source": "i", "target": "c", "kind": "ref"}],
    }


def visible_ids(page):
    return page.evaluate("cy.nodes(':visible').map(n => n.id()).sort()")


def capture(page, name):
    directory = os.environ.get("LINEXCEL_VIEWER_ARTIFACTS")
    if directory:
        target = Path(directory)
        target.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(target / f"{name}.png"))


def test_sheet_kind_and_discrepancy_filters_compose(report):
    page, errors = report(filtered_graph())
    page.locator("#lin-rail-kinds button").filter(has_text="Source data").click()
    assert visible_ids(page) == ["c"]
    page.select_option("#lin-sheet-filter", "Out")
    assert visible_ids(page) == ["c"]
    page.select_option("#lin-sheet-filter", "__all__")
    assert visible_ids(page) == ["c"]
    page.locator("#lin-rail-kinds button").filter(has_text="Source data").click()
    page.click("#lin-diffs-only")
    assert visible_ids(page) == ["c"]
    page.select_option("#lin-sheet-filter", "Other")
    assert visible_ids(page) == []
    page.select_option("#lin-sheet-filter", "__all__")
    page.wait_for_function("cy.getElementById('c').visible()", timeout=3000)
    assert visible_ids(page) == ["c"]
    assert not errors


def test_search_does_not_select_hidden_results(report):
    page, errors = report(filtered_graph())
    page.select_option("#lin-sheet-filter", "Out")
    page.fill("#lin-search", "Unrelated")
    page.press("#lin-search", "Enter")
    assert page.locator("#lin-search-status").inner_text() == "No matches"
    assert page.evaluate("cy.nodes(':selected').length") == 0
    assert not errors


def test_portrait_flow_stays_readable_and_respects_explicit_layout(report):
    graph = {
        "nodes": [
            {"id": "input1", "label": "Inputs!A1:A100", "kind": "input"},
            {"id": "input2", "label": "Inputs!B1:B100", "kind": "input"},
            {"id": "parameter", "label": "Params!C4", "kind": "input"},
            {"id": "name", "label": "TargetRate", "kind": "name"},
            {"id": "group", "label": "Sales!E4", "kind": "group", "count": 100},
            {"id": "total", "label": "Summary!C5", "kind": "cell"},
            {"id": "mean", "label": "Summary!C6", "kind": "cell"},
            {"id": "status", "label": "Summary!C7", "kind": "cell"},
        ],
        "edges": [
            {"id": f"edge{i}", "source": source, "target": target, "kind": "ref"}
            for i, (source, target) in enumerate(
                [
                    ("input1", "group"),
                    ("input2", "group"),
                    ("group", "total"),
                    ("group", "mean"),
                    ("group", "status"),
                    ("total", "status"),
                    ("parameter", "name"),
                    ("name", "status"),
                ]
            )
        ],
    }
    page, errors = report(graph, width=390, height=844)
    assert page.locator("#lin-lay-dagre").get_attribute("aria-pressed") == "true"
    metrics = page.evaluate("""() => ({
        pixels:parseFloat(cy.nodes().first().style('font-size'))*cy.zoom(),
        before:cy.getElementById('input1').position('y'),
        after:cy.getElementById('group').position('y'),
        framed:cy.nodes().every(n=>{
            const b=n.renderedBoundingBox();
            return b.x1>=0&&b.y1>=0&&b.x2<=cy.width()&&b.y2<=cy.height();
        }),edges:cy.edges(':visible').length})""")
    assert metrics["pixels"] >= 10
    assert metrics["before"] < metrics["after"]
    assert metrics["framed"] and metrics["edges"] == 8
    page.click("#lin-options-toggle")
    page.click("#lin-lay-fcose")
    page.keyboard.press("Escape")
    page.set_viewport_size({"width": 844, "height": 390})
    page.set_viewport_size({"width": 390, "height": 844})
    assert page.locator("#lin-lay-fcose").get_attribute("aria-pressed") == "true"
    assert not errors


def test_readme_capture_selects_result_before_photographing(report):
    import runpy

    capture = runpy.run_path(
        str(Path(__file__).resolve().parents[1] / "scripts" / "capture_viewer.py")
    )

    graph = filtered_graph()
    graph["nodes"][1]["label"] = "Summary!C5"
    page, errors = report(graph)
    assert capture["apply_shot"](page, capture["SHOTS"][1])
    assert page.locator("#lin-search-results").is_hidden()
    assert page.locator("#lin-panel h2").inner_text() == "Summary!C5"
    assert not errors


def test_dense_mobile_local_view_starts_with_three_readable_neighbors(report):
    graph = {
        "nodes": [{"id": "hub", "label": "Lookups!B21", "kind": "cell", "value": 10}]
        + [
            {"id": str(i), "label": f"'Sales Data'!E{i}", "kind": "input", "value": i}
            for i in range(100)
        ],
        "edges": [
            {"id": f"edge{i}", "source": str(i), "target": "hub", "kind": "ref"}
            for i in range(100)
        ],
    }
    page, errors = report(graph, width=390, height=844)
    page.evaluate("() => {cy.getElementById('hub').emit('tap');}")
    page.click("#lin-explore")
    page.wait_for_function("cy.nodes(':visible').length === 4")
    assert (
        page.evaluate("parseFloat(cy.nodes().first().style('font-size'))*cy.zoom()")
        >= 11
    )
    assert "4/101" in page.locator("#lin-explore-status").inner_text()
    page.click("#lin-explore-more")
    page.wait_for_function("cy.nodes(':visible').length === 7")
    assert "7/101" in page.locator("#lin-explore-status").inner_text()
    assert not errors


def test_group_discrepancy_preserves_the_representative_reading(report):
    group = {
        "id": "g",
        "label": "Group",
        "kind": "group",
        "formula": "=B1*2",
        "r1c1": "=RC[1]*2",
        "count": 2,
        "bbox": "A1:A2",
        "value": 4,
        "cachedValue": 4,
        "valueSource": "engine",
        "cachedAgreement": "same",
        "groupCachedAgreement": "differ",
        "samples": [
            {
                "addr": "A1",
                "value": 4,
                "cachedValue": 4,
                "source": "engine",
                "cachedAgreement": "same",
            },
            {
                "addr": "A2",
                "value": 6,
                "cachedValue": 4,
                "source": "engine",
                "cachedAgreement": "differ",
            },
        ],
    }
    page, errors = report({"nodes": [group], "edges": []})
    page.click("#lin-diffs-only")
    assert visible_ids(page) == ["g"]
    page.fill("#lin-search", "Group")
    page.press("#lin-search", "Enter")
    assert "differs from the file" in page.locator(".lin-verdict").inner_text()
    assert page.locator(".lin-vtable tbody tr").count() == 2
    assert page.locator(".lin-vtable tr.is-diff th").inner_text() == "A2"
    assert page.evaluate("cy.nodes(':selected').length") == 1
    assert not errors


def test_dense_search_updates_detail_once_per_frame(report):
    graph = {
        "nodes": [
            {"id": str(i), "kind": "cell", "label": f"Match {i}"} for i in range(400)
        ],
        "edges": [],
    }
    page, errors = report(graph)
    aspect = page.evaluate("""() => {
        const box = cy.nodes().boundingBox();
        return box.w / box.h;
    }""")
    assert 0.2 < aspect < 5, "Isolated cards must fill an area, not one thin rank"
    # Count graph-wide scans during the synchronous selection transaction.
    # The ceiling is independent of node count; timing thresholds vary by CI.
    calls = page.evaluate("""() => {
        const original = cy.nodes;
        let calls = 0;
        cy.nodes = function (...args) { calls++; return original.apply(this, args); };
        const input = document.querySelector('#lin-search');
        input.value = 'Match';
        input.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'Enter', bubbles: true
        }));
        cy.nodes = original;
        return calls;
    }""")
    assert calls < 20
    assert page.evaluate("cy.nodes(':selected').length") == 400
    page.wait_for_timeout(600)
    capture(page, "dense-search")
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
@pytest.mark.parametrize("dark", [False, True])
def test_report_controls_fit_viewport(report, width, height, dark):
    page, errors = report(filtered_graph(), width=width, height=height)
    if dark:
        page.click("#lin-theme")
    page.fill("#lin-search", "Output")
    page.press("#lin-search", "Enter")
    page.wait_for_timeout(100)
    assert page.evaluate("cy.getElementById('c').style('opacity')") == "1"
    for selector in ("#lin-search", "#lin-theme", ".lin-close"):
        box = page.locator(selector).bounding_box()
        assert box is not None
        assert box["x"] >= 0 and box["x"] + box["width"] <= width
    page.wait_for_timeout(600)
    capture(page, f"report-{width}-{'dark' if dark else 'light'}")
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_neighbor_action_fits_the_visible_neighborhood(report, width, height):
    graph = {
        "nodes": [{"id": "hub", "kind": "cell", "label": "Hub"}]
        + [{"id": str(i), "kind": "input", "label": f"Input {i}"} for i in range(80)],
        "edges": [
            {"id": f"e{i}", "source": str(i), "target": "hub", "kind": "ref"}
            for i in range(80)
        ],
    }
    page, errors = report(graph, width=width, height=height)
    if width <= 900:
        page.click("#lin-options-toggle")
    page.click("#lin-lay-dagre")
    if width <= 900:
        page.click("#lin-options-close")
    page.evaluate("() => { cy.getElementById('hub').emit('tap'); }")
    page.click("#lin-fit-neighbors")
    page.wait_for_function("""() => {
        const rect = document.querySelector('#lin-cy').getBoundingClientRect();
        const box = cy.nodes(':visible').renderedBoundingBox();
        return box.x1 >= 0 && box.y1 >= 0
            && box.x2 <= rect.width && box.y2 <= rect.height;
    }""")
    # A mobile reader can also reach the original canvas controls while the
    # details are open. The old full-height overlay intercepted this click.
    page.click("#lin-fit-sel")
    assert page.evaluate("cy.getElementById('hub').selected()")
    assert not errors


@pytest.mark.parametrize(
    "value,evaluated,state",
    [
        ("#VALUE!", True, "is-error"),
        ({"type": "Error", "kind": "Div"}, True, "is-error"),
        (None, False, "is-pending"),
    ],
)
def test_failed_or_pending_final_step_does_not_signal_success(
    report, value, evaluated, state
):
    graph = filtered_graph()
    graph["nodes"][1]["steps"] = {
        "label": "SUM",
        "expr": "SUM(A1)",
        "value": value,
        "evaluated": evaluated,
        "inputs": [{"ref": "A1", "value": value}],
    }
    page, errors = report(graph)
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    assert state in page.locator(".lin-final").get_attribute("class")
    assert (
        page.locator(".lin-final").evaluate("e => getComputedStyle(e).borderLeftColor")
        != "rgb(27, 175, 122)"
    )
    assert not errors


def rich_graph():
    graph = filtered_graph()
    graph["nodes"][1]["doc"] = "## Synthetic cell documentation\nReads the input."
    image = (
        "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' "
        "width='120' height='80'><rect width='120' height='80' fill='gray'/></svg>"
    )
    graph["meta"] = {
        "workbookDoc": (
            "# Synthetic overview\n\n| Item | Value |\n|---|---|\n| Test | 2 |"
        ),
        "workbookContext": {
            "sheets": [
                {
                    "name": "Out",
                    "visibility": "visible",
                    "dimensions": {"rows": 1, "columns": 2},
                    "preview_range": "A1:B1",
                    "preview": [{"row": 1, "values": ["Test", 2]}],
                }
            ]
        },
        "screenshots": [image, image.replace("gray", "blue")],
    }
    return graph


def test_workbook_search_reveals_hidden_result_only_when_chosen(report):
    page, errors = report(filtered_graph())
    page.select_option("#lin-sheet-filter", "Out")
    page.fill("#lin-search", "Unrelated")
    page.press("#lin-search", "Enter")
    assert page.locator("#lin-search-status").inner_text() == "No matches"
    page.select_option("#lin-search-scope", "workbook")
    assert visible_ids(page) == ["c", "i"]
    assert "reset filters" in page.locator("#lin-result-list").inner_text()
    page.press("#lin-search", "ArrowDown")
    assert page.evaluate("document.activeElement.dataset.nodeId") == "u"
    page.keyboard.press("Enter")
    assert page.locator("#lin-sheet-filter").input_value() == "__all__"
    assert page.evaluate("cy.nodes(':selected').map(n => n.id())") == ["u"]
    assert page.locator("#lin-search-results").is_hidden()
    assert not errors


def test_search_results_keyboard_and_safe_formula_excerpt(report):
    graph = filtered_graph()
    graph["nodes"][0]["formula"] = "=<img onerror=alert(1)>"
    page, errors = report(graph)
    page.fill("#lin-search", "")
    page.click("#lin-search-submit")
    page.select_option("#lin-search-scope", "workbook")
    page.fill("#lin-search", "i")
    page.press("#lin-search", "Enter")
    assert page.locator("#lin-result-list img").count() == 0
    page.press("#lin-search", "ArrowDown")
    page.keyboard.press("End")
    assert page.evaluate("document.activeElement.dataset.nodeId") == "c"
    page.keyboard.press("Home")
    assert page.evaluate("document.activeElement.dataset.nodeId") == "i"
    page.keyboard.press("Escape")
    assert page.locator("#lin-search-results").is_hidden()
    assert page.locator("#lin-search").evaluate("e => e === document.activeElement")
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_sheet_roundtrip_preserves_graph_and_hides_graph_only_controls(
    report, width, height
):
    page, errors = report(rich_graph(), width=width, height=height)
    page.select_option("#lin-sheet-filter", "Out")
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    page.wait_for_timeout(100)
    page.evaluate("() => { cy.stop(); cy.zoom(.8); cy.pan({x: 42, y: 63}); }")
    before = page.evaluate("({zoom: cy.zoom(), pan: cy.pan()})")
    page.click("#lin-see-sheet")
    assert "Output" in page.locator("#lin-back-graph").inner_text()
    assert page.locator("#lin-options-host").is_hidden()
    assert page.locator("#lin-tools").is_hidden()
    page.click("#lin-back-graph")
    assert page.locator("#lin-sheet-filter").input_value() == "Out"
    assert page.evaluate("({zoom: cy.zoom(), pan: cy.pan()})") == before
    assert page.evaluate("cy.nodes(':selected').map(n => n.id())") == ["c"]
    capture(page, f"investigation-{width}")
    assert not errors


def test_history_restores_filters_and_camera_after_revealing_result(report):
    page, errors = report(filtered_graph())
    page.select_option("#lin-sheet-filter", "Out")
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    page.evaluate("() => { cy.stop(); cy.zoom(.8); cy.pan({x: 42, y: 63}); }")
    before = page.evaluate("({zoom: cy.zoom(), pan: cy.pan()})")
    page.fill("#lin-search", "Unrelated")
    page.select_option("#lin-search-scope", "workbook")
    page.locator("#lin-result-list button").click()
    page.click("#lin-history-back")
    assert page.locator("#lin-sheet-filter").input_value() == "Out"
    assert page.evaluate("({zoom: cy.zoom(), pan: cy.pan()})") == before
    assert page.locator("#lin-panel h2").inner_text() == "Output"
    page.click("#lin-history-forward")
    assert page.locator("#lin-panel h2").inner_text() == "Unrelated"
    assert page.locator("#lin-sheet-filter").input_value() == "__all__"
    assert not errors


def test_copy_formula_and_local_selection_link_are_exact(report):
    graph = filtered_graph()
    formula = '=LET(x,"<tag>&value", A1*B1 + C1*D1)'
    graph["nodes"][1]["formula"] = formula
    page, errors = report(graph)
    page.evaluate("""() => {
      window.copied = [];
      Object.defineProperty(navigator, 'clipboard', {value: {
        writeText: value => { window.copied.push(value); return Promise.resolve(); }
      }});
      cy.getElementById('c').emit('tap');
    }""")
    page.locator("#lin-panel details summary").click()
    page.click("#lin-copy-formula")
    page.click("#lin-copy-link")
    values = page.evaluate("window.copied")
    assert values[0] == formula
    assert values[1].endswith("#node=c")
    page.goto(values[1])
    page.reload()
    page.wait_for_function(
        "document.querySelector('#lin-panel h2')?.textContent === 'Output'"
    )
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_progressive_exploration_bounds_hub_and_restores_layout(report, width, height):
    graph = {
        "nodes": [{"id": "hub", "label": "Hub", "kind": "cell", "sheet": "S"}]
        + [
            {"id": f"n{i:03}", "label": f"Input {i}", "kind": "input", "sheet": "S"}
            for i in range(40)
        ],
        "edges": [
            {"id": f"e{i}", "source": f"n{i:03}", "target": "hub", "kind": "ref"}
            for i in range(40)
        ],
    }
    page, errors = report(graph, width=width, height=height)
    page.evaluate("() => { cy.getElementById('hub').emit('tap'); cy.stop(); }")
    before = page.evaluate("cy.nodes().map(n => ({id:n.id(), p:n.position()}))")
    page.click("#lin-explore")
    initial_count = 7 if width < 900 else 13
    page.wait_for_function("n => cy.nodes(':visible').length === n", arg=initial_count)
    assert len(visible_ids(page)) == initial_count
    assert f"{initial_count}/41" in page.locator("#lin-explore-status").inner_text()
    capture(page, f"exploration-initial-{width}")
    page.click("#lin-explore-more")
    page.wait_for_function(
        "n => cy.nodes(':visible').length === n", arg=13 if width < 900 else 25
    )
    assert len(visible_ids(page)) == (13 if width < 900 else 25)
    capture(page, f"exploration-{width}")
    page.click("#lin-explore-exit")
    page.wait_for_function("cy.nodes(':visible').length === 41")
    assert len(visible_ids(page)) == 41
    assert page.evaluate("cy.nodes().map(n => ({id:n.id(), p:n.position()}))") == before
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_coverage_counts_graph_nodes_and_opens_drilldown(report, width, height):
    graph = filtered_graph()
    graph["meta"] = {
        "coverage": {
            "scope": "graph_nodes",
            "totalNodes": 3,
            "categories": {
                "engine": {"count": 1, "nodeIds": ["c"]},
                "other": {"count": 2, "nodeIds": ["i", "u"]},
                "divergent": {"count": 1, "nodeIds": ["c"]},
            },
            "omissions": {
                "status": "not_certified",
                "warnings": ["Trace incomplete: source omitted"],
            },
        }
    }
    page, errors = report(graph, width=width, height=height)
    if width < 900:
        page.click("#lin-options-toggle")
    assert "groups count as one node" in page.locator("#lin-coverage").inner_text()
    assert "do not certify" in page.locator("#lin-coverage").inner_text()
    page.locator("#lin-coverage summary").click()
    assert "Trace incomplete" in page.locator("#lin-coverage").inner_text()
    page.locator('[data-coverage="divergent"]').click()
    page.wait_for_function("document.activeElement.id === 'lin-panel'")
    page.wait_for_function("""() => {
        const b=cy.getElementById('c').renderedBoundingBox();
        return b.x1>=0 && b.y1>=0 && b.x2<=cy.width() && b.y2<=cy.height();
    }""")
    page.locator("#lin-panel .lin-result").click()
    assert page.locator("#lin-panel h2").inner_text() == "Output"
    if width < 900:
        page.click("#lin-options-toggle")
    with page.expect_download() as download:
        page.click("#lin-export-view")
    assert download.value.suggested_filename == "linexcel-view.png"
    assert not errors


@pytest.mark.parametrize("width,height", [(320, 844), (768, 900), (844, 390)])
def test_investigation_controls_remain_bounded_in_french(report, width, height):
    graph = rich_graph()
    graph["meta"]["screenshots"] = {"Out": graph["meta"]["screenshots"][:1]}
    page, errors = report(graph, width=width, height=height, language="fr")
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    page.click("#lin-explore")
    assert page.locator("#lin-cy").bounding_box()["height"] >= 100
    assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
    capture(page, f"investigation-fr-{width}x{height}")
    page.fill("#lin-search", "Output")
    page.press("#lin-search", "Enter")
    bounds = page.locator("#lin-search-results").bounding_box()
    assert bounds["x"] >= 0 and bounds["x"] + bounds["width"] <= width
    page.click("#lin-theme")
    page.click("#lin-options-toggle")
    assert page.locator("#lin-options-dialog").is_visible()
    assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
    capture(page, f"options-fr-{width}x{height}")
    page.keyboard.press("Escape")
    assert not errors


def test_semantic_limitation_is_scoped_and_never_declares_value_incorrect(report):
    graph = filtered_graph()
    graph["nodes"][1].update(
        verification="unverified_engine_semantics",
        semanticRisks=[{"feature": "<img onerror=alert(1)>", "origin": "dependency"}],
        doc="## A generated explanation",
    )
    page, errors = report(graph)
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    notice = page.locator("#lin-semantic-notice")
    assert "Potentially affected" in notice.inner_text()
    assert "not independently verified" in notice.inner_text()
    assert notice.locator("img").count() == 0
    assert notice.bounding_box()["y"] < page.locator(".lin-ai-box").bounding_box()["y"]
    page.evaluate("() => { cy.getElementById('i').emit('tap'); }")
    assert page.locator("#lin-semantic-notice").count() == 0
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_interrupted_empty_report_notice_persists_across_tabs(
    browser, tmp_path, width, height
):
    graph = rich_graph()
    graph["nodes"] = []
    graph["edges"] = []
    graph["meta"]["stats"] = {"totalNodes": 0, "totalEdges": 0, "totalFormulas": None}
    graph["meta"]["execution"] = {
        "status": "timed_out",
        "phase": "<img onerror=alert(1)>",
        "elapsedSeconds": 1.3,
        "budgetSeconds": 1,
        "isolated": True,
    }
    path = tmp_path / "interrupted.html"
    path.write_text(render_html(graph, language="fr"), encoding="utf-8")
    page = browser.new_page(viewport={"width": width, "height": height})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    try:
        page.goto(path.as_uri())
        assert "— formules" in page.locator("#lin-stats").inner_text()
        assert page.locator(".lin-emptygraph-title").inner_text() == (
            "Aucun nœud de lignage disponible dans ce rapport."
        )
        for tab in ("graph", "sheets", "overview"):
            page.click(f"#lin-tab-{tab}")
            notice = page.locator("#lin-execution-notice")
            assert notice.is_visible()
            assert "Résultats incomplets" in notice.inner_text()
            assert notice.locator("img").count() == 0
            assert page.locator("#lin-options-host").is_hidden()
            assert page.evaluate(
                "document.documentElement.scrollWidth <= window.innerWidth"
            )
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_source_recovery_notice_is_visible_across_tabs(report, width, height):
    graph = rich_graph()
    graph["meta"]["execution"] = {
        "status": "memory_limit",
        "phase": "engine evaluation",
        "recovery": {"status": "completed", "mode": "source_only"},
    }
    page, errors = report(graph, width=width, height=height, language="fr")
    for tab in ("graph", "sheets", "overview"):
        page.click(f"#lin-tab-{tab}")
        notice = page.locator("#lin-execution-notice")
        assert notice.is_visible()
        assert "Inventaire partiel" in notice.inner_text()
        assert "sans recalcul ni graphe de dépendances" in notice.inner_text()
        box = notice.bounding_box()
        assert box["x"] >= 0 and box["x"] + box["width"] <= width
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_all_rich_tabs_and_search_navigation(report, width, height):
    page, errors = report(rich_graph(), width=width, height=height)
    page.click("#lin-tab-overview")
    assert (
        page.locator("#lin-overview .lin-doc h1").inner_text() == "Synthetic overview"
    )
    assert page.locator("#lin-overview td").count() == 2
    page.click("#lin-tab-sheets")
    assert page.locator("#lin-sheet-details td").count() == 2
    page.click("#lin-tab-screenshots")
    page.locator("#lin-screenshots .lin-chip").nth(1).click()
    assert (
        page.locator("#lin-screenshots .lin-chip").nth(1).get_attribute("aria-pressed")
        == "true"
    )
    page.wait_for_function("document.querySelector('.lin-shot').naturalWidth === 120")
    # Search is available in the shared header: submitting a match must make
    # the selected cell visible instead of modifying a hidden graph.
    page.fill("#lin-search", "Output")
    page.press("#lin-search", "Enter")
    assert page.locator("#lin-graph-main").is_visible()
    assert page.locator("#lin-panel .lin-doc h2").inner_text() == (
        "Synthetic cell documentation"
    )
    assert page.evaluate("cy.getElementById('c').selected()")
    assert not errors


def test_documentation_renders_nested_lists_headings_and_literal_code(report):
    graph = rich_graph()
    graph["meta"]["workbookDoc"] = (
        "# Workbook `example.xlsx`\n\n"
        "- Outer\n  - Inner **bold**\n  - Another child\n- Next\n\n"
        "1. First step\n2. Second step\n\n"
        '```excel\n=IF(A1<2,"<tag>","ok")\n```'
    )
    page, errors = report(graph)
    page.click("#lin-tab-overview")
    assert page.locator("#lin-overview h1 code").inner_text() == "example.xlsx"
    assert page.locator("#lin-overview ul > li > ul > li").count() == 2
    assert page.locator("#lin-overview ul ul strong").inner_text() == "bold"
    assert page.locator("#lin-overview ol > li").count() == 2
    assert page.locator("#lin-overview pre code").inner_text() == (
        '=IF(A1<2,"<tag>","ok")'
    )
    assert page.locator("#lin-overview tag").count() == 0
    assert not errors


def test_sheet_context_limits_are_visible_and_escaped(report):
    graph = rich_graph()
    warning = 'Comments not scanned: <img src=x onerror="window.injected=true">'
    context = graph["meta"]["workbookContext"]
    context["warnings"] = [warning]
    context["sheets"][0]["dimensions"] = {"rows": None, "columns": None}
    page, errors = report(graph)
    page.click("#lin-tab-sheets")
    note = page.locator(".lin-context-warnings")
    assert note.is_visible()
    assert note.inner_text() == warning
    assert note.locator("img").count() == 0
    assert "— rows × — columns" in page.locator("#lin-sheet-details").inner_text()
    assert page.evaluate("window.injected === undefined")
    assert not errors


def test_sheet_warnings_are_scoped_and_icon_is_compact(report):
    graph = rich_graph()
    context = graph["meta"]["workbookContext"]
    context["sheets"] = [
        {
            "name": "FirstSheet",
            "visibility": "visible",
            "dimensions": {"rows": 10, "columns": 5},
            "preview_range": "A1:E10",
            "preview": [{"row": 1, "values": [1, 2, 3, 4, 5]}],
        },
        {
            "name": "SecondSheet",
            "visibility": "visible",
            "dimensions": {"rows": 20, "columns": 8},
            "preview_range": "A1:H20",
            "preview": [{"row": 1, "values": [10, 20]}],
        },
    ]
    context["warnings"] = [
        "Large workbook: sheet context uses bounded previews",
        "Comments on 'FirstSheet' were truncated for inspection",
    ]
    page, errors = report(graph)
    page.click("#lin-tab-sheets")

    # First sheet is active by default
    first_note = page.locator(".lin-context-warnings")
    assert first_note.is_visible()
    first_text = first_note.inner_text()
    assert "Large workbook: sheet context uses bounded previews" in first_text
    assert "Comments on 'FirstSheet' were truncated for inspection" in first_text

    # The warning icon must be compact and aligned, not an unbounded giant SVG
    icon_box = page.locator(".lin-context-warnings svg").bounding_box()
    assert icon_box is not None
    assert 0 < icon_box["width"] <= 24
    assert 0 < icon_box["height"] <= 24

    # Switch to SecondSheet
    page.locator("#lin-sheets-sidebar button").nth(1).click()
    second_note = page.locator(".lin-context-warnings")
    assert second_note.is_visible()
    second_text = second_note.inner_text()
    # Global warning is present on SecondSheet, but FirstSheet-specific warning is NOT
    assert "Large workbook: sheet context uses bounded previews" in second_text
    assert "Comments on 'FirstSheet' were truncated for inspection" not in second_text

    assert not errors


def test_sheet_without_warnings_has_no_notice_and_scroll_resets(report):
    graph = rich_graph()
    context = graph["meta"]["workbookContext"]
    context["sheets"] = [
        {
            "name": "TruncatedSheet",
            "visibility": "visible",
            "dimensions": {"rows": 50, "columns": 5},
            "preview_range": "A1:E50",
            "preview": [{"row": i, "values": [i]} for i in range(1, 51)],
        },
        {
            "name": "CleanSheet",
            "visibility": "visible",
            "dimensions": {"rows": 10, "columns": 5},
            "preview_range": "A1:E10",
            "preview": [{"row": 1, "values": [1]}],
        },
    ]
    context["warnings"] = [
        "Comments on 'TruncatedSheet' were truncated for inspection",
    ]
    page, errors = report(graph)
    page.click("#lin-tab-sheets")

    assert page.locator(".lin-context-warnings").count() == 1

    # Scroll down on the first sheet
    page.evaluate(
        "() => { document.querySelector('.lin-sheet-body').scrollTop = 200; }"
    )
    assert (
        page.evaluate("() => document.querySelector('.lin-sheet-body').scrollTop") > 0
    )

    # Switch to CleanSheet: no notice rendered, and scroll resets to 0
    page.locator("#lin-sheets-sidebar button").nth(1).click()
    assert page.locator(".lin-context-warnings").count() == 0
    assert (
        page.evaluate("() => document.querySelector('.lin-sheet-body').scrollTop") == 0
    )

    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_flat_page_vision_follows_explicit_identity_and_native_zoom(
    report, width, height
):
    graph = rich_graph()
    graph["meta"]["screenshots"] = [
        s.replace("width='120'", "width='2200'") for s in graph["meta"]["screenshots"]
    ]
    graph["meta"]["screenshotNames"] = ["finance-01", "finance-02"]
    graph["meta"]["screenshotDocs"] = {"finance-02": "## Second page\n**Blue** inputs."}
    page, errors = report(graph, width=width, height=height)
    page.click("#lin-tab-screenshots")
    assert page.locator(".lin-shot-description .lin-ai-box").count() == 0
    page.locator(".lin-shot-tabs button").nth(1).click()
    assert page.locator(".lin-shot-description h2").inner_text() == "Second page"
    assert page.locator(".lin-shot-description strong").inner_text() == "Blue"
    page.wait_for_function("document.querySelector('.lin-shot').naturalWidth === 2200")
    page.locator("#lin-screenshots .lin-shot-viewer > button").click()
    assert page.locator("#lin-screenshots .lin-frame").evaluate(
        "e => e.scrollWidth > e.clientWidth"
    )
    assert page.locator("#lin-screenshots-main").evaluate(
        "e => e.scrollWidth === e.clientWidth"
    )
    page.locator(".lin-shot-tabs button").nth(0).click()
    assert page.locator(".lin-shot-description .lin-ai-box").count() == 0
    assert not errors


@pytest.mark.parametrize("with_context", [True, False])
def test_unmapped_chart_images_and_vision_remain_reachable(report, with_context):
    graph = rich_graph()
    first, second = graph["meta"]["screenshots"]
    graph["meta"]["screenshots"] = {"Out": [first], "Chart": [second]}
    graph["meta"]["screenshotDocs"] = {
        "Out": "Sheet description",
        "Chart": "Chart description",
    }
    if not with_context:
        graph["meta"].pop("workbookContext")
    page, errors = report(graph, width=390, height=844)
    if with_context:
        page.click("#lin-tab-sheets")
        assert (
            page.locator("#lin-sheet-details .lin-doc").inner_text()
            == "Sheet description"
        )
        assert page.locator("#lin-sheet-details img").count() == 1
    page.click("#lin-tab-screenshots")
    page.locator(".lin-shot-tabs button", has_text="Chart").click()
    assert (
        page.locator(".lin-shot-description .lin-doc").inner_text()
        == "Chart description"
    )
    assert page.locator(".lin-shot").get_attribute("src") == second
    if not with_context:
        page.locator(".lin-shot-tabs button", has_text="Out").click()
        assert (
            page.locator(".lin-shot-description .lin-doc").inner_text()
            == "Sheet description"
        )
    assert not errors


def test_legacy_embedded_pages_keep_unpaired_descriptions_named_and_escaped(report):
    graph = rich_graph()
    graph["meta"]["screenshotDocs"] = {"old-page": "**Visible** <img onerror='evil()'>"}
    page, errors = report(graph)
    page.click("#lin-tab-screenshots")
    assert page.locator("#lin-screenshots h3").inner_text() == "old-page"
    assert page.locator("#lin-screenshots .lin-doc strong").inner_text() == "Visible"
    assert page.locator("#lin-screenshots .lin-doc img").count() == 0
    assert page.locator(".lin-shot-description .lin-doc").count() == 0
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_markdown_tables_keep_words_whole_and_scroll_long_tokens(report, width, height):
    graph = rich_graph()
    graph["nodes"][1]["doc"] = (
        "| Sheet | Rows | Formula | Notable layout |\n|---|---|---|---|\n"
        "| Summary | 7 x 3 | 100 | Revenue model with a long description |\n\n"
        "| Reference | Meaning |\n|---|---|\n| " + "A" * 180 + " | Long token |"
    )
    page, errors = report(graph, width=width, height=height)
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    for selector, word in (("td", "Summary"), ("th", "Formula")):
        cell = page.locator("#lin-panel .lin-mdt " + selector).filter(has_text=word)
        assert (
            cell.evaluate("""e => {
            const range=document.createRange();range.selectNodeContents(e);
            return range.getClientRects().length;
        }""")
            == 1
        )
    assert (
        page.locator("#lin-panel .lin-mdt")
        .nth(1)
        .evaluate("e=>e.scrollWidth>e.clientWidth")
    )
    assert page.locator("#lin-panel").evaluate("e=>e.scrollWidth===e.clientWidth")
    assert not errors


def test_group_card_shows_count_once_without_changing_source_label(report):
    graph = filtered_graph()
    graph["nodes"] = [
        {"id": "g", "kind": "group", "label": "Sales!E4 x100", "count": 100},
    ]
    graph["edges"] = []
    page, errors = report(graph)
    assert page.evaluate("cy.getElementById('g').data('label')") == "Sales!E4\n×100"
    page.evaluate("() => { cy.getElementById('g').emit('tap'); }")
    assert page.locator("#lin-panel h2").inner_text() == "Sales!E4 x100"
    assert not errors


def test_search_submit_clear_and_filter_feedback(report):
    page, errors = report(filtered_graph())
    page.fill("#lin-search", "Output")
    page.click("#lin-search-submit")
    assert page.evaluate("cy.getElementById('c').selected()")
    assert page.locator("#lin-search-status").inner_text()
    page.fill("#lin-search", "missing")
    assert not page.locator("#lin-search-status").inner_text()
    page.press("#lin-search", "Enter")
    assert page.locator("#lin-search-status").inner_text() == "No matches"
    assert page.evaluate("cy.nodes(':selected').length") == 0
    assert "lin-empty" in page.locator("#lin-panel").get_attribute("class")
    for change in ("sheet", "kind", "diffs"):
        page.fill("#lin-search", "Output")
        page.click("#lin-search-submit")
        if change == "sheet":
            page.select_option("#lin-sheet-filter", "Out")
        elif change == "kind":
            page.locator("#lin-rail-kinds button", has_text="Source data").click()
        else:
            page.click("#lin-diffs-only")
        assert not page.locator("#lin-search-status").inner_text()
        assert page.locator("#lin-search").input_value() == ""
    page.fill("#lin-search", "Output")
    page.click("#lin-search-submit")
    page.click("#lin-search-clear")
    assert page.evaluate("cy.nodes(':selected').length") == 0
    assert not errors


def test_redundant_preview_is_hidden_and_three_tabs_keep_keyboard_navigation(report):
    graph = rich_graph()
    image = graph["meta"]["screenshots"][0]
    graph["meta"]["screenshots"] = {"Out": [image]}
    graph["meta"]["screenshotDocs"] = {"Out": "A visible description"}
    page, errors = report(graph)
    assert page.locator("#lin-tab-screenshots").is_hidden()
    page.focus("#lin-tab-graph")
    for key, expected in [
        ("End", "sheets"),
        ("ArrowRight", "graph"),
        ("ArrowLeft", "sheets"),
        ("Home", "graph"),
    ]:
        page.keyboard.press(key)
        assert (
            page.locator(f"#lin-tab-{expected}").get_attribute("aria-selected")
            == "true"
        )
    page.click("#lin-tab-sheets")
    assert page.locator("#lin-sheet-details img").get_attribute("src") == image
    assert "A visible description" in page.locator("#lin-sheet-details").inner_text()
    assert not errors


@pytest.mark.parametrize(
    "width,height", [(1440, 900), (768, 540), (844, 390), (390, 844), (320, 700)]
)
def test_command_bar_keeps_selection_above_detail_panel(report, width, height):
    graph = filtered_graph()
    name = "Prévisions financières — scénario consolidé international 2026"
    graph["nodes"][1]["sheet"] = name
    page, errors = report(graph, width=width, height=height, language="fr")
    page.get_by_role("combobox", name="Filtre par feuille").select_option(name)
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    page.focus("#lin-sheet-filter")
    page.keyboard.press("Escape")
    assert page.evaluate("cy.getElementById('c').selected()")
    page.click("#lin-fit-sel")
    page.wait_for_timeout(650)
    bounds = page.evaluate("""() => {
        const canvas = document.querySelector('#lin-cy').getBoundingClientRect();
        const panel = document.querySelector('#lin-panel').getBoundingClientRect();
        const commands = document.querySelector('#lin-tools').getBoundingClientRect();
        const selected = cy.getElementById('c').renderedBoundingBox();
        return {height: canvas.height, overlap: canvas.bottom - panel.top,
                commandsBottom: commands.bottom, canvasTop: canvas.top,
                x1: selected.x1, x2: selected.x2, y1: selected.y1, y2: selected.y2,
                width: canvas.width,
                overflow: document.documentElement.scrollWidth > innerWidth};
    }""")
    assert not bounds["overflow"]
    assert bounds["height"] >= 80
    assert bounds["commandsBottom"] <= bounds["canvasTop"] + 1
    if width <= 900:
        assert bounds["overlap"] <= 1
    assert bounds["x1"] >= 0 and bounds["x2"] <= bounds["width"]
    assert bounds["y1"] >= 0 and bounds["y2"] <= bounds["height"]
    assert page.locator("#lin-sheet-filter").input_value() == name
    assert not errors


def test_zoom_preserves_center_reports_percentage_and_clamps_bounds(report):
    page, errors = report(filtered_graph())
    page.evaluate("() => { cy.zoom(1); cy.pan({x:30,y:60}); }")
    center = (
        "({x:(cy.width()/2-cy.pan().x)/cy.zoom(),"
        "y:(cy.height()/2-cy.pan().y)/cy.zoom()})"
    )
    before = page.evaluate(center)
    page.click("#lin-zoom-in")
    page.wait_for_function(
        "document.querySelector('#lin-zoom-level').textContent === '140%'"
    )
    assert page.evaluate(center) == pytest.approx(before)
    page.click("#lin-zoom-out")
    assert page.evaluate("cy.zoom()") == pytest.approx(1)
    page.evaluate("() => { cy.zoom(cy.maxZoom()); }")
    page.wait_for_function("document.querySelector('#lin-zoom-in').disabled")
    page.evaluate("() => { cy.zoom(cy.minZoom()); }")
    page.wait_for_function("document.querySelector('#lin-zoom-out').disabled")
    # All elements can be hidden by composing valid filters; Fit remains safe.
    page.click("#lin-diffs-only")
    page.select_option("#lin-sheet-filter", "Other")
    assert visible_ids(page) == []
    zoom = page.evaluate("cy.zoom()")
    page.click("#lin-fit")
    assert page.evaluate("cy.zoom()") == zoom
    assert page.locator("#lin-fit-sel").is_disabled()
    assert not errors


@pytest.mark.parametrize("missing_library", [False, True])
def test_search_button_stands_down_with_unavailable_graph(
    browser, tmp_path, missing_library
):
    graph = filtered_graph() if missing_library else {"nodes": [], "edges": []}
    graph["meta"] = {
        "stats": {
            "totalNodes": len(graph["nodes"]),
            "totalEdges": len(graph["edges"]),
            "totalFormulas": 0,
        }
    }
    html = render_html(graph)
    if missing_library:
        html = html.replace(
            "function boot() {", "function boot() { window.cytoscape = undefined;"
        )
    path = tmp_path / "unavailable.html"
    path.write_text(html, encoding="utf-8")
    page = browser.new_page()
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    try:
        page.goto(path.as_uri())
        assert page.locator("#lin-search").is_disabled()
        assert page.locator("#lin-search-submit").is_disabled()
        assert page.locator("#lin-options-toggle").is_hidden()
        assert not errors
    finally:
        page.close()


def test_inline_code_preserves_operators_and_literal_markup_inside_emphasis(report):
    graph = rich_graph()
    graph["meta"]["workbookDoc"] = (
        "Formula `C4 * D4` (R1C1: `RC[-2] * RC[-1]`).\n\n"
        '`**literal**` and `<img src=x onerror="window.injected=true">`.\n\n'
        "**before `A1*B1` after** and *before `x` after*.\n\n"
        "Literal marker \x000\x00 and `2*3` with `4*5`."
    )
    page, errors = report(graph)
    page.click("#lin-tab-overview")
    code = page.locator("#lin-overview code")
    assert code.all_text_contents() == [
        "C4 * D4",
        "RC[-2] * RC[-1]",
        "**literal**",
        '<img src=x onerror="window.injected=true">',
        "A1*B1",
        "x",
        "2*3",
        "4*5",
    ]
    assert page.locator("#lin-overview code *").count() == 0
    assert page.locator("#lin-overview strong").inner_text() == "before A1*B1 after"
    assert page.locator("#lin-overview em").inner_text() == "before x after"
    assert page.locator("#lin-overview img").count() == 0
    assert page.evaluate("window.injected === undefined")
    assert not errors


@pytest.mark.parametrize("width,height", [(1440, 900), (390, 844)])
def test_qualification_quotes_are_readable_scoped_and_escaped(report, width, height):
    graph = rich_graph()
    qualification = (
        "> **Qualification.** Check this formula quotation against the source.\n"
        ">\n"
        '> `=IF(A1>0,A1,0)` and <img src=x onerror="window.injected=true">\n\n'
        "Ordinary prose with `A1 > 0` follows outside the qualification.\n\n"
        "```text\n> literal code marker\n```"
    )
    graph["nodes"][1]["doc"] = qualification
    graph["meta"]["workbookDoc"] = qualification
    graph["meta"]["screenshots"] = {"Out": graph["meta"]["screenshots"]}
    graph["meta"]["screenshotDocs"] = {"Out": qualification}
    page, errors = report(graph, width=width, height=height)

    for tab, container in (
        ("overview", "#lin-overview"),
        ("graph", "#lin-panel"),
        ("sheets", "#lin-sheet-details"),
    ):
        page.click(f"#lin-tab-{tab}")
        if tab == "graph":
            page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
        quote = page.locator(f"{container} blockquote")
        assert quote.count() == 1
        assert quote.locator("strong").inner_text() == "Qualification."
        assert quote.locator("p").count() == 2
        assert quote.locator("code").inner_text() == "=IF(A1>0,A1,0)"
        assert not any(line.startswith(">") for line in quote.inner_text().splitlines())
        assert "Ordinary prose" not in quote.inner_text()
        literal = page.locator(f"{container} pre code").inner_text()
        assert literal == "> literal code marker"
        assert quote.locator("img").count() == 0
        assert page.evaluate("window.injected === undefined")
        quote.scroll_into_view_if_needed()
        geometry = quote.evaluate("""element => {
            const box = element.getBoundingClientRect();
            return {left: box.left, right: box.right,
                overflow: element.scrollWidth > element.clientWidth + 1,
                border: parseFloat(getComputedStyle(element).borderLeftWidth)};
        }""")
        assert geometry["left"] >= 0 and geometry["right"] <= width
        assert not geometry["overflow"] and geometry["border"] > 0
        capture(page, f"qualification-{tab}-{width}")
    assert not errors


@pytest.mark.parametrize(
    "language,singular,plural",
    [
        ("en", "1 step ·", "2 steps ·"),
        ("fr", "1 étape ·", "2 étapes ·"),
        ("es", "1 paso ·", "2 pasos ·"),
        ("de", "1 Schritt ·", "2 Schritten ·"),
        ("it", "1 passo ·", "2 passi ·"),
        ("pt", "1 passo ·", "2 passos ·"),
        ("nl", "1 stap ·", "2 stappen ·"),
        ("ja", "1ステップ", "2ステップ"),
        ("zh", "1步", "2步"),
    ],
)
def test_local_exploration_uses_localized_step_count(
    report, language, singular, plural
):
    page, errors = report(filtered_graph(), language=language)
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    page.click("#lin-explore")
    assert singular in page.locator("#lin-explore-status").inner_text()
    page.click("#lin-explore-depth")
    assert plural in page.locator("#lin-explore-status").inner_text()
    page.click("#lin-explore-depth")
    assert singular in page.locator("#lin-explore-status").inner_text()
    assert not errors


@pytest.mark.parametrize("width,height", [(320, 700), (390, 844), (844, 390)])
def test_graph_options_dialog_keyboard_and_camera_state(report, width, height):
    graph = rich_graph()
    graph["meta"]["screenshots"] = {"Out": [graph["meta"]["screenshots"][0]]}
    page, errors = report(graph, width=width, height=height, language="fr")
    page.select_option("#lin-sheet-filter", "Out")
    page.evaluate("() => { cy.getElementById('c').emit('tap'); }")
    page.click("#lin-fit-sel")
    page.wait_for_timeout(650)
    state = page.evaluate(
        "({zoom:cy.zoom(),pan:cy.pan(),selected:cy.nodes(':selected').map(n=>n.id())})"
    )
    assert page.locator(".lin-legend").count() == 0
    assert (
        page.locator("#lin-tab-overview").get_attribute("aria-label")
        == "Synthèse générale"
    )
    tabs = page.locator("[role=tab]:visible")
    assert tabs.count() == 3
    for tab in tabs.all():
        rect = tab.bounding_box()
        assert rect["x"] >= 0 and rect["x"] + rect["width"] <= width
    page.get_by_role("button", name="Options du graphe", exact=True).click()
    dialog = page.get_by_role("dialog", name="Options du graphe")
    assert dialog.is_visible()
    assert page.evaluate("document.activeElement.parentElement.id") == "lin-rail-kinds"
    assert dialog.locator("#lin-rail-kinds .lin-sw").count() == 2
    for _ in range(12):
        page.keyboard.press("Tab")
        assert page.evaluate(
            "document.activeElement.closest('#lin-options-dialog') !== null"
        )
    page.keyboard.press("/")
    assert page.evaluate(
        "document.activeElement.closest('#lin-options-dialog') !== null"
    )
    page.keyboard.press("Escape")
    page.wait_for_function("!document.querySelector('#lin-options-dialog').open")
    assert page.evaluate("document.activeElement.id") == "lin-options-toggle"
    assert page.locator("#lin-sheet-filter").input_value() == "Out"
    assert (
        page.evaluate(
            "({zoom:cy.zoom(),pan:cy.pan(),selected:cy.nodes(':selected').map(n=>n.id())})"
        )
        == state
    )
    assert page.evaluate("cy.height()") >= (101 if width == 844 else 80)
    page.click("#lin-tab-overview")
    assert page.locator("#lin-options-toggle").is_hidden()
    page.click("#lin-tab-graph")
    assert page.locator("#lin-options-toggle").is_visible()
    assert not errors


def test_graph_options_filters_survive_reopen_and_desktop_resize(report):
    page, errors = report(filtered_graph(), width=390, height=844)
    page.click("#lin-options-toggle")
    page.locator("#lin-rail-kinds button", has_text="Source data").click()
    page.click("#lin-options-close")
    assert visible_ids(page) == ["c"]
    page.click("#lin-options-toggle")
    assert (
        page.locator("#lin-rail-kinds button", has_text="Source data").get_attribute(
            "aria-pressed"
        )
        == "false"
    )
    page.set_viewport_size({"width": 1440, "height": 900})
    page.wait_for_function("!document.querySelector('#lin-options-dialog').open")
    assert page.locator("#lin-options-toggle").is_hidden()
    page.locator("#lin-options-host #lin-rail-kinds").wait_for(state="visible")
    assert page.locator("#lin-options-host #lin-rail-kinds").is_visible()
    assert visible_ids(page) == ["c"]
    page.locator("#lin-rail-kinds button", has_text="Source data").click()
    page.wait_for_function("cy.getElementById('i').visible()")
    assert visible_ids(page) == ["c", "i", "u"]
    assert not errors


def test_graph_options_immediate_reopen_keeps_content_and_focus(report):
    page, errors = report(filtered_graph(), width=390, height=844)
    page.click("#lin-options-toggle")
    page.evaluate("""() => {
        const dialog = document.querySelector('#lin-options-dialog');
        window.previousCloseDelivered = false;
        dialog.addEventListener('close', () => {
            window.previousCloseDelivered = true;
        }, {once: true});
        document.querySelector('#lin-options-close').click();
        document.querySelector('#lin-options-toggle').click();
    }""")
    page.wait_for_function("window.previousCloseDelivered")
    assert page.locator("#lin-options-dialog").is_visible()
    assert page.locator("#lin-options-dialog #lin-rail-kinds").is_visible()
    assert page.evaluate(
        "document.activeElement.closest('#lin-options-dialog') !== null"
    )
    page.locator("#lin-rail-kinds button", has_text="Source data").click()
    assert visible_ids(page) == ["c"]
    page.keyboard.press("Escape")
    page.wait_for_function(
        "document.querySelector('#lin-options-body').parentElement.id"
        " === 'lin-options-host'"
    )
    assert page.evaluate("document.activeElement.id") == "lin-options-toggle"
    assert not errors
