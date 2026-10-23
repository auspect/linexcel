"""Engine input reads and name resolution must not claim formula recalculation."""

from pathlib import Path

import pytest
from test_viewer_browser import browser as browser

from linexcel.i18n import LANGUAGES, UI_STRINGS
from linexcel.viewer import render_html


def value_graph():
    nodes = [
        {
            "id": "input",
            "kind": "input",
            "value": 5000,
            "cachedValue": 5000,
        },
        {
            "id": "name",
            "kind": "name",
            "targets": ["Params!C4"],
            "value": 5000,
            "cachedValue": 5000,
        },
        {"id": "input-no-cache", "kind": "input", "value": 5000},
        {
            "id": "name-no-cache",
            "kind": "name",
            "targets": ["Params!C4"],
            "value": 5000,
        },
        {"id": "input-unavailable", "kind": "input", "value": None},
        {
            "id": "input-different",
            "kind": "input",
            "value": 5000,
            "cachedValue": 4000,
        },
        {
            "id": "input-samples",
            "kind": "input",
            "value": 5000,
            "samples": [
                {
                    "addr": "C4",
                    "value": 5000,
                    "cachedValue": 5000,
                    "source": "engine",
                },
                {
                    "addr": "C5",
                    "value": 4000,
                    "source": "engine",
                },
            ],
        },
        {
            "id": "formula",
            "kind": "cell",
            "formula": "=2+3",
            "value": 5,
            "cachedValue": 5,
        },
        {
            "id": "formula-name",
            "kind": "name",
            "targets": ["Params!C4"],
            "formula": "=2+3",
            "value": 5,
            "cachedValue": 5,
        },
        {
            "id": "formula-samples",
            "kind": "group",
            "formula": "=B1*2",
            "count": 2,
            "samples": [
                {
                    "addr": "A1",
                    "value": 4,
                    "cachedValue": 4,
                    "valueSource": "engine",
                }
            ],
            "value": 4,
        },
    ]
    for node in nodes:
        node.update(label=node["id"], sheet="Params", valueSource="engine")
    return {
        "nodes": nodes,
        "edges": [],
        "meta": {"stats": {"totalFormulas": 3, "totalNodes": len(nodes)}},
    }


def open_values(browser, tmp_path, language, *, width=1440, fallback=False):
    page = browser.new_page(
        viewport={"width": width, "height": 844 if width == 390 else 900}
    )
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    html = render_html(value_graph(), language=language)
    if fallback:
        html = html.replace(
            "function boot() {", "function boot() { window.cytoscape = undefined;"
        )
    path = Path(tmp_path) / f"values-{language}-{width}-{fallback}.html"
    path.write_text(html, encoding="utf-8")
    page.goto(path.as_uri())
    if fallback:
        page.locator(".lin-fallback-list button").first.wait_for()
    else:
        page.wait_for_function("document.querySelector('#lin-cy')._cyreg?.cy != null")
    return page, errors


def select(page, node_id, *, fallback):
    if fallback:
        page.locator(f'.lin-fallback-list button[data-node-id="{node_id}"]').click()
    else:
        page.evaluate(
            "id => document.querySelector('#lin-cy')._cyreg.cy"
            ".getElementById(id).emit('tap')",
            node_id,
        )
    page.locator("#lin-panel .lin-vcard").first.wait_for()
    return page.locator("#lin-panel")


@pytest.mark.parametrize("language", LANGUAGES)
def test_input_and_name_readings_keep_their_provenance_in_every_language(
    browser, tmp_path, language
):
    page, errors = open_values(browser, tmp_path, language, fallback=True)
    text = UI_STRINGS[language]
    try:
        for node_id in ["input", "name"]:
            panel = select(page, node_id, fallback=True)
            assert (
                panel.locator(".lin-verdict").inner_text() == text["value_read_match"]
            )
            assert text["value_col_read"] in panel.text_content()
            assert text["value_engine_read_desc"] in panel.text_content()
            assert text["value_col_calc"] not in panel.text_content()
            assert text["value_match"] not in panel.text_content()
        for node_id in ["input-no-cache", "name-no-cache"]:
            panel = select(page, node_id, fallback=True)
            assert not panel.locator(".lin-verdict").count()
            assert text["value_col_read"] in panel.text_content()
            assert text["value_read_no_cache_desc"] in panel.text_content()
            assert text["value_no_cache_desc"] not in panel.text_content()
        panel = select(page, "input-unavailable", fallback=True)
        assert text["coverage_unavailable"] in panel.text_content()
        assert text["value_not_recalc"] not in panel.text_content()
        panel = select(page, "input-different", fallback=True)
        assert panel.locator(".lin-verdict").inner_text() == text["value_read_mismatch"]
        assert panel.locator(".lin-vcard.is-diff").count() == 1
        panel = select(page, "input-samples", fallback=True)
        assert (
            text["value_col_read"] in panel.locator(".lin-vtable thead").text_content()
        )
        assert text["value_col_calc"] not in panel.text_content()
        for node_id in ["formula", "formula-name", "formula-samples"]:
            panel = select(page, node_id, fallback=True)
            assert text["value_match"] in panel.text_content()
            assert text["value_col_calc"] in panel.text_content()
            assert text["value_col_read"] not in panel.text_content()
        assert not errors
    finally:
        page.close()


@pytest.mark.parametrize("language", ["en", "fr"])
@pytest.mark.parametrize("width", [1440, 390])
def test_named_reading_and_formula_keep_distinct_labels_in_the_live_graph(
    browser, tmp_path, language, width
):
    page, errors = open_values(browser, tmp_path, language, width=width)
    text = UI_STRINGS[language]
    try:
        panel = select(page, "name", fallback=False)
        assert text["value_read_match"] in panel.text_content()
        assert text["value_col_read"] in panel.text_content()
        assert text["value_match"] not in panel.text_content()
        panel = select(page, "formula", fallback=False)
        assert text["value_match"] in panel.text_content()
        assert text["value_col_calc"] in panel.text_content()
        assert text["value_col_read"] not in panel.text_content()
        assert not page.evaluate("document.documentElement.scrollWidth > innerWidth")
        assert not errors
    finally:
        page.close()
