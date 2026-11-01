"""Native evidence stays compact and literal in both shipped browser surfaces."""

import base64
import json
import os
from pathlib import Path

import pytest
import test_app_viewer_regressions as navigation_tests
from graph_test_helpers import open_workspace_menu
from vision_helpers import image_bytes

from linexcel.viewer import render_html
from linexcel.vision_evidence import prepare_image, render_observations

LITERAL = "O'Brien & B (x) *literal* `code` <script> | &#x27; &amp; 90+"
app = navigation_tests.app
browser = navigation_tests.browser


def native_markdown():
    native = prepare_image(image_bytes(6400, 30))
    observations = [
        {
            "kind": "text",
            "text": LITERAL if index == 0 else f"Interior quote {index}",
            "certainty": "visible",
            "region_id": region.id,
            "original_bbox": list(region.bbox),
            "crop_sha256": region.sha256,
        }
        for index, region in enumerate(native.regions)
    ]
    return render_observations(
        [{"evidence": native.metadata(), "observations": observations}]
    )


def assert_disclosure(page, root):
    disclosure = root.locator("details")
    assert disclosure.count() == 1
    assert not disclosure.evaluate("element => element.open")
    hidden = disclosure.locator("li")
    assert hidden.count() == 1 and not hidden.is_visible()
    first = root.locator("li").first
    assert first.inner_text().startswith("“" + LITERAL + "”")
    assert first.locator("script,img,strong,em").count() == 0
    # Literal backticks remain literal; only generated receipt metadata is code.
    assert "code" not in first.locator("code").all_text_contents()
    assert root.locator("li details,li summary").count() == 0
    disclosure.locator("summary").click()
    assert disclosure.evaluate("element => element.open") and hidden.is_visible()
    assert root.evaluate(
        "element => {const r=element.getBoundingClientRect();"
        "return r.left>=0 && r.right<=innerWidth+1}"
    )
    assert page.evaluate("window.nativeEvidenceInjected || false") is False


def capture(page, name):
    directory = os.environ.get("LINEXCEL_VISION_CAPTURE_DIR")
    if directory:
        folder = Path(directory)
        folder.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(folder / name), full_page=True)


def install_serve_document(page, markdown):
    node = {
        "id": "First!A1",
        "sheet": "First",
        "cell": "A1",
        "kind": "input",
        "value": 1,
    }
    payload = {
        "project": {"id": "p", "name": "Native proof", "revision": 1},
        "graph": {
            "lazy": False,
            "nodes": [node],
            "edges": [],
            "meta": {"sheets": ["First"]},
        },
        "tasks": [
            {
                "id": "document",
                "operation": "document",
                "nodeId": node["id"],
                "status": "succeeded",
                "options": {"language": "fr"},
                "result": {"markdown": markdown, "language": "fr"},
            }
        ],
    }
    page.route(
        "**/api/projects/p",
        lambda route: route.fulfill(
            body=json.dumps(payload), content_type="application/json"
        ),
    )
    page.reload()
    page.locator("#node-list [data-node='First!A1']").wait_for()
    open_workspace_menu(page)
    page.locator("[data-project='p']").click()
    page.locator("#node-list [data-node='First!A1']").click()
    page.locator("#node-ai .ai-result").wait_for()


def test_serve_disclosure_is_collapsed_and_quotes_are_literal(app):
    page, _, errors, _, _ = app
    install_serve_document(page, native_markdown())
    root = page.locator("#node-ai .ai-result")
    width = page.viewport_size["width"]
    capture(page, f"serve-native-collapsed-{width}.png")
    assert_disclosure(page, root)
    capture(page, f"serve-native-expanded-{width}.png")
    assert not errors


@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_export_disclosure_is_collapsed_and_quotes_are_literal(browser, size):
    width, height = size
    page = browser.new_page(viewport={"width": width, "height": height})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    uri = "data:image/png;base64," + base64.b64encode(image_bytes()).decode()
    graph = {
        "nodes": [],
        "edges": [],
        "meta": {
            "filename": "Native proof",
            "stats": {"totalFormulas": 0, "totalNodes": 0, "totalEdges": 0},
            "screenshots": {"Proof": [uri]},
            "screenshotDocs": {"Proof": native_markdown()},
        },
    }
    try:
        page.set_content(render_html(graph), wait_until="load")
        assert not errors
        page.locator("#lin-tab-screenshots").click()
        root = page.locator(".lin-shot-description .lin-doc")
        capture(page, f"export-native-collapsed-{width}.png")
        assert_disclosure(page, root)
        capture(page, f"export-native-expanded-{width}.png")
        assert not errors
    finally:
        page.close()


def test_serve_recognizes_only_balanced_attribute_free_disclosures(app):
    page, _, errors, _, _ = app
    snippets = [
        "<details open><summary>Unsafe</summary>\n"
        "<script>window.nativeEvidenceInjected=true</script>\n</details>",
        "<details><summary>Unclosed</summary>\n- Body",
        "```html\n<details><summary>Code</summary>\n</details>\n```",
        "<details><summary>"
        "<img src=x onerror=window.nativeEvidenceInjected=true></summary>\n</details>",
    ]
    for text in snippets:
        install_serve_document(page, text)
        assert (
            page.locator(
                "#node-detail details,#node-detail script,#node-detail img"
            ).count()
            == 0
        )
        assert page.evaluate("window.nativeEvidenceInjected || false") is False
    assert not errors


@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_export_never_activates_unbalanced_attributed_or_fenced_html(browser, size):
    width, height = size
    page = browser.new_page(viewport={"width": width, "height": height})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    snippets = [
        "<details open><summary>Unsafe</summary>\n"
        "<img src=x onerror=window.nativeEvidenceInjected=true>\n</details>",
        "<details><summary>Unclosed</summary>\n- Body",
        "```html\n<details><summary>Code</summary>\n</details>\n```",
        "<details><summary><script>window.nativeEvidenceInjected=true</script>"
        "</summary>\n</details>",
    ]
    uri = "data:image/png;base64," + base64.b64encode(image_bytes()).decode()
    try:
        for markdown in snippets:
            graph = {
                "nodes": [],
                "edges": [],
                "meta": {
                    "filename": "Malformed proof",
                    "stats": {"totalFormulas": 0, "totalNodes": 0, "totalEdges": 0},
                    "screenshots": {"Proof": [uri]},
                    "screenshotDocs": {"Proof": markdown},
                },
            }
            page.set_content(render_html(graph), wait_until="load")
            page.locator("#lin-tab-screenshots").click()
            root = page.locator(".lin-shot-description .lin-doc")
            assert root.locator("details,img,script").count() == 0
            assert page.evaluate("window.nativeEvidenceInjected || false") is False
        assert not errors
    finally:
        page.close()
