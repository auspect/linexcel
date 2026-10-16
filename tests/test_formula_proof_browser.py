"""Model code wrappers must not hide deterministic source proof blocks."""

import pytest
import test_app_viewer_regressions as navigation_tests
from test_vision_disclosure import capture, install_serve_document

from linexcel.documentation_proof import render_formula_tokens
from linexcel.viewer import render_html

app = navigation_tests.app
browser = navigation_tests.browser
XML = "(" * 40 + "1+2" + ")" * 40
SOURCE = "=" + XML
CANONICAL = "=1 + 2"


def proof():
    raw = (
        "## Proof\n\n- Original: `{{SOURCE_FORMULA}}`\n"
        "- XML: ``{{STORED_XML_FORMULA}}``\n"
        "\n```text\n{{CANONICAL_FORMULA}}\n```\n\nEnd of proof."
    )
    markdown, _ = render_formula_tokens(
        raw,
        {
            "source_formula": SOURCE,
            "formula_provenance": {"xmlText": XML},
            "canonical_formula": CANONICAL,
        },
    )
    return markdown


def assert_proof(page, root):
    assert root.locator("pre").all_text_contents() == [SOURCE, XML, CANONICAL]
    assert "End of proof." in root.inner_text()
    assert "{{" not in root.inner_text() and "```" not in root.inner_text()
    assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")


def test_serve_model_code_wrappers_keep_three_readable_source_blocks(app):
    page, _, errors, _, _ = app
    install_serve_document(page, proof())
    assert_proof(page, page.locator("#node-ai .ai-result"))
    capture(page, f"serve-formula-wrappers-{page.viewport_size['width']}.png")
    assert not errors


@pytest.mark.parametrize("size", [(1440, 900), (390, 844)])
def test_export_model_code_wrappers_keep_three_readable_source_blocks(browser, size):
    page = browser.new_page(viewport={"width": size[0], "height": size[1]})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    graph = {
        "nodes": [],
        "edges": [],
        "meta": {
            "workbookDoc": proof(),
            "stats": {"totalFormulas": 0, "totalNodes": 0, "totalEdges": 0},
        },
    }
    try:
        page.set_content(render_html(graph), wait_until="load")
        page.locator("#lin-tab-overview").click()
        assert_proof(page, page.locator("#lin-overview .lin-doc"))
        capture(page, f"export-formula-wrappers-{size[0]}.png")
        assert not errors
    finally:
        page.close()
