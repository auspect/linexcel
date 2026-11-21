"""Physical access to secondary graph controls keeps the primary UI closed."""

import re

# French that an English page must not show: function words, accented letters
# and the French units. A French sentence left behind has at least one.
FRENCH = re.compile(
    r"\b(?:le|la|les|des|du|une|un|et|est|sur|pour|dans|avec|aux|au|ou|par|qui"
    r"|que|sa|son|ses|ce|cette|pas|ne|vous|nous|sont|Mio)\b"
    r"|[àâçéèêëîïôûùœ]|\d\s?Go\b",
    re.IGNORECASE,
)


def french_words(text):
    return sorted({match.group(0) for match in FRENCH.finditer(text)})


def open_workspace_menu(page):
    menu = page.locator("#workspace-menu")
    if menu.is_visible() and menu.get_attribute("aria-expanded") != "true":
        menu.click()


def close_workspace_menu(page):
    close = page.locator("#workspace-close")
    if close.is_visible():
        close.click()


def open_settings(page, section="limits"):
    """Open the settings dialog from the sidebar, on one of its sections."""
    open_workspace_menu(page)
    page.locator("#settings-open").click()
    page.locator(f"#settings [data-settings='{section}']").click()
    page.locator(f"#settings-{section}").wait_for()


def close_settings(page):
    page.locator("#settings-done").click()
    page.locator("#settings").wait_for(state="hidden")


def open_graph_depth(page):
    if page.locator("#graph-depth").is_hidden():
        open_graph_options(page)


def open_data_options(page):
    options = page.locator("#data-options")
    if options.is_visible() and not options.evaluate("element => element.open"):
        options.locator(":scope > summary").click()


def open_graph_options(page):
    options = page.locator("#graph-options")
    if not options.evaluate("element => element.open"):
        options.locator(":scope > summary").click()
    assert options.evaluate("element => element.open")


def close_graph_options(page):
    """Close the options; on phones they are a bottom sheet over the page."""
    options = page.locator("#graph-options")
    if options.evaluate("element => element.open"):
        options.locator(":scope > summary").click()
    assert not options.evaluate("element => element.open")


def tap_graph_node(page, node_id):
    """Click the rendered HTML card, falling back to its accessible relation."""
    canvas = page.locator("#graph-canvas")
    canvas.scroll_into_view_if_needed()
    # Titles can be user-facing sheet labels; data-id is the stable identity.
    card = page.locator(".react-flow__node").locator(
        f'xpath=./self::*[@data-id="{node_id}"]'
    )
    page.evaluate(
        "() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))"
    )
    box = card.bounding_box()
    if box:
        point = {"x": box["x"] + box["width"] / 2, "y": box["y"] + box["height"] / 2}
        hit = page.evaluate(
            "([p,id]) => document.elementFromPoint(p.x,p.y)"
            "?.closest('.react-flow__node')?.dataset.id === id",
            [point, node_id],
        )
        if hit:
            page.mouse.click(point["x"], point["y"])
            return
    relation = page.locator(f'#graph [data-graph-node="{node_id}"]').first
    if relation.count():
        details = page.locator("#graph-list")
        if not details.evaluate("e => e.open"):
            details.locator(":scope > summary").click()
        relation.click()
        return
    card.click()
    return


def open_neighborhood(page):
    """Choose the explicit cell mode after the automatic workbook overview."""
    page.locator("#tab-graph").click()
    page.wait_for_function("""() => {
      const back=document.querySelector('#graph-neighborhood');
      const status=document.querySelector('#graph-build-status');
      return !document.querySelector('#graph-build').disabled ||
        (back&&!back.hidden) || status.textContent.includes('unavailable');
    }""")
    button = page.locator("#graph-neighborhood")
    if button.is_visible():
        button.click()
    if page.locator("#graph-sheet").is_hidden():
        open_graph_options(page)
