"""Physical access to secondary graph controls keeps the primary UI closed."""


def open_graph_options(page):
    options = page.locator("#graph-options")
    if not options.evaluate("element => element.open"):
        options.locator(":scope > summary").click()
    assert options.evaluate("element => element.open")


def tap_graph_node(page, node_id):
    canvas = page.locator("#graph-canvas")
    canvas.scroll_into_view_if_needed()
    page.wait_for_function("!document.querySelector('#graph-canvas')._cy.animated()")
    page.evaluate(
        "() => new Promise(resolve => requestAnimationFrame(() => "
        "requestAnimationFrame(resolve)))"
    )
    point = canvas.evaluate(
        """(element, id) => {
            const p=element._cy.getElementById(id).renderedPosition();
            const b=element.getBoundingClientRect();
            return {x:b.left+element.clientLeft+p.x,
                y:b.top+element.clientTop+p.y};
        }""",
        node_id,
    )
    assert page.evaluate(
        "p => document.elementFromPoint(p.x,p.y)?.closest('#graph-canvas') "
        "=== document.querySelector('#graph-canvas')",
        point,
    )
    page.mouse.click(point["x"], point["y"])
