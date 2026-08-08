"""Tests for the report's chrome: layout, theming and accessibility.

The viewer is a single template string that ships CSS, DOM and JS together, so
these assert on the generated document rather than on a running browser. Each
test names the defect it guards against, because the value of a presentation
test is only ever "this regression cannot come back quietly".
"""

import re

from linexcel.i18n import LANGUAGES, UI_STRINGS
from linexcel.viewer import render_html

EN = UI_STRINGS["en"]


def demo_graph() -> dict:
    """Two sheets, one formula, one input — enough to exercise the chrome."""
    return {
        "nodes": [
            {
                "id": "i:In!A1",
                "label": "In!A1",
                "kind": "input",
                "sheet": "In",
                "value": 3,
            },
            {
                "id": "c:Out!B1",
                "label": "Out!B1",
                "kind": "cell",
                "sheet": "Out",
                "formula": "=In!A1*2",
                "value": 6,
            },
        ],
        "edges": [
            {"id": "e1", "source": "i:In!A1", "target": "c:Out!B1", "kind": "ref"}
        ],
    }


def script_of(html: str) -> str:
    """The viewer's own script block, without the embedded Cytoscape bundles."""
    return html.split("var GRAPH = ", 1)[1]


class TestInlineStylesLiftedToCss:
    """Layout used to be dozens of ``element.style.x = ...`` assignments.

    They are now CSS classes; the only assignments left are the ones that
    genuinely depend on data — a palette colour and the step indent.
    """

    def test_only_data_driven_style_assignments_survive(self):
        props = set(re.findall(r"\.style\.(\w+) =", render_html(demo_graph())))
        assert props == {"background", "color", "marginLeft"}, props

    def test_the_screenshot_pane_no_longer_injects_a_stylesheet(self):
        html = render_html(demo_graph())
        assert "document.head.appendChild(styleNode)" not in html
        assert '.lin-chip[aria-pressed="true"]' in html

    def test_the_layout_classes_are_declared(self):
        html = render_html(demo_graph())
        for cls in (
            ".lin-panel-head",
            ".lin-card-head",
            ".lin-badges",
            ".lin-list",
            ".lin-note",
            ".lin-gallery",
            ".lin-shots",
            ".lin-shot-tabs",
            ".lin-empty-icon",
            ".lin-sheets-wrap",
            ".lin-sheet-body",
        ):
            assert cls + " {" in html or cls + "," in html, cls

    def test_the_markup_carries_no_style_attribute(self):
        """The template's own DOM used to inline flex rules and display:none."""
        head = render_html(demo_graph()).split("<script>\n(function", 1)[0]
        assert 'style="' not in head.split("</style>", 1)[1]


class TestTopBarGroups:
    def test_identity_navigation_and_tools_are_three_containers(self):
        html = render_html(demo_graph())
        assert '<div class="lin-id">' in html
        assert '<div class="lin-tabs" role="tablist" id="lin-tablist">' in html
        assert '<div class="lin-tools" id="lin-tools">' in html

    def test_the_graph_tools_live_in_the_tools_group(self):
        bar = render_html(demo_graph()).split('id="lin-tools"', 1)[1]
        tools = bar.split("</header>", 1)[0]
        for control in (
            "lin-search",
            "lin-sheet-filter",
            "lin-lay-dagre",
            "lin-lay-fcose",
            "lin-zoom-in",
            "lin-zoom-out",
            "lin-fit",
            "lin-fit-sel",
        ):
            assert control in tools, control

    def test_the_tools_are_hidden_away_from_the_graph_tab(self):
        html = render_html(demo_graph())
        assert (
            "document.getElementById('lin-tools').hidden = "
            "activeMain.id !== 'lin-graph-main';" in html
        )
        assert ".lin-tools[hidden] { display: none; }" in html

    def test_fit_selection_toggles_through_the_hidden_property(self):
        html = render_html(demo_graph())
        assert "fitSelBtn.hidden = cy.nodes(':selected').length === 0;" in html
        assert 'id="lin-fit-sel" hidden' in html


class TestTablistSemantics:
    def test_the_tabs_declare_the_aria_tab_pattern(self):
        html = render_html(demo_graph())
        assert 'role="tablist"' in html
        for key in ("graph", "overview", "sheets", "screenshots"):
            assert f'id="lin-tab-{key}" class="lin-tab" role="tab"' in html
            assert f'aria-controls="lin-{key}-main"' in html
            assert f'id="lin-{key}-main" role="tabpanel" ' in html or (
                f'id="lin-{key}-main" role="tabpanel"' in html
            )

    def test_selection_moves_aria_selected_and_the_roving_tabindex(self):
        html = render_html(demo_graph())
        assert "btn.setAttribute('aria-selected', on ? 'true' : 'false');" in html
        assert "btn.tabIndex = on ? 0 : -1;" in html
        assert '.lin-tab[aria-selected="true"] {' in html

    def test_the_arrow_keys_walk_the_visible_tabs(self):
        html = render_html(demo_graph())
        assert "function setupTablistKeys()" in html
        assert "setupTablistKeys();" in html
        for key in ("ArrowRight", "ArrowLeft", "Home", "End"):
            assert f"e.key === '{key}'" in html
        assert "return !b.hidden;" in html  # a hidden tab is not a stop

    def test_the_layout_and_page_buttons_report_their_state(self):
        html = render_html(demo_graph())
        assert "on.setAttribute('aria-pressed', 'true');" in html
        assert "child.setAttribute('aria-current', 'false');" in html


class TestAccessibleNames:
    def test_every_icon_only_control_is_named(self):
        html = render_html(demo_graph())
        assert "named('lin-zoom-in', _t('zoom_in'));" in html
        assert "named('lin-zoom-out', _t('zoom_out'));" in html
        assert "node.setAttribute('aria-label', text);" in html
        assert "close.setAttribute('aria-label', _t('close'));" in html

    def test_the_detail_panel_is_a_named_region(self):
        html = render_html(demo_graph())
        assert 'id="lin-panel" role="region" tabindex="-1"' in html
        assert "panel.setAttribute('aria-label', _t('details_panel'));" in html

    def test_focus_is_visible(self):
        html = render_html(demo_graph())
        assert ".lin-root :focus-visible { outline: 2px solid var(--focus)" in html

    def test_the_search_box_moves_its_focus_ring_rather_than_dropping_it(self):
        """The input's own outline is suppressed only because the wrapper rings."""
        html = render_html(demo_graph())
        assert ".lin-search input:focus-visible { outline: none; }" in html
        ring = "border-color: var(--focus); box-shadow: 0 0 0 3px var(--focus-soft);"
        assert ring in html

    def test_navigating_a_precedent_keeps_focus_in_the_document(self):
        """The clicked button is destroyed by the re-render; focus must land."""
        html = render_html(demo_graph())
        assert "document.getElementById('lin-panel').focus();" in html

    def test_the_badge_ink_is_computed_rather_than_assumed(self):
        """White on #eda100 was 2:1; the pair is derived from the fill now."""
        html = render_html(demo_graph())
        assert "function onColor(hex)" in html
        assert ".lin-badge { font-size:" in html  # no blanket color: #fff

    def test_the_link_ink_is_not_the_node_fill(self):
        html = render_html(demo_graph())
        assert "color: var(--link); cursor: pointer;" in html
        assert "--link: #1b62b8;" in html


class TestDarkTheme:
    """Dark exists, but it is opt-in.

    Excel is a light application and its users work in light; a report that
    opened dark because the reader's OS happened to be dark was the wrong
    default for this audience.
    """

    def test_light_is_the_default_whatever_the_os_prefers(self):
        html = render_html(demo_graph())
        assert "@media (prefers-color-scheme" not in html
        assert "matchMedia" not in script_of(html)
        assert "<meta name='color-scheme' content='light'>" in html
        assert '<div class="lin-root" data-theme="light">' in html

    def test_dark_is_a_state_on_the_root_rather_than_a_media_query(self):
        html = render_html(demo_graph())
        assert '.lin-root[data-theme="dark"] {' in html
        assert "color-scheme: dark;" in html

    def test_the_toggle_lives_outside_the_graph_only_tools(self):
        """It must stay reachable on the tabs where .lin-tools is hidden."""
        html = render_html(demo_graph())
        assert '<div class="lin-bar-right">' in html
        between = html.split('id="lin-tools"', 1)[1].split('id="lin-theme"', 1)[0]
        # The tools container has closed by the time the toggle appears, so the
        # `hidden` that blanks .lin-tools on the other tabs cannot reach it.
        assert between.count("</div>") == between.count("<div") + 1
        assert "named('lin-theme', _t('theme_dark'));" in html

    def test_the_toggle_reports_its_state(self):
        html = render_html(demo_graph())
        assert 'id="lin-theme" class="lin-theme" aria-pressed="false"' in html
        assert (
            "if (btn) btn.setAttribute('aria-pressed', "
            "theme === 'dark' ? 'true' : 'false');" in html
        )

    def test_storage_failure_is_swallowed_rather_than_thrown(self):
        """localStorage throws in the sandboxed data: iframe notebooks use."""
        script = script_of(render_html(demo_graph()))
        assert (
            "try { return window.localStorage.getItem(THEME_KEY); } "
            "catch (e) { return null; }" in script
        )
        assert (
            "try { window.localStorage.setItem(THEME_KEY, value); } "
            "catch (e) { /* not available */ }" in script
        )

    def test_the_meta_is_kept_honest_with_the_current_state(self):
        script = script_of(render_html(demo_graph()))
        lookup = "var meta = document.querySelector('meta[name=\"color-scheme\"]');"
        assert lookup in script
        assert "if (meta) meta.setAttribute('content', theme);" in script

    def test_the_dark_block_redefines_the_surfaces_and_the_ink(self):
        dark = render_html(demo_graph()).split('.lin-root[data-theme="dark"] {', 1)[1]
        for token in ("--surface:", "--ink:", "--ink2:", "--code-bg:", "--shadow:"):
            assert token in dark.split("</style>", 1)[0], token

    def test_the_canvas_reads_the_tokens_back(self):
        """Cytoscape paints to a canvas and cannot resolve var()."""
        html = render_html(demo_graph())
        assert "function themeInk()" in html
        assert "color: t.ink2, 'border-width': 1.5" in html
        assert "'border-color': t.nodeBorder" in html
        assert "cy.style(buildStyle(big));" in html
        assert "themeHooks.push(function () { cy.style(buildStyle(big)); });" in html
        assert "for (var i = 0; i < themeHooks.length; i++) themeHooks[i]();" in html

    def test_the_sheet_badges_no_longer_carry_hard_coded_pairs(self):
        """They were written per call site, so dark mode could not reach them."""
        script = script_of(render_html(demo_graph()))
        for gone in ("#e5dbff", "#e3faf2", "#fff0f6", "#e8f7ff", "#d6336c", "#0ca678"):
            assert gone not in script, gone
        assert "metaBadge(metaRow, _t('visibility')" in script
        assert "PALETTE.violet" in script


class TestResponsivePanel:
    def test_the_panel_is_wide_enough_for_a_formula_and_two_readings(self):
        """340px was cramped once the value card carried two columns."""
        html = render_html(demo_graph())
        assert "width: 440px; flex-shrink: 0;" in html

    def test_the_panel_overlays_below_900px(self):
        narrow = render_html(demo_graph()).split("@media (max-width: 900px) {", 1)[1]
        assert "position: absolute;" in narrow
        assert "width: min(440px, 92vw);" in narrow

    def test_the_two_readings_stack_once_the_panel_is_an_overlay(self):
        narrow = render_html(demo_graph()).split("@media (max-width: 560px) {", 1)[1]
        assert ".lin-vgrid { grid-template-columns: 1fr; }" in narrow

    def test_the_empty_panel_stands_down_when_it_would_cover_the_graph(self):
        html = render_html(demo_graph())
        assert ".lin-panel.lin-empty { display: none; }" in html
        assert "p.classList.add('lin-empty');" in html
        assert "p.classList.remove('lin-empty');" in html

    def test_a_populated_panel_keeps_its_close_button(self):
        html = render_html(demo_graph())
        assert "var close = el('button', 'lin-close', '✕');" in html
        assert "close.onclick = function () { clearSel(cy); };" in html

    def test_the_sheets_columns_stack(self):
        narrow = render_html(demo_graph()).split("@media (max-width: 900px) {", 1)[1]
        assert ".lin-sheets-wrap { flex-direction: column; }" in narrow


class TestNewInterfaceStrings:
    NEW_KEYS = (
        "close",
        "layout_flow",
        "layout_organic",
        "details_panel",
        "sheet_dims",
        "visibility",
        "freeze_panes",
        "hidden_columns",
        "merged_ranges",
        "comments",
    )

    def test_every_language_defines_them(self):
        for language in LANGUAGES:
            missing = set(self.NEW_KEYS) - set(UI_STRINGS[language])
            assert not missing, f"{language} is missing {missing}"

    def test_the_dimension_line_keeps_both_placeholders(self):
        for language in LANGUAGES:
            text = UI_STRINGS[language]["sheet_dims"]
            assert "{rows}" in text and "{cols}" in text, language

    def test_the_sheet_pane_no_longer_hard_codes_english(self):
        html = render_html(demo_graph())
        for literal in ("' rows × '", "'❄ Freeze: '", "Hidden cols: ", "Comments ("):
            assert literal not in html, literal
        assert "_t('sheet_dims'" in html
        assert "_t('comments')" in html

    def test_the_layout_buttons_are_translated(self):
        html = render_html(demo_graph(), language="fr")
        assert UI_STRINGS["fr"]["layout_flow"] in html
        assert UI_STRINGS["fr"]["layout_organic"] in html
        assert "label('lin-lay-dagre', _t('layout_flow'));" in html

    def test_interpolation_does_not_read_the_value_as_a_capture_group(self):
        """`{recalc}` carries a computed cell value, which may contain `$&`.

        `String.replace` with a *string* replacement expands `$&` and `$1` into
        capture references, so such a value would blank itself out and leave the
        sentence mangled. A function replacement is the only form that does not.
        """
        html = render_html(demo_graph())
        assert "function () { return replacements[k]; }" in html
        assert "str.replace('{' + k + '}', replacements[k])" not in html
