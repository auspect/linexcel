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

    def test_the_graph_tools_float_over_the_canvas(self):
        """The toolbar holds only what makes sense above a graph; the search
        stays in the header, the layout switch lives in the rail."""
        html = render_html(demo_graph())
        tools = html.split('id="lin-tools"', 1)[1].split('id="lin-legend"', 1)[0]
        for control in (
            "lin-sheet-filter",
            "lin-zoom-in",
            "lin-zoom-out",
            "lin-fit",
            "lin-fit-sel",
        ):
            assert control in tools, control
        header = html.split("</header>", 1)[0]
        assert "lin-search" in header
        assert "lin-sheet-filter" not in header
        rail = html.split('class="lin-rail"', 1)[1].split('id="lin-graph-main"', 1)[0]
        for control in ("lin-lay-dagre", "lin-lay-fcose", "lin-tablist"):
            assert control in rail, control

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


class TestLeftRail:
    def test_the_rail_groups_views_kinds_and_layout(self):
        html = render_html(demo_graph())
        assert '<nav class="lin-rail">' in html
        for rid in (
            "lin-rail-views",
            "lin-rail-kinds-label",
            "lin-rail-layout-label",
            "lin-rail-kinds",
        ):
            assert f'id="{rid}"' in html, rid

    def test_the_rail_labels_are_translated(self):
        html = render_html(demo_graph(), language="fr")
        assert "_t('rail_views')" in html
        assert "_t('rail_kinds')" in html
        assert "_t('rail_layout')" in html

    def test_kind_filters_toggle_a_display_none_class(self):
        html = render_html(demo_graph())
        assert "function buildKindFilters(" in html
        assert "buildKindFilters(cy, layoutOpts" in html
        assert "{ selector: '.kind-off', style: { 'display': 'none' } }" in html
        assert "node.addClass('kind-off'); else node.removeClass('kind-off');" in html


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
        """It must stay reachable on the tabs where .lin-tools is hidden:
        the toggle sits in the header, the tools float over the graph canvas."""
        html = render_html(demo_graph())
        assert '<div class="lin-bar-right">' in html
        header = html.split("</header>", 1)[0]
        canvas = html.split("</header>", 1)[1]
        assert 'id="lin-theme"' in header
        assert 'id="lin-tools"' not in header
        assert 'id="lin-tools"' in canvas
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
        "rail_views",
        "rail_kinds",
        "rail_layout",
        "shots_empty_title",
        "shots_empty_desc",
        "shots_hint",
        "shots_in_sheets",
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


def sheets_graph(screenshots=None) -> dict:
    """A graph carrying the workbook context the Sheets tab reads."""
    graph = demo_graph()
    graph["meta"] = {
        "workbookContext": {
            "sheets": [
                {
                    "name": "In",
                    "visibility": "visible",
                    "dimensions": {"rows": 2, "columns": 2},
                    "preview_range": "A1:B2",
                    "preview": [
                        {"row": 1, "values": ["Label", "Amount"]},
                        {"row": 2, "values": ["Rent", 1200]},
                    ],
                    "freeze_panes": None,
                    "merged_ranges": [],
                    "hidden_columns": ["B"],
                    "comments": [],
                }
            ]
        }
    }
    if screenshots is not None:
        graph["meta"]["screenshots"] = screenshots
    return graph


class TestSheetsTab:
    """The tab used to hold badges and comments and nothing else.

    The first cells of every sheet were already embedded in the document and
    never drawn, and the rendered image was only ever shown when the caller
    happened to hand in a per-sheet mapping — which nothing produced.
    """

    def test_the_first_cells_of_the_sheet_are_drawn(self):
        script = script_of(render_html(sheets_graph()))
        assert "previewGrid" in script
        assert "lin-grid" in script

    def test_the_preview_is_labelled_with_the_range_it_covers(self):
        script = script_of(render_html(sheets_graph()))
        assert "sheet.preview_range" in script

    def test_a_hidden_column_is_marked_in_the_preview_header(self):
        """The badge says column B is hidden; the grid must agree with it."""
        script = script_of(render_html(sheets_graph()))
        assert "hidden.indexOf(letter)" in script
        assert "is-hidden" in render_html(sheets_graph())

    def test_the_rendered_sheet_has_its_own_heading(self):
        html = render_html(sheets_graph())
        assert EN["sheet_render"] in html
        assert EN["sheet_preview"] in html

    def test_a_flat_page_list_is_not_shown_under_a_sheet(self):
        """Print pages belong to the workbook; only a mapping names sheets."""
        script = script_of(render_html(sheets_graph(["a.png", "b.png"])))
        assert "Array.isArray(screens)" in script

    def test_the_rendered_image_scrolls_inside_a_frame(self):
        """A sheet renders onto one page, so a long one is a very tall image."""
        css = render_html(sheets_graph({"In": ["a.png"]}))
        assert ".lin-frame" in css
        assert "max-height: min(52vh, 460px); overflow: auto" in css

    def test_a_wide_preview_scrolls_rather_than_widening_the_card(self):
        assert ".lin-gridwrap { overflow-x: auto" in render_html(sheets_graph())


class TestScreenshotsTabStates:
    """The Visual preview tab used to vanish whenever the report carried no
    screenshots (the default run), which read as a missing feature rather
    than a missing render. It now always opens, on one of three states:
    the page gallery, a pointer to the per-sheet renders, or an empty state
    that names the command producing them.
    """

    def test_setup_is_called_once_and_before_the_cytoscape_guard(self):
        """A second call before the guard used to double-render the gallery;
        and a call after it left the tab unset in the fallback states."""
        script = script_of(render_html(demo_graph()))
        assert script.count("setupScreenshots();") == 1
        call_at = script.index("setupScreenshots();")
        guard_at = script.index("typeof cytoscape === 'undefined'")
        assert call_at < guard_at

    def test_setup_is_idempotent(self):
        script = script_of(render_html(demo_graph()))
        assert "if (container.firstChild) return;" in script

    def test_the_empty_state_says_so_and_names_the_command(self):
        script = script_of(render_html(sheets_graph()))
        assert "_t('shots_empty_title')" in script
        assert "_t('shots_empty_desc')" in script
        assert "_t('shots_hint')" in script
        html = render_html(sheets_graph())
        assert ".lin-shots-empty {" in html
        assert "--screenshots DIR" in EN["shots_hint"]

    def test_a_per_sheet_mapping_points_to_the_sheets_tab(self):
        """Renders keyed by sheet name live under each sheet's card; the tab
        says where they are instead of duplicating the gallery."""
        script = script_of(render_html(sheets_graph({"In": ["a.png"]})))
        assert "_t('shots_in_sheets')" in script

    def test_the_tab_is_unhidden_in_every_state(self):
        """The early `return` that kept the tab hidden when no flat page list
        existed is gone: the button is revealed before the shape is read."""
        script = script_of(render_html(sheets_graph()))
        unhide_at = script.index("btn.hidden = false;")
        branch_at = script.index("Array.isArray(images)")
        assert unhide_at < branch_at


class TestNodeCards:
    """Nodes are cards now: the label — and a secondary annotation line — sits
    inside the node instead of floating under a dot, and the node is sized
    from its pre-truncated lines so a long label cannot overflow."""

    def test_the_label_is_built_as_a_card_with_a_secondary_line(self):
        script = script_of(render_html(demo_graph()))
        assert "function cardLabel(n)" in script
        assert "function subLabel(n)" in script
        assert "lines.join('\\n')" in script

    def test_the_card_is_sized_from_its_lines(self):
        script = script_of(render_html(demo_graph()))
        assert "id: n.id, label: card.label, w: card.w, h: card.h," in script
        assert "width: 'data(w)', height: 'data(h)'," in script

    def test_every_kind_is_a_rounded_card(self):
        script = script_of(render_html(demo_graph()))
        kinds = script.split("var KIND = {", 1)[1].split("};", 1)[0]
        assert kinds.count("shape: 'round-rectangle'") == 8
        for gone in ("'ellipse'", "'diamond'", "'hexagon'", "'tag'", "'octagon'"):
            assert gone not in kinds, gone

    def test_the_label_ink_is_derived_from_the_card_fill(self):
        script = script_of(render_html(demo_graph()))
        assert "color: onColor(KIND[k].color)" in script

    def test_an_unresolved_external_gets_a_dashed_border(self):
        script = script_of(render_html(demo_graph()))
        assert "function unresolvedExternal(n)" in script
        assert "' unresolved' : ''" in script
        rule = "{ selector: 'node.unresolved', style: { 'border-style': 'dashed',"
        assert rule in script

    def test_the_unresolved_annotation_is_translated(self):
        script = script_of(render_html(demo_graph(), language="fr"))
        assert "_t('external_unresolved')" in script


class TestLabelLevelOfDetail:
    """Labels fade with the zoom band: none but the selection far out, the
    selection and its neighbourhood in the middle band of a crowded graph,
    everything close in."""

    def test_two_zoom_thresholds_frame_the_middle_band(self):
        script = script_of(render_html(demo_graph()))
        assert "var LABEL_MIN_ZOOM = 0.55;" in script
        assert "var LABEL_FULL_ZOOM = 1.6;" in script

    def test_the_middle_band_labels_only_the_neighbourhood(self):
        script = script_of(render_html(demo_graph()))
        assert "sel.closedNeighborhood().nodes()" in script
        assert "z >= LABEL_FULL_ZOOM" in script

    def test_a_small_graph_stays_fully_labelled(self):
        script = script_of(render_html(demo_graph()))
        assert "var CROWDED = GRAPH.nodes.length > 80;" in script


class TestLayoutAutoSwitch:
    """Past a few hundred nodes the organic layout collapses into a hairball,
    so a wide workbook opens on the hierarchical flow. The toggle keeps both
    layouts one click away either way."""

    def test_the_organic_default_is_capped_by_a_node_threshold(self):
        script = script_of(render_html(demo_graph()))
        assert "var ORGANIC_MAX_NODES = 300;" in script
        assert "GRAPH.nodes.length <= ORGANIC_MAX_NODES" in script

    def test_the_manual_toggle_still_reaches_both_layouts(self):
        script = script_of(render_html(demo_graph()))
        fcose = "cy.elements(':visible').layout(layoutOpts('fcose', hasFcose)).run();"
        dagre = "cy.elements(':visible').layout(layoutOpts('dagre', hasFcose)).run();"
        assert fcose in script
        assert dagre in script


class TestDiffsOnlyFilter:
    """The rail's "discrepancies only" toggle keeps just the nodes whose
    recalculated value differs from the file's."""

    def test_the_toggle_sits_in_the_rail_under_its_own_label(self):
        html = render_html(demo_graph())
        assert 'id="lin-rail-filters-label"' in html
        assert 'id="lin-diffs-only" class="lin-kind" aria-pressed="false"' in html

    def test_the_toggle_hides_under_its_own_class(self):
        """A dedicated class, so switching a kind off cannot lift the filter."""
        script = script_of(render_html(demo_graph()))
        assert "{ selector: '.diff-off', style: { 'display': 'none' } }" in script
        assert "node.addClass('diff-off');" in script
        assert "function nodeDiffers(n)" in script

    def test_a_workbook_without_discrepancies_disables_the_toggle(self):
        script = script_of(render_html(demo_graph()))
        assert "if (!count) { btn.disabled = true; return; }" in script

    def test_the_group_filters_stand_down_without_a_graph(self):
        script = script_of(render_html(demo_graph()))
        assert "'lin-rail-filters-label'" in script


class TestFormulaHighlighting:
    def test_the_panel_formula_is_tokenised_and_coloured(self):
        script = script_of(render_html(demo_graph()))
        assert "function highlightFormula(src)" in script
        assert "fcode.innerHTML = highlightFormula(n.formula);" in script

    def test_the_token_colours_are_theme_tokens(self):
        html = render_html(demo_graph())
        tokens = (
            ".lin-tok-fn",
            ".lin-tok-ref",
            ".lin-tok-sheet",
            ".lin-tok-num",
            ".lin-tok-str",
        )
        for cls in tokens:
            assert cls + " {" in html, cls
        assert ".lin-tok-fn { color: var(--ai-fg);" in html


class TestSearchShortcut:
    def test_slash_focuses_the_search_box(self):
        script = script_of(render_html(demo_graph()))
        assert "function setupSearchShortcut()" in script
        assert "setupSearchShortcut();" in script
        assert "e.key !== '/' || e.ctrlKey || e.metaKey || e.altKey" in script

    def test_typing_in_a_field_never_triggers_the_shortcut(self):
        script = script_of(render_html(demo_graph()))
        assert "t.isContentEditable" in script
        assert "t.tagName === 'TEXTAREA'" in script

    def test_the_hint_is_on_the_box_title(self):
        script = script_of(render_html(demo_graph()))
        assert "search.title = _t('search_label') + ' (/)';" in script


class TestDiffsAndCardStrings:
    NEW_KEYS = ("rail_filters", "diffs_only", "external_unresolved")

    def test_every_language_defines_them(self):
        for language in LANGUAGES:
            missing = set(self.NEW_KEYS) - set(UI_STRINGS[language])
            assert not missing, f"{language} is missing {missing}"
