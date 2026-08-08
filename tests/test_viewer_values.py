"""Tests for how the viewer surfaces value provenance, dates and the result.

These build graph dictionaries by hand instead of going through ``analyze``:
the point here is the *rendering* contract — which strings, data and wiring end
up in the generated document — not the value resolution that produces them.

The panel itself is built client-side, so a Python-level test can only assert
on what the document ships: the graph JSON, the localized strings, and the
template logic that binds one to the other. Each test therefore checks the data
*and* the code path that consumes it.
"""

import base64

from linexcel.i18n import UI_STRINGS
from linexcel.viewer import render_html, wrap_iframe

EN = UI_STRINGS["en"]
FR = UI_STRINGS["fr"]


def graph(**node_fields) -> dict:
    """A one-node graph carrying ``node_fields`` on a formula cell."""
    node = {
        "id": "S!A1",
        "label": "S!A1",
        "kind": "cell",
        "formula": "=B1/C1",
        "value": 1,
    }
    node.update(node_fields)
    return {"nodes": [node], "edges": []}


class TestValueProvenance:
    def test_file_sourced_value_ships_the_read_from_file_label(self):
        html = render_html(graph(valueSource="file", cachedValue=1))
        assert EN["value_from_file"] == "Read from Excel file"
        assert EN["value_from_file"] in html
        assert '"valueSource": "file"' in html
        assert "file:     { labelKey: 'value_from_file'" in html

    def test_fallback_sourced_value_ships_the_fallback_label(self):
        html = render_html(graph(valueSource="fallback", value=0))
        assert EN["value_fallback"] == "linexcel fallback (error-guarded)"
        assert EN["value_fallback"] in html
        assert '"valueSource": "fallback"' in html
        assert "fallback: { labelKey: 'value_fallback'" in html

    def test_engine_sourced_value_ships_the_recalculated_label(self):
        html = render_html(graph(valueSource="engine"))
        assert EN["value_recalc"] in html
        assert "engine:   { labelKey: 'value_recalc'" in html

    def test_the_reading_is_swatched_with_the_color_of_its_source(self):
        html = render_html(graph(valueSource="engine"))
        assert "color: PALETTE.blue" in html  # engine
        assert "color: PALETTE.green" in html  # file
        assert "color: PALETTE.amber" in html  # fallback
        # the palette still ships the literal hex the canvas needs
        assert "'#1baf7a'" in html
        assert "'#eda100'" in html
        assert "sw.style.background = color;" in html

    def test_the_badge_ink_is_derived_from_the_fill(self):
        """White-on-amber failed contrast; the pair is now computed."""
        html = render_html(graph(valueSource="fallback"))
        assert "function onColor(hex)" in html
        assert "node.style.color = onColor(hex);" in html

    def test_each_state_carries_a_sentence_of_its_own(self):
        """A badge alone left 'engine' and 'file' looking interchangeable."""
        html = render_html(graph(valueSource="file"))
        for key in ("value_recalc_desc", "value_from_file_desc", "value_fallback_desc"):
            assert f"descKey: '{key}'" in html, key
        assert EN["value_from_file_desc"] in html
        assert "if (src) sv.appendChild(el('p', 'lin-hint', _t(src.descKey)));" in html

    def test_a_guarded_fallback_is_marked_on_the_card(self):
        html = render_html(graph(valueSource="fallback"))
        mark = (
            "if (n.valueSource === 'fallback') card.className = 'lin-vcard is-guarded';"
        )
        assert mark in html
        assert ".lin-vcard.is-guarded { border-color: var(--amber); }" in html


class TestValueDate:
    def test_the_date_string_reaches_the_document(self):
        html = render_html(graph(valueDate="2026-08-07", value=46236))
        assert "2026-08-07" in html
        assert '"valueDate": "2026-08-07"' in html

    def test_the_panel_prefers_the_date_over_the_raw_serial(self):
        html = render_html(graph(valueDate="2026-08-07", value=46236))
        assert "var shown = n.valueDate || fmt(n.value);" in html


class TestReadVersusRecalculated:
    """Reading the workbook's own values back is the headline feature.

    The panel used to show one headline figure and only mention the file's
    stored value when it *differed*, so in the common case the reader could not
    tell whose number they were looking at. Both readings are now shown,
    labelled, with the verdict stated either way.
    """

    def test_both_readings_are_shown_side_by_side(self):
        html = render_html(graph(value=42, cachedValue=41, valueSource="engine"))
        assert '"value": 42' in html
        assert '"cachedValue": 41' in html
        left = "reading(_t('value_col_file'), cached, SRC.file.color)"
        right = "reading(_t('value_col_calc'), shown, SRC.engine.color)"
        assert left in html
        assert right in html
        assert EN["value_col_file"] == "Excel file"
        assert EN["value_col_calc"] in html

    def test_agreement_is_stated_rather_than_left_silent(self):
        html = render_html(graph(value=42, cachedValue=42, valueSource="engine"))
        assert EN["value_match"] == "The recalculated value matches the file"
        assert EN["value_match"] in html
        assert "_t(agree ? 'value_match' : 'value_mismatch')" in html

    def test_a_disagreement_is_loud(self):
        html = render_html(graph(value=42, cachedValue=41, valueSource="engine"))
        assert EN["value_mismatch"] in html
        assert "if (!agree) card.className = 'lin-vcard is-diff';" in html
        assert ".lin-vcard.is-diff .lin-verdict {" in html
        assert "background: var(--warn-bg); color: var(--warn-fg);" in html

    def test_the_verdict_is_not_carried_by_colour_alone(self):
        """Green vs orange is the classic CVD trap: icon and sentence too."""
        html = render_html(graph(value=42, cachedValue=41, valueSource="engine"))
        assert "verdict.appendChild(icon(agree ? 'check' : 'warn'));" in html
        assert "check: 'M2.6 8.4 L6.3 12 L13.4 4.2'" in html

    def test_a_single_reading_is_never_dressed_up_as_a_corroboration(self):
        """Under 'file' the shown value *is* the stored one, not a second read."""
        html = render_html(graph(value=42, cachedValue=42, valueSource="file"))
        assert (
            "var paired = hasCached && hasValue && n.valueSource === 'engine';" in html
        )
        assert "grid.className = 'lin-vgrid is-single';" in html

    def test_a_date_node_compares_the_cached_date_part_only(self):
        # a file-cached "2026-08-07 00:00:00" is the same day as valueDate:
        # the comparison must not fire on the time suffix
        html = render_html(
            graph(valueDate="2026-08-07", value=46236, cachedValue="2026-08-07")
        )
        assert "n.cachedValue.slice(0, 10) === n.valueDate;" in html
        assert "var agree = !!sameDate || cached === shown;" in html
        assert '"cachedValue": "2026-08-07"' in html

    def test_a_genuinely_different_cached_date_still_flags(self):
        html = render_html(
            graph(valueDate="2026-08-07", value=46236, cachedValue="2026-08-06")
        )
        assert '"cachedValue": "2026-08-06"' in html
        assert EN["value_mismatch"] in html

    def test_the_section_is_no_longer_titled_as_if_everything_were_computed(self):
        html = render_html(graph(valueSource="file"))
        assert EN["value_heading"] == "Value"
        assert "var sv = section(_t('value_heading'));" in html


class TestSamples:
    def test_samples_carry_their_own_date_and_source(self):
        html = render_html(
            graph(
                samples=[
                    {"addr": "S!A2", "value": 46236, "date": "2026-08-08"},
                    {"addr": "S!A3", "value": 7, "valueSource": "file"},
                ]
            )
        )
        assert '"date": "2026-08-08"' in html
        assert '"valueSource": "file"' in html
        assert "if (s.date) line += ' (' + s.date + ')';" in html
        assert "SRC[ssrc] && ssrc !== 'engine'" in html

    def test_both_spellings_of_the_sample_source_field_are_read(self):
        """The graph names it valueSource; older graphs said source."""
        html = render_html(graph(samples=[{"addr": "S!A3", "value": 7}]))
        assert "var ssrc = s.valueSource || s.source;" in html


class TestErrorValues:
    def test_a_div_error_step_renders_as_the_excel_error_text(self):
        steps = {
            "label": "/",
            "expr": "B1/C1",
            "evaluated": True,
            "value": {"type": "Error", "kind": "Div"},
            "children": [],
            "inputs": [{"ref": "B1", "value": 1}],
        }
        html = render_html(graph(steps=steps))
        assert "#DIV/0!" in html
        assert '"kind": "Div"' in html
        assert "v.type === 'Error' && ERRTEXT[v.kind]" in html

    def test_every_excel_error_kind_is_mapped(self):
        html = render_html(graph())
        errors = ("#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!")
        for text in errors:
            assert text in html

    def test_an_unknown_error_kind_falls_back_to_the_raw_dict(self):
        html = render_html(graph())
        assert "return JSON.stringify(v);" in html


class TestFinalResult:
    def test_the_root_step_is_labeled_as_the_final_result(self):
        html = render_html(
            graph(
                steps={
                    "label": "SUM",
                    "expr": "SUM(B1:B3)",
                    "evaluated": True,
                    "value": 6,
                    "children": [
                        {
                            "label": "range",
                            "expr": "B1:B3",
                            "evaluated": True,
                            "value": 6,
                            "children": [],
                        }
                    ],
                }
            )
        )
        assert EN["final_result"] == "Final result"
        assert EN["final_result"] in html
        assert "d.appendChild(el('div', 'lin-val', _t('final_result')))" in html

    def test_only_the_root_step_gets_the_final_styling(self):
        html = render_html(graph())
        assert "'lin-step' + (depth === 0 ? ' lin-final' : '')" in html
        assert ".lin-step.lin-final { border-left-color: var(--green); }" in html


class TestLocalization:
    def test_the_french_render_ships_the_french_labels(self):
        html = render_html(graph(valueSource="file", cachedValue=2), language="fr")
        assert FR["value_from_file"] == "Lue du fichier Excel"
        for key in (
            "value_from_file",
            "value_recalc",
            "value_fallback",
            "value_col_file",
            "value_col_calc",
            "value_match",
            "value_mismatch",
            "final_result",
        ):
            assert FR[key] in html

    def test_the_english_render_does_not_ship_the_french_labels(self):
        html = render_html(graph(valueSource="file"))
        assert FR["value_from_file"] not in html
        assert FR["final_result"] not in html

    def test_every_language_defines_the_new_keys(self):
        new_keys = {
            "value_heading",
            "value_from_file",
            "value_recalc",
            "value_fallback",
            "value_recalc_desc",
            "value_from_file_desc",
            "value_fallback_desc",
            "value_col_file",
            "value_col_calc",
            "value_match",
            "value_mismatch",
            "final_result",
        }
        for language, strings in UI_STRINGS.items():
            assert new_keys <= set(strings), f"{language} is missing new value keys"

    def test_the_two_column_headers_stay_short_enough_to_sit_side_by_side(self):
        """They share a 440px panel: a paragraph in a column header wraps ugly."""
        for language, strings in UI_STRINGS.items():
            for key in ("value_col_file", "value_col_calc"):
                assert len(strings[key]) <= 28, f"{language}/{key} is too long"


def two_sheet_graph() -> dict:
    """A graph spanning ``Inputs`` and ``Calc``, wired across the two."""
    return {
        "nodes": [
            {
                "id": "i:Inputs!A1",
                "label": "Inputs!A1",
                "kind": "input",
                "sheet": "Inputs",
                "value": 3,
            },
            {
                "id": "c:Calc!B1",
                "label": "Calc!B1",
                "kind": "cell",
                "sheet": "Calc",
                "formula": "=Inputs!A1*2",
                "value": 6,
            },
            {
                "id": "x:<external>",
                "label": "<external>",
                "kind": "opaque",
                "sheet": None,
            },
        ],
        "edges": [
            {"id": "e1", "source": "i:Inputs!A1", "target": "c:Calc!B1", "kind": "ref"},
            {
                "id": "e2",
                "source": "x:<external>",
                "target": "c:Calc!B1",
                "kind": "ref",
            },
        ],
    }


class TestSheetFilter:
    def test_the_toolbar_ships_a_sheet_filter_select(self):
        html = render_html(two_sheet_graph())
        assert '<select id="lin-sheet-filter"' in html

    def test_every_sheet_of_the_graph_gets_an_option(self):
        html = render_html(two_sheet_graph())
        assert '<option value="Inputs">Inputs</option>' in html
        assert '<option value="Calc">Calc</option>' in html

    def test_nodes_without_a_sheet_get_no_option(self):
        html = render_html(two_sheet_graph())
        assert '<option value="">' not in html
        assert "<option" in html.split('id="lin-sheet-filter"', 1)[1]

    def test_the_default_option_is_the_localized_all_sheets_label(self):
        html = render_html(two_sheet_graph())
        assert EN["all_sheets"] == "All sheets"
        assert f'<option value="__all__">{EN["all_sheets"]}</option>' in html

    def test_the_french_render_ships_the_french_all_sheets_label(self):
        html = render_html(two_sheet_graph(), language="fr")
        assert FR["all_sheets"] == "Toutes les feuilles"
        assert f'<option value="__all__">{FR["all_sheets"]}</option>' in html
        assert FR["sheet_filter"] == "Filtre par feuille"
        assert FR["sheet_filter"] in html

    def test_the_select_is_wired_to_the_filter_function(self):
        html = render_html(two_sheet_graph())
        assert "function applySheetFilter(cy, sheet, layoutName, hasFcose)" in html
        assert "sheetSel.onchange = function () {" in html
        assert "applySheetFilter(cy, sheetSel.value, curLayout, hasFcose);" in html

    def test_cross_sheet_neighbours_are_kept_and_dimmed(self):
        html = render_html(two_sheet_graph())
        assert "var keep = inSheet.union(inSheet.neighborhood().nodes());" in html
        assert "function dimExternal(cy)" in html
        assert "}).addClass('dimmed');" in html

    def test_edges_survive_only_between_two_visible_nodes(self):
        html = render_html(two_sheet_graph())
        assert "return keep.contains(e.source()) && keep.contains(e.target());" in html
        assert "keep.show(); keepEdges.show();" in html

    def test_the_sentinel_restores_the_whole_graph(self):
        html = render_html(two_sheet_graph())
        assert "var ALL_SHEETS = '__all__';" in html
        assert "if (CUR_SHEET === ALL_SHEETS) { cy.elements().show(); return; }" in html

    def test_the_filter_relayouts_the_visible_graph(self):
        html = render_html(two_sheet_graph())
        assert "visible.layout(layoutOpts(layoutName, hasFcose)).run();" in html
        assert (
            "if (cy.nodes(':visible').length !== before) cy.fit(visible, 40);" in html
        )

    def test_the_nodes_carry_their_sheet_into_cytoscape(self):
        html = render_html(two_sheet_graph())
        assert "degree: degreeMap[n.id] || 0, sheet: n.sheet || ''" in html

    def test_the_filter_is_disabled_without_cytoscape(self):
        html = render_html(two_sheet_graph())
        assert "sheetSel.disabled = true;" in html

    def test_a_sheet_name_is_escaped_into_the_option(self):
        html = render_html(
            {
                "nodes": [
                    {
                        "id": "c:X!A1",
                        "label": "X!A1",
                        "kind": "cell",
                        "sheet": '<b>"Q1"</b>',
                    },
                ],
                "edges": [],
            }
        )
        assert '<option value="&lt;b&gt;&quot;Q1&quot;&lt;/b&gt;">' in html
        assert '<option value="<b>' not in html

    def test_every_language_defines_the_sheet_filter_keys(self):
        for language, strings in UI_STRINGS.items():
            assert {"all_sheets", "sheet_filter"} <= set(strings), (
                f"{language} is missing the sheet filter keys"
            )


def search_graph() -> dict:
    """Three nodes, two of which carry the token ``marge``."""
    return {
        "nodes": [
            {
                "id": "c:S!A1",
                "label": "S!A1",
                "kind": "cell",
                "sheet": "S",
                "formula": "=Marge*2",
            },
            {
                "id": "g:S!B2#4",
                "label": "Marge nette",
                "kind": "group",
                "sheet": "S",
                "formula": "=B1-C1",
            },
            {
                "id": "c:S!C1",
                "label": "S!C1",
                "kind": "cell",
                "sheet": "S",
                "formula": "=SUM(D1:D9)",
            },
        ],
        "edges": [],
    }


class TestSearchAll:
    """Enter in the search box must reach every match, not only the first.

    The handler runs client-side, so these assert on the shipped template: the
    predicate now sweeps the whole node list and the result drives a multi-node
    selection framed by one fit.
    """

    def test_the_predicate_sweeps_every_node(self):
        html = render_html(search_graph())
        assert "var matched = GRAPH.nodes.filter(function (n) {" in html
        assert "GRAPH.nodes.find(function (n) {" not in html

    def test_the_predicate_still_reads_label_and_formula(self):
        html = render_html(search_graph())
        assert "return (n.label || '').toLowerCase().indexOf(q) >= 0" in html
        assert "|| (n.formula || '').toLowerCase().indexOf(q) >= 0;" in html

    def test_the_searchable_text_of_every_node_ships(self):
        html = render_html(search_graph())
        assert '"label": "Marge nette"' in html
        assert '"formula": "=Marge*2"' in html
        assert '"formula": "=SUM(D1:D9)"' in html

    def test_a_single_match_still_selects_and_opens_the_panel(self):
        """The first hit keeps the old path: select() drives panel + highlight."""
        html = render_html(search_graph())
        assert "select(cy, matched[0].id);" in html
        assert "function select(cy, id)" in html

    def test_every_other_match_joins_the_selection(self):
        html = render_html(search_graph())
        assert "for (var i = 0; i < matched.length; i++) {" in html
        assert "ele.select(); ele.removeClass('dimmed');" in html
        assert "cy.batch(function () {" in html

    def test_the_view_frames_the_whole_matched_set(self):
        html = render_html(search_graph())
        fit = "cy.animate({ fit: { eles: eles, padding: 40 }, duration: 300 });"
        assert fit in html

    def test_an_empty_result_leaves_the_graph_alone_and_says_so(self):
        html = render_html(search_graph())
        guard = "if (!matched.length) { setStatus(_t('search_none'), true); return; }"
        assert guard in html
        assert EN["search_none"] == "No matches"
        assert EN["search_none"] in html

    def test_the_count_is_visible_feedback_not_a_tooltip(self):
        html = render_html(search_graph())
        assert "search.title = _t('search_matches'" not in html
        assert (
            "setStatus(_t('search_matches', { count: matched.length }), false);" in html
        )
        assert EN["search_matches"] == "Matches: {count}"
        assert EN["search_matches"] in html

    def test_the_status_is_a_live_region_that_is_shown_before_it_is_written(self):
        """A region still `hidden` when its text changes is never announced."""
        html = render_html(search_graph())
        assert (
            '<span class="lin-search-status" id="lin-search-status" '
            'role="status" aria-live="polite" hidden>' in html
        )
        body = html.split("function setStatus(text, none) {", 1)[1]
        assert body.index("searchStatus.hidden = !text;") < body.index(
            "searchStatus.textContent = text || '';"
        )

    def test_the_french_render_ships_the_french_match_count(self):
        html = render_html(search_graph(), language="fr")
        assert FR["search_matches"] == "Résultats : {count}"
        assert FR["search_matches"] in html

    def test_every_language_defines_the_search_count_key(self):
        for language, strings in UI_STRINGS.items():
            assert "search_matches" in strings, f"{language} is missing search_matches"
            assert "{count}" in strings["search_matches"], (
                f"{language} lost the count placeholder"
            )

    def test_no_language_writes_a_bare_count_before_a_plural_noun(self):
        """The count is on screen now, so "1 matches" would be visible."""
        for language, strings in UI_STRINGS.items():
            if language in ("ja", "zh"):
                continue  # no plural agreement to get wrong
            assert not strings["search_matches"].startswith("{count} "), language


class TestSearchControl:
    """The box used to be a bare input explaining itself with a `⏎` glyph."""

    def test_the_affordance_is_inline_svg_rather_than_a_glyph(self):
        html = render_html(search_graph())
        box = html.split('<div class="lin-search" id="lin-searchbox">', 1)[1]
        box = box.split("</div>", 1)[0]
        assert "<svg viewBox=" in box
        assert "⏎" not in html
        for language, strings in UI_STRINGS.items():
            assert "⏎" not in strings["search"], language

    def test_the_clear_button_appears_only_once_there_is_text(self):
        html = render_html(search_graph())
        assert 'id="lin-search-clear" class="lin-search-clear" hidden' in html
        assert "searchClear.hidden = !search.value;" in html
        assert ".lin-search-clear[hidden] { display: none; }" in html
        assert "named('lin-search-clear', _t('search_clear'));" in html

    def test_escape_clears_and_restores_the_graph(self):
        html = render_html(search_graph())
        esc = "else if (e.key === 'Escape') { e.preventDefault(); restoreGraph(); }"
        assert esc in html
        assert "function restoreGraph() {" in html
        assert "resetSearch();" in html
        assert "clearSel(cy);" in html

    def test_enter_still_submits(self):
        html = render_html(search_graph())
        assert "if (e.key === 'Enter') { e.preventDefault(); runSearch(); }" in html

    def test_the_box_is_named_rather_than_relying_on_the_placeholder(self):
        html = render_html(search_graph())
        assert "search.setAttribute('aria-label', _t('search_label'));" in html
        assert EN["search_label"] == "Search cells and formulas (Enter)"
        assert EN["search_label"] in html

    def test_the_name_still_says_how_to_submit_now_that_the_glyph_is_gone(self):
        """The placeholder used to read "Search… ⏎"; the hint moved, not died."""
        for language, strings in UI_STRINGS.items():
            # ja/zh close the aside with the full-width parenthesis
            assert strings["search_label"].rstrip().endswith((")", "）")), language

    def test_a_report_without_cytoscape_disables_the_box(self):
        html = render_html(search_graph())
        assert "search.disabled = true;" in html

    def test_every_language_defines_the_new_search_keys(self):
        for language, strings in UI_STRINGS.items():
            missing = {"search_label", "search_none", "search_clear"} - set(strings)
            assert not missing, f"{language} is missing {missing}"


class TestIframeWrapping:
    def test_the_labels_survive_the_base64_iframe_wrapper(self):
        html = render_html(graph(valueSource="file", valueDate="2026-08-07"))
        iframe = wrap_iframe(html)
        payload = iframe.split("base64,", 1)[1].split('"', 1)[0]
        decoded = base64.b64decode(payload).decode("utf-8")
        assert EN["value_from_file"] in decoded
        assert "2026-08-07" in decoded
