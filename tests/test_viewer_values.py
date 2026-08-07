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

    def test_the_badge_is_rendered_with_a_color_per_source(self):
        html = render_html(graph(valueSource="engine"))
        assert "lin-src" in html
        assert "var(--blue)" in html  # engine
        assert "'#1baf7a'" in html  # file
        assert "'#eda100'" in html  # fallback
        assert "badge.style.background = src.color;" in html


class TestValueDate:
    def test_the_date_string_reaches_the_document(self):
        html = render_html(graph(valueDate="2026-08-07", value=46236))
        assert "2026-08-07" in html
        assert '"valueDate": "2026-08-07"' in html

    def test_the_panel_prefers_the_date_over_the_raw_serial(self):
        html = render_html(graph(valueDate="2026-08-07", value=46236))
        assert "var shown = n.valueDate || fmt(n.value);" in html


class TestCachedValueDiscrepancy:
    def test_both_values_appear_when_they_differ(self):
        html = render_html(graph(value=42, cachedValue=41, valueSource="engine"))
        assert '"value": 42' in html
        assert '"cachedValue": 41' in html
        assert EN["value_cached"] in html
        assert EN["differs_from_file"] in html

    def test_the_second_line_is_emitted_only_on_a_mismatch(self):
        html = render_html(graph(value=42, cachedValue=42))
        assert "if (!sameDate && cached !== shown) {" in html

    def test_a_date_node_compares_the_cached_date_part_only(self):
        # a file-cached "2026-08-07 00:00:00" is the same day as valueDate:
        # the comparison must not fire on the time suffix
        html = render_html(
            graph(valueDate="2026-08-07", value=46236, cachedValue="2026-08-07")
        )
        assert "n.cachedValue.slice(0, 10) === n.valueDate;" in html
        assert '"cachedValue": "2026-08-07"' in html

    def test_a_genuinely_different_cached_date_still_flags(self):
        html = render_html(
            graph(valueDate="2026-08-07", value=46236, cachedValue="2026-08-06")
        )
        assert '"cachedValue": "2026-08-06"' in html
        assert EN["differs_from_file"] in html

    def test_the_discrepancy_message_keeps_its_placeholder(self):
        assert "{recalc}" in EN["differs_from_file"]
        assert "{recalc}" in FR["differs_from_file"]


class TestSamples:
    def test_samples_carry_their_own_date_and_source(self):
        html = render_html(
            graph(
                samples=[
                    {"addr": "S!A2", "value": 46236, "date": "2026-08-08"},
                    {"addr": "S!A3", "value": 7, "source": "file"},
                ]
            )
        )
        assert '"date": "2026-08-08"' in html
        assert '"source": "file"' in html
        assert "if (s.date) line += ' (' + s.date + ')';" in html
        assert "SRC[s.source] && s.source !== 'engine'" in html


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
        assert ".lin-step.lin-final { border-left-color: #1baf7a; }" in html


class TestLocalization:
    def test_the_french_render_ships_the_french_labels(self):
        html = render_html(graph(valueSource="file", cachedValue=2), language="fr")
        assert FR["value_from_file"] == "Lue du fichier Excel"
        for key in (
            "value_from_file",
            "value_recalc",
            "value_fallback",
            "value_cached",
            "final_result",
        ):
            assert FR[key] in html

    def test_the_english_render_does_not_ship_the_french_labels(self):
        html = render_html(graph(valueSource="file"))
        assert FR["value_from_file"] not in html
        assert FR["final_result"] not in html

    def test_every_language_defines_the_new_keys(self):
        new_keys = {
            "value_from_file",
            "value_recalc",
            "value_fallback",
            "value_cached",
            "differs_from_file",
            "final_result",
        }
        for language, strings in UI_STRINGS.items():
            assert new_keys <= set(strings), f"{language} is missing new value keys"


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

    def test_an_empty_result_changes_nothing(self):
        html = render_html(search_graph())
        assert "if (!matched.length) return;" in html

    def test_the_box_reports_how_many_nodes_matched(self):
        html = render_html(search_graph())
        assert "search.title = _t('search_matches', { count: matched.length });" in html
        assert EN["search_matches"] == "{count} matches"
        assert EN["search_matches"] in html

    def test_the_french_render_ships_the_french_match_count(self):
        html = render_html(search_graph(), language="fr")
        assert FR["search_matches"] == "{count} résultats"
        assert FR["search_matches"] in html

    def test_every_language_defines_the_search_count_key(self):
        for language, strings in UI_STRINGS.items():
            assert "search_matches" in strings, f"{language} is missing search_matches"
            assert "{count}" in strings["search_matches"], (
                f"{language} lost the count placeholder"
            )


class TestIframeWrapping:
    def test_the_labels_survive_the_base64_iframe_wrapper(self):
        html = render_html(graph(valueSource="file", valueDate="2026-08-07"))
        iframe = wrap_iframe(html)
        payload = iframe.split("base64,", 1)[1].split('"', 1)[0]
        decoded = base64.b64decode(payload).decode("utf-8")
        assert EN["value_from_file"] in decoded
        assert "2026-08-07" in decoded
