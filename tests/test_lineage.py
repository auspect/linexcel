"""Tests for the lineage module: references, grouping, graph, VBA, API."""

import io
from pathlib import Path
from typing import Any, cast

import pytest

from linexcel import LineageResult, analyze
from linexcel import viewer as viewer_module
from linexcel.analyzer import analyze_workbook
from linexcel.refs import (
    Rect,
    col_to_num,
    num_to_col,
    parse_ref,
    parse_ref_detailed,
    ref_to_r1c1,
    stretch_ref,
)
from linexcel.rewrite import canonical_r1c1, qualify_sheet
from linexcel.vba import analyze_vba


class TestRefs:
    def test_col_roundtrip(self):
        for col in ("A", "Z", "AA", "XFD"):
            assert num_to_col(col_to_num(col)) == col

    def test_parse_cell(self):
        rect = parse_ref("B2", "S")
        assert rect is not None
        assert rect == Rect("S", 2, 2, 2, 2)

    def test_parse_range_with_sheet(self):
        rect = parse_ref("'Ma Feuille'!A1:C10")
        assert rect is not None
        assert rect.sheet == "Ma Feuille"
        assert (rect.r1, rect.c1, rect.r2, rect.c2) == (1, 1, 10, 3)

    def test_parse_whole_column(self):
        rect = parse_ref("A:B")
        assert rect is not None
        assert rect.c1 == 1 and rect.c2 == 2
        assert rect.r1 == 1 and rect.r2 == 1_048_576

    def test_structured_ref_rejected(self):
        assert parse_ref("Table1[Col]") is None
        assert parse_ref("MyName") is None

    def test_r1c1_relative(self):
        assert ref_to_r1c1("A2", 2, 2) == "RC[-1]"
        assert ref_to_r1c1("$A$1:A10", 5, 3) == "R1C1:R[5]C[-2]"

    def test_stretch_relative_follows_group(self):
        detail = parse_ref_detailed("A2", "S")
        assert detail is not None
        rect = stretch_ref(detail, 2, 4, (2, 101), (4, 4))
        assert rect is not None
        assert (rect.r1, rect.r2) == (2, 101)
        assert rect.c1 == rect.c2 == 1

    def test_stretch_anchored_stays_fixed(self):
        detail = parse_ref_detailed("$A$2", "S")
        assert detail is not None
        rect = stretch_ref(detail, 2, 4, (2, 101), (4, 4))
        assert rect is not None
        assert (rect.r1, rect.r2) == (2, 2)


class TestRewrite:
    def test_stretched_formulas_share_canonical_form(self):
        assert canonical_r1c1("A2*2+1", 2, 2) == canonical_r1c1("A9*2+1", 9, 2)

    def test_different_logic_differs(self):
        assert canonical_r1c1("A1", 1, 2) != canonical_r1c1("A1", 2, 2)

    def test_qualify_leaves_names_and_strings(self):
        out = qualify_sheet('SUM(A1:A3)+MonNom&"B2"', "Feuil 1")
        assert "'Feuil 1'!A1:A3" in out
        assert "MonNom" in out and "'Feuil 1'!MonNom" not in out
        assert '"B2"' in out

    def test_qualify_keeps_existing_sheet(self):
        out = qualify_sheet("Data!B2*C3", "S1")
        assert "Data!B2" in out and "S1!C3" in out


class TestAnalyze:
    def test_graph_structure(self, lineage_excel):
        result = analyze_workbook(lineage_excel, "test.xlsx")
        graph = result["graph"]
        stats = graph["meta"]["stats"]
        assert stats["totalFormulas"] == 103
        kinds = {n["kind"] for n in graph["nodes"]}
        assert {"group", "cell", "input", "name"} <= kinds

    def test_stretched_column_becomes_one_group(self, lineage_excel):
        graph = analyze_workbook(lineage_excel, "test.xlsx")["graph"]
        groups = [n for n in graph["nodes"] if n["kind"] == "group"]
        assert len(groups) == 1
        assert groups[0]["count"] == 100
        assert groups[0]["bbox"] == "D2:D101"

    def test_group_inputs_are_aggregated(self, lineage_excel):
        graph = analyze_workbook(lineage_excel, "test.xlsx")["graph"]
        input_labels = {n["label"] for n in graph["nodes"] if n["kind"] == "input"}
        assert "Ventes!B2:B101" in input_labels
        assert "Ventes!C2:C101" in input_labels

    def test_defined_name_resolved(self, lineage_excel):
        graph = analyze_workbook(lineage_excel, "test.xlsx")["graph"]
        names = [n for n in graph["nodes"] if n["kind"] == "name"]
        assert names and names[0]["label"] == "TauxCible"
        # name is fed by Params!A1 and feeds Synthese!B3
        edges = graph["edges"]
        assert any(
            e["target"] == names[0]["id"] and "Params" in e["source"] for e in edges
        )
        assert any(
            e["source"] == names[0]["id"] and e["target"].endswith("Synthese!B3")
            for e in edges
        )

    def test_composed_formula_steps_evaluated(self, lineage_excel):
        graph = analyze_workbook(lineage_excel, "test.xlsx")["graph"]
        node = next(n for n in graph["nodes"] if n["id"].endswith("Synthese!B3"))
        steps = node["steps"]
        assert steps["label"] == "IF"
        assert steps["evaluated"] and steps["value"] == node["value"]
        # the comparison and inner SUM are evaluated individually
        flat = _flatten(steps)
        by_label = {s["label"]: s for s in flat}
        assert by_label[">"]["value"] is True
        assert isinstance(by_label["SUM"]["value"], float)
        assert by_label["ROUND"]["evaluated"]

    def test_values_computed_by_engine(self, lineage_excel):
        graph = analyze_workbook(lineage_excel, "test.xlsx")["graph"]
        b1 = next(n for n in graph["nodes"] if n["id"].endswith("Synthese!B1"))
        assert isinstance(b1["value"], float) and b1["value"] > 0


def _flatten(step):
    out = [step]
    for child in step.get("children", []):
        out.extend(_flatten(child))
    return out


class TestPackageApi:
    """The tool must be usable as a library, without FastAPI or AI."""

    def test_analyze_from_bytes(self, lineage_excel):
        result = analyze(lineage_excel, filename="demo.xlsx")
        assert isinstance(result, LineageResult)
        assert result.stats["totalFormulas"] == 103
        assert "Ventes" in result.sheets

    def test_analyze_from_path(self, tmp_path, lineage_excel):
        path = tmp_path / "workbook.xlsx"
        path.write_bytes(lineage_excel)
        result = analyze(path)
        assert result.stats["totalFormulas"] == 103

    def test_analyze_from_filelike(self, lineage_excel):
        result = analyze(io.BytesIO(lineage_excel), filename="stream.xlsx")
        assert result.stats["totalFormulas"] == 103

    def test_navigation_helpers(self, lineage_excel):
        result = analyze(lineage_excel)
        b3 = result.find("Synthese!B3")
        assert b3 and b3[0]["id"].endswith("Synthese!B3")
        node_id = b3[0]["id"]
        prec_labels = {n["label"] for n in result.precedents(node_id)}
        assert "TauxCible" in prec_labels
        node = result.node(node_id)
        assert node and node["formula"].startswith("=IF")

    def test_to_json_roundtrip(self, lineage_excel):
        import json

        result = analyze(lineage_excel)
        data = json.loads(result.to_json())
        assert data["meta"]["stats"]["totalFormulas"] == 103

    def test_workbook_context_preserves_preview_and_comments(self, lineage_excel):
        result = analyze(lineage_excel, filename="context.xlsx")
        context = result.workbook_context
        ventes = next(sheet for sheet in context["sheets"] if sheet["name"] == "Ventes")
        assert ventes["preview"][0]["values"][:4] == [
            "Produit",
            "Qté",
            "Prix",
            "CA",
        ]
        assert ventes["comments"] == [
            {
                "cell": "A1",
                "author": "Data team",
                "text": "Exported product category",
            }
        ]
        assert ventes["freeze_panes"] == "A2"
        assert ventes["hidden_columns"] == ["C"]
        assert "F1:G1" in ventes["merged_ranges"]

    def test_screenshots_report_a_missing_renderer(
        self, lineage_excel, monkeypatch, tmp_path
    ):
        from linexcel import WorkbookRenderError

        monkeypatch.setattr("linexcel.insights.shutil.which", lambda _: None)
        monkeypatch.setattr("linexcel.insights._launcher_install_paths", lambda: ())
        monkeypatch.setattr("linexcel.insights._pdftoppm_install_paths", lambda: ())
        with pytest.raises(WorkbookRenderError, match="LibreOffice, pdftoppm"):
            analyze(lineage_excel).save_screenshots(tmp_path)

    def test_renderer_is_found_in_a_standard_install_outside_path(
        self, monkeypatch, tmp_path
    ):
        """Windows and macOS installers do not extend PATH."""
        from linexcel.insights import find_libreoffice, find_pdftoppm

        installed_office = tmp_path / "soffice.com"
        installed_office.write_text("", encoding="utf-8")
        installed_converter = tmp_path / "pdftoppm.exe"
        installed_converter.write_text("", encoding="utf-8")

        monkeypatch.setattr("linexcel.insights.shutil.which", lambda _: None)
        monkeypatch.setattr(
            "linexcel.insights._launcher_install_paths",
            lambda: (tmp_path / "absent.com", installed_office),
        )
        monkeypatch.setattr(
            "linexcel.insights._pdftoppm_install_paths",
            lambda: (installed_converter,),
        )
        assert find_libreoffice() == str(installed_office)
        assert find_pdftoppm() == str(installed_converter)

    def test_screenshots_run_headless_conversion_pipeline(
        self, lineage_excel, monkeypatch, tmp_path
    ):
        calls = []

        def fake_run(command, **_kwargs):
            calls.append(command)
            if command[0] == "libreoffice":
                pdf_dir = Path(command[command.index("--outdir") + 1])
                (pdf_dir / "workbook.pdf").write_bytes(b"%PDF-1.4")
            else:
                Path(f"{command[-1]}-1.png").write_bytes(b"png")

        commands = {"libreoffice": "libreoffice", "pdftoppm": "pdftoppm"}
        monkeypatch.setattr("linexcel.insights.shutil.which", commands.get)
        monkeypatch.setattr("linexcel.insights.subprocess.run", fake_run)
        screenshots = analyze(lineage_excel).save_screenshots(tmp_path)
        assert [path.name for path in screenshots] == ["workbook-1.png"]
        assert calls[0][2:5] == ["--headless", "--norestore", "--convert-to"]
        assert calls[1][1:3] == ["-png", "-r"]

    def test_screenshots_use_a_throwaway_libreoffice_profile(
        self, lineage_excel, monkeypatch, tmp_path
    ):
        """A desktop LibreOffice owns the default profile; sharing it makes the
        headless run exit 0 having converted nothing."""
        calls = []

        def fake_run(command, **_kwargs):
            calls.append(command)
            if command[0] == "libreoffice":
                pdf_dir = Path(command[command.index("--outdir") + 1])
                (pdf_dir / "workbook.pdf").write_bytes(b"%PDF-1.4")
            else:
                Path(f"{command[-1]}-1.png").write_bytes(b"png")

        commands = {"libreoffice": "libreoffice", "pdftoppm": "pdftoppm"}
        monkeypatch.setattr("linexcel.insights.shutil.which", commands.get)
        monkeypatch.setattr("linexcel.insights.subprocess.run", fake_run)
        analyze(lineage_excel).save_screenshots(tmp_path)
        profile = calls[0][1]
        assert profile.startswith("-env:UserInstallation=file://")
        assert "linexcel-render-" in profile

    def test_to_html_is_offline_and_self_contained(self, lineage_excel):
        result = analyze(lineage_excel)
        html = result.to_html()
        assert html.startswith("<!doctype html>")
        # Cytoscape embedded, no network dependency
        assert "cdn.jsdelivr" not in html
        assert "cytoscape" in html
        # the composite formula and its decomposition are in the injected data
        assert "Synthese!B3" in html

    def test_workbook_doc_has_a_separate_html_tab(self, lineage_excel):
        result = analyze(lineage_excel)
        html = result.to_html(workbook_doc="# Workbook role\n\nA test overview.")
        assert "Workbook overview" in html
        assert "workbookDoc" in html
        assert "A test overview." in html

    def test_build_workbook_dossier(self, lineage_excel):
        from linexcel.aidoc import build_workbook_dossier

        dossier = build_workbook_dossier(analyze(lineage_excel).graph)
        sheets = {sheet["name"]: sheet for sheet in dossier["sheets"]}
        assert sheets["Ventes"]["formula_cells"] == 100
        assert sheets["Ventes"]["dimensions"]["columns"] == 7
        assert dossier["defined_names"] == [
            {"name": "TauxCible", "targets": ["Params!A1"]}
        ]
        assert dossier["formula_patterns"][0]["cells"] == 100

    def test_screenshot_tab_is_translated_like_the_rest_of_the_viewer(
        self, lineage_excel
    ):
        """The screenshot pane used to hard-code its French heading and the
        'Page N' labels, so an English report was partly French."""
        png = "data:image/png;base64,aGVsbG8="
        html = analyze(lineage_excel).to_html(language="en", screenshots=[png])
        assert "Aperçu" not in html
        assert '"page": "Page {n}"' in html
        assert "_t('page'" in html

    def test_every_language_defines_the_screenshot_page_label(self):
        from linexcel.i18n import LANGUAGES, UI_STRINGS

        for language in LANGUAGES:
            assert "{n}" in UI_STRINGS[language]["page"]

    def test_repr_html_wraps_in_data_iframe(self, lineage_excel):
        result = analyze(lineage_excel)
        frame = result._repr_html_()
        assert frame.startswith('<iframe src="data:text/html;base64,')

    def test_save_html(self, tmp_path, lineage_excel):
        result = analyze(lineage_excel)
        out = result.save_html(tmp_path / "graph.html")
        assert out.exists() and out.stat().st_size > 100_000

    def test_version_is_exposed(self):
        import linexcel

        assert isinstance(linexcel.__version__, str) and linexcel.__version__

    def test_unsupported_language_is_rejected(self, lineage_excel):
        result = analyze(lineage_excel)
        with pytest.raises(ValueError, match="Unsupported language"):
            result.to_html(language='"; alert(1); var x="')

    def test_to_html_without_source_bytes_drops_the_sheet_tab(self, lineage_excel):
        """A result built without the workbook bytes still renders."""
        payload = analyze_workbook(lineage_excel, "test.xlsx")
        result = LineageResult(graph=payload["graph"], engine=payload["engine"])
        html = result.to_html()
        assert html.startswith("<!doctype html>")
        assert '"workbookContext": null' in html or '"workbookContext":null' in html

    def test_title_does_not_overwrite_workbook_data(self, lineage_excel):
        """Placeholders must not be substituted inside the injected graph."""
        from linexcel.viewer import render_html

        graph = {
            "nodes": [{"id": "n1", "kind": "cell", "label": "__TITLE__"}],
            "edges": [],
            "sheets": [],
            "meta": {"stats": {}, "warnings": []},
        }
        html = render_html(graph, title="Report", language="fr")
        assert '"label": "__TITLE__"' in html
        assert "<h1>Report</h1>" in html

    def test_title_cannot_break_the_embedded_json(self, lineage_excel):
        """A backslash-terminated title used to truncate the GRAPH literal."""
        import json
        import re

        result = analyze(lineage_excel)
        html = result.to_html(title="report\\")
        embedded = re.search(r"var GRAPH = (.*?);\n", html, re.S)
        assert embedded
        assert json.loads(embedded.group(1))["meta"]["stats"]["totalFormulas"] == 103

    def test_document_without_provider_raises_aidocerror(
        self, lineage_excel, monkeypatch
    ):
        """Bare calls no longer fall back to any vendor: no implicit default."""
        from linexcel.aidoc import AiDocError

        for var in (
            "GOOGLE_API_KEY",
            "GEMINI_API_KEY",
            "GEMINI_MODEL",
            "LINEXCEL_AI_BASE_URL",
            "OPENAI_BASE_URL",
        ):
            monkeypatch.delenv(var, raising=False)
        result = analyze(lineage_excel)
        try:
            result.document()
        except AiDocError as exc:
            assert "No AI provider selected" in str(exc)
        else:  # pragma: no cover
            raise AssertionError("AiDocError expected when no provider is configured")


class TestLanguages:
    """The language set is a closed allowlist, and every registry must cover it.

    A language present in one registry but not another would surface as raw
    interface keys in the report or as a KeyError when prompting.
    """

    def test_registries_cover_every_language(self):
        from linexcel.aidoc import _SYSTEM, _WORKBOOK_SYSTEM
        from linexcel.i18n import LANGUAGES, UI_STRINGS

        assert set(UI_STRINGS) == set(LANGUAGES)
        assert set(_SYSTEM) == set(LANGUAGES)
        assert set(_WORKBOOK_SYSTEM) == set(LANGUAGES)

    def test_every_language_defines_every_ui_key(self):
        from linexcel.i18n import DEFAULT_LANGUAGE, UI_STRINGS

        expected = set(UI_STRINGS[DEFAULT_LANGUAGE])
        for language, strings in UI_STRINGS.items():
            assert set(strings) == expected, f"{language} key set differs"

    def test_ui_keys_match_what_the_viewer_asks_for(self):
        """Guards against a key used by the template but defined nowhere."""
        import re as _re
        from pathlib import Path as _Path

        from linexcel.i18n import DEFAULT_LANGUAGE, UI_STRINGS

        source = _Path(viewer_module.__file__).read_text(encoding="utf-8")
        used = set(_re.findall(r"_t\('([a-z_]+)'", source))
        used |= set(_re.findall(r"labelKey: '([a-z_]+)'", source))
        assert used, "no i18n key found in the viewer template"
        assert used <= set(UI_STRINGS[DEFAULT_LANGUAGE]), (
            f"keys used but not defined: {used - set(UI_STRINGS[DEFAULT_LANGUAGE])}"
        )

    def test_placeholders_are_preserved_in_every_translation(self):
        import re as _re

        from linexcel.i18n import DEFAULT_LANGUAGE, UI_STRINGS

        reference = UI_STRINGS[DEFAULT_LANGUAGE]
        for language, strings in UI_STRINGS.items():
            for key, text in strings.items():
                assert set(_re.findall(r"\{(\w+)\}", text)) == set(
                    _re.findall(r"\{(\w+)\}", reference[key])
                ), f"{language}/{key} lost or invented a placeholder"

    @pytest.mark.parametrize("language", ["es", "de", "it", "pt", "nl", "ja", "zh"])
    def test_each_language_renders(self, lineage_excel, language):
        html = analyze(lineage_excel).to_html(language=language)
        assert f"<html lang='{language}'>" in html
        assert f'"{language}"' in html

    def test_only_the_requested_locale_and_english_are_embedded(self, lineage_excel):
        """Shipping all nine locales in every report would be dead weight."""
        from linexcel.i18n import UI_STRINGS

        html = analyze(lineage_excel).to_html(language="ja")
        assert UI_STRINGS["ja"]["sheets_tab"] in html
        assert UI_STRINGS["en"]["sheets_tab"] in html
        assert UI_STRINGS["zh"]["placeholder_title"] not in html

    def test_unsupported_language_is_rejected_by_both_entry_points(self, lineage_excel):
        result = analyze(lineage_excel)
        with pytest.raises(ValueError, match="Unsupported language"):
            result.to_html(language="klingon")
        with pytest.raises(ValueError, match="Unsupported language"):
            result.document_workbook(
                provider=lambda s, u, *, temperature=0.2: "x", language="klingon"
            )


class TestAiProviders:
    """Provider resolution and failure handling — no network call involved."""

    @staticmethod
    def _calc_ids(result: LineageResult) -> list[str]:
        return [n["id"] for n in result.nodes if n["kind"] in ("cell", "group")]

    def test_no_provider_selected_raises_with_guidance(
        self, lineage_excel, monkeypatch
    ):
        """No implicit vendor: a bare call is an explicit error, not Google."""
        from linexcel.aidoc import AiDocError

        for var in (
            "GOOGLE_API_KEY",
            "GEMINI_API_KEY",
            "GEMINI_MODEL",
            "LINEXCEL_AI_BASE_URL",
            "OPENAI_BASE_URL",
        ):
            monkeypatch.delenv(var, raising=False)
        result = analyze(lineage_excel)
        with pytest.raises(AiDocError, match="No AI provider selected"):
            result.document_workbook()

    def test_gemini_requires_an_explicit_model(self, lineage_excel, monkeypatch):
        """A Google key alone no longer selects Gemini — a model must be named."""
        from linexcel.aidoc import AiDocError

        monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
        monkeypatch.delenv("GEMINI_API_KEY", raising=False)
        monkeypatch.delenv("LINEXCEL_AI_BASE_URL", raising=False)
        result = analyze(lineage_excel)
        with pytest.raises(AiDocError) as exc:
            result.document(model="gemini-3.1-flash-lite")
        # model= routes to Gemini, which fails on the missing key or package
        # (dev env) — but never on the neutral "no provider" gate.
        assert "No AI provider selected" not in str(exc.value)

    def test_model_routes_to_gemini_only_when_requested(
        self, lineage_excel, monkeypatch
    ):
        """model= is the explicit opt-in for Gemini; it reaches the client."""
        from linexcel import aidoc

        captured = {}

        class StubGemini:
            def __init__(self, *, api_key, model):
                captured["api_key"] = api_key
                captured["model"] = model

            def generate(
                self, system_prompt, user_prompt, *, temperature=0.2, max_tokens=None
            ):
                return "# gemini"

        monkeypatch.setattr(aidoc, "_GeminiProvider", StubGemini)
        monkeypatch.setenv("GOOGLE_API_KEY", "test-key")
        result = analyze(lineage_excel)
        assert result.document_workbook(model="gemini-3.1-flash-lite") == "# gemini"
        assert captured["model"] == "gemini-3.1-flash-lite"
        assert captured["api_key"] is None  # key read from env by the client

    def test_gemini_model_env_also_opts_in(self, lineage_excel, monkeypatch):
        """GEMINI_MODEL is the env equivalent of model= for Gemini."""
        from linexcel import aidoc

        captured = {}

        class StubGemini:
            def __init__(self, *, api_key, model):
                captured["model"] = model

            def generate(
                self, system_prompt, user_prompt, *, temperature=0.2, max_tokens=None
            ):
                return "# gemini"

        monkeypatch.setattr(aidoc, "_GeminiProvider", StubGemini)
        monkeypatch.setenv("GOOGLE_API_KEY", "test-key")
        monkeypatch.setenv("GEMINI_MODEL", "gemini-2.5-flash")
        result = analyze(lineage_excel)
        assert result.document([self._calc_ids(result)[0]]) == {
            self._calc_ids(result)[0]: "# gemini"
        }
        assert captured["model"] == "gemini-2.5-flash"

    def test_plain_callable_is_accepted(self, lineage_excel):
        seen = []

        def my_llm(system_prompt, user_prompt, *, temperature=0.2):
            seen.append((system_prompt, user_prompt, temperature))
            return "# card"

        result = analyze(lineage_excel)
        node_id = self._calc_ids(result)[0]
        assert result.document([node_id], provider=my_llm) == {node_id: "# card"}
        assert seen[0][2] == 0.2
        assert node_id in seen[0][1]

    def test_object_exposing_generate_is_accepted(self, lineage_excel):
        class Provider:
            def generate(self, system_prompt, user_prompt, *, temperature=0.2, **kw):
                return "# from object"

        result = analyze(lineage_excel)
        node_id = self._calc_ids(result)[0]
        assert result.document([node_id], provider=Provider()) == {
            node_id: "# from object"
        }

    def test_workbook_overview_accepts_a_callable(self, lineage_excel):
        result = analyze(lineage_excel)
        assert (
            result.document_workbook(
                provider=lambda system, user, *, temperature=0.2, **kw: "# overview"
            )
            == "# overview"
        )

    def test_provider_without_generate_is_rejected(self, lineage_excel):
        from linexcel.aidoc import AiDocError

        result = analyze(lineage_excel)
        with pytest.raises(AiDocError, match="must expose a generate"):
            # cast: the point of the test is the runtime rejection
            result.document_workbook(provider=cast(Any, object()))

    def test_partial_failure_keeps_the_successful_cards(self, lineage_excel):
        result = analyze(lineage_excel)
        ids = self._calc_ids(result)
        failing = ids[0]

        def flaky(system_prompt, user_prompt, *, temperature=0.2):
            if f'"node_id": "{failing}"' in user_prompt:
                raise RuntimeError("rate limited")
            return "# card"

        with pytest.warns(UserWarning, match=f"failed for 1 of {len(ids)} nodes"):
            docs = result.document(ids, provider=flaky, max_workers=1)
        assert set(docs) == set(ids) - {failing}

    def test_total_failure_raises(self, lineage_excel):
        from linexcel.aidoc import AiDocError

        result = analyze(lineage_excel)
        ids = self._calc_ids(result)[:2]

        def broken(system_prompt, user_prompt, *, temperature=0.2):
            raise RuntimeError("no backend")

        with pytest.raises(AiDocError, match="failed for all 2 nodes"):
            result.document(ids, provider=broken, max_workers=1)


class TestTokenUsage:
    """Token accounting: real counts when the provider reports them, else
    an estimate that is flagged as such."""

    @staticmethod
    def _calc_ids(result: LineageResult) -> list[str]:
        return [n["id"] for n in result.nodes if n["kind"] in ("cell", "group")]

    def test_estimate_counts_latin_words(self):
        from linexcel.aidoc import estimate_tokens

        assert estimate_tokens("") == 0
        assert estimate_tokens("the quick brown fox jumps") == 6  # 5 words * 4/3

    def test_estimate_counts_cjk_per_character(self):
        """A spaceless Japanese sentence is one `\\w+` match, not one token."""
        import re

        from linexcel.aidoc import estimate_tokens

        text = "日本語のテキストです"
        assert len(re.findall(r"\w+", text)) == 1  # what a naive counter sees
        assert estimate_tokens(text) == len(text) == 10

    def test_mixed_script_counts_both_parts(self):
        from linexcel.aidoc import estimate_tokens

        assert estimate_tokens("SUM 合計 total") == 2 + (2 * 4 // 3)

    def test_callable_provider_usage_is_estimated_and_flagged(self, lineage_excel):
        result = analyze(lineage_excel)
        node_id = self._calc_ids(result)[0]
        result.document([node_id], provider=lambda s, u, *, temperature=0.2: "# card")

        usage = result.token_usage
        assert usage.requests == 1
        assert usage.estimated is True
        assert usage.input_tokens > 0 and usage.output_tokens > 0
        assert usage.total == usage.input_tokens + usage.output_tokens

    def test_provider_reported_counts_are_used_verbatim(self, lineage_excel):
        from linexcel.aidoc import TokenUsage

        class Reporting:
            def generate(self, system_prompt, user_prompt, *, temperature=0.2, **kw):
                return "# card"

            def generate_with_usage(
                self, system_prompt, user_prompt, *, temperature=0.2, **kw
            ):
                return "# card", TokenUsage(
                    input_tokens=1234,
                    output_tokens=56,
                    requests=1,
                    model="m",
                    provider="p",
                )

        result = analyze(lineage_excel)
        result.document([self._calc_ids(result)[0]], provider=Reporting())

        usage = result.token_usage
        assert (usage.input_tokens, usage.output_tokens) == (1234, 56)
        assert usage.estimated is False
        assert usage.model == "m" and usage.provider == "p"

    def test_usage_accumulates_across_calls(self, lineage_excel):
        result = analyze(lineage_excel)
        ids = self._calc_ids(result)
        provider = lambda s, u, *, temperature=0.2: "# card"  # noqa: E731

        result.document(ids, provider=provider, max_workers=1)
        after_nodes = result.token_usage.requests
        result.document_workbook(provider=provider)

        assert after_nodes == len(ids)
        assert result.token_usage.requests == len(ids) + 1

    def test_tokens_spent_before_a_failure_are_still_counted(self, lineage_excel):
        """Tokens already spent are billed even if a later node fails."""
        result = analyze(lineage_excel)
        ids = self._calc_ids(result)
        failing = ids[0]

        def flaky(system_prompt, user_prompt, *, temperature=0.2):
            if f'"node_id": "{failing}"' in user_prompt:
                raise RuntimeError("rate limited")
            return "# card"

        with pytest.warns(UserWarning):
            result.document(ids, provider=flaky, max_workers=1)
        assert result.token_usage.requests == len(ids) - 1

    def test_usage_starts_empty_and_reads_cleanly(self, lineage_excel):
        usage = analyze(lineage_excel).token_usage
        assert (usage.total, usage.requests) == (0, 0)
        assert "0 tokens" in str(usage)

    def test_gemini_usage_metadata_is_read(self):
        from linexcel.aidoc import _usage_from

        class Meta:
            prompt_token_count = 900
            candidates_token_count = 100

        usage = _usage_from(
            Meta(),
            ("prompt_token_count", "candidates_token_count"),
            "prompt",
            "text",
            model="gemini",
            provider="gemini",
        )
        assert (usage.input_tokens, usage.output_tokens, usage.total) == (
            900,
            100,
            1000,
        )
        assert usage.estimated is False

    def test_missing_usage_block_falls_back_to_estimation(self):
        """Local OpenAI-compatible runtimes do not always report usage."""
        from linexcel.aidoc import _usage_from

        usage = _usage_from(
            None,
            ("prompt_tokens", "completion_tokens"),
            "one two three four",
            "five six",
            model="local",
            provider="openai-compatible",
        )
        assert usage.estimated is True
        assert usage.input_tokens > 0 and usage.output_tokens > 0


class TestVba:
    MODULES = {
        "Module1": (
            """Public Sub MAJ()
        total = WorksheetFunction.Sum(Worksheets("Ventes").Range("D2:D101"))
        Worksheets("Synthese").Range("B10").Value = total * Taux()
        Cells(3, 2) = "ok"
    End Sub
    Private Function Taux() As Double
        Taux = Sheets("Params").Range("A1").Value
    End Function
"""
        )
    }

    def test_procedures_and_calls(self):
        procs = analyze_vba(self.MODULES)
        names = {p.name: p for p in procs}
        assert set(names) == {"MAJ", "Taux"}
        assert names["MAJ"].calls == ["Taux"]
        assert names["MAJ"].kind == "Sub"
        assert names["Taux"].kind == "Function"

    def test_read_write_detection(self):
        procs = analyze_vba(self.MODULES)
        maj = next(p for p in procs if p.name == "MAJ")
        accesses = {(r.sheet, r.ref): r.access for r in maj.refs}
        assert accesses[("Ventes", "D2:D101")] == "read"
        assert accesses[("Synthese", "B10")] == "write"
        assert accesses[(None, "B3")] == "write"

    def test_comments_ignored(self):
        procs = analyze_vba(
            {"M": 'Sub S()\n    \' Range("Z9") = 1 in comment\nEnd Sub\n'}
        )
        assert procs[0].refs == []

    def test_calls_are_case_insensitive(self):
        """VBA is case-insensitive: `helper` must reach `Helper`."""
        procs = analyze_vba(
            {
                "M": (
                    "Sub A()\n    Call helper\nEnd Sub\n"
                    "Sub Helper()\n    x = 1\nEnd Sub\n"
                )
            }
        )
        assert next(p for p in procs if p.name == "A").calls == ["Helper"]


class TestVbaGraph:
    """VBA branch of the analyzer.

    openpyxl cannot author a ``vbaProject.bin``, so extraction is stubbed and
    the graph-building code below it runs for real.
    """

    @staticmethod
    def _graph(workbook: bytes, monkeypatch, modules: dict[str, str]) -> dict:
        monkeypatch.setattr(
            "linexcel.analyzer.extract_vba_modules",
            lambda data, filename: dict(modules),
        )
        return analyze_workbook(workbook, "macro.xlsm")["graph"]

    def test_call_edge_links_the_two_procedures(self, lineage_excel, monkeypatch):
        graph = self._graph(
            lineage_excel,
            monkeypatch,
            {
                "Module1": (
                    "Public Sub Refresh()\n"
                    '    Worksheets("Synthese").Range("B10").Value = Rate()\n'
                    "End Sub\n"
                    "Private Function Rate() As Double\n"
                    '    Rate = Sheets("Params").Range("A1").Value\n'
                    "End Function\n"
                )
            },
        )
        assert {n["id"] for n in graph["nodes"] if n["kind"] == "vba"} == {
            "vp:Module1.Refresh",
            "vp:Module1.Rate",
        }
        calls = [
            (e["source"], e["target"]) for e in graph["edges"] if e["kind"] == "call"
        ]
        assert calls == [("vp:Module1.Refresh", "vp:Module1.Rate")]

    def test_read_and_write_edges_reach_the_sheets(self, lineage_excel, monkeypatch):
        graph = self._graph(
            lineage_excel,
            monkeypatch,
            {
                "Module1": (
                    "Sub Refresh()\n"
                    '    Worksheets("Synthese").Range("B10").Value = 1\n'
                    '    x = Sheets("Params").Range("A1").Value\n'
                    "End Sub\n"
                )
            },
        )
        by_id = {n["id"]: n for n in graph["nodes"]}
        writes = [e for e in graph["edges"] if e["kind"] == "vba-write"]
        reads = [e for e in graph["edges"] if e["kind"] == "vba-read"]
        assert [by_id[e["target"]]["label"] for e in writes] == ["Synthese!B10"]
        assert [by_id[e["source"]]["label"] for e in reads] == ["Params!A1"]

    def test_call_resolves_in_the_calling_module_first(
        self, lineage_excel, monkeypatch
    ):
        graph = self._graph(
            lineage_excel,
            monkeypatch,
            {
                "ModA": (
                    "Sub Run()\n    Helper\nEnd Sub\nSub Helper()\n    x = 1\nEnd Sub\n"
                ),
                "ModB": "Sub Helper()\n    y = 2\nEnd Sub\n",
            },
        )
        calls = {
            (e["source"], e["target"]) for e in graph["edges"] if e["kind"] == "call"
        }
        assert ("vp:ModA.Run", "vp:ModA.Helper") in calls
        assert ("vp:ModA.Run", "vp:ModB.Helper") not in calls

    def test_ambiguous_call_stays_unresolved(self, lineage_excel, monkeypatch):
        """A name owned by two other modules must not pick one arbitrarily.

        ``Solo`` is the positive control: without it an empty edge list would
        also pass if call edges stopped being emitted altogether.
        """
        graph = self._graph(
            lineage_excel,
            monkeypatch,
            {
                "ModA": "Sub Helper()\n    x = 1\nEnd Sub\n",
                "ModB": "Sub Helper()\n    y = 2\nEnd Sub\n",
                "ModC": (
                    "Sub Run()\n    Helper\n    Solo\nEnd Sub\n"
                    "Sub Solo()\n    z = 3\nEnd Sub\n"
                ),
            },
        )
        calls = {
            (e["source"], e["target"]) for e in graph["edges"] if e["kind"] == "call"
        }
        assert calls == {("vp:ModC.Run", "vp:ModC.Solo")}

    def test_call_resolution_ignores_casing(self, lineage_excel, monkeypatch):
        """VBA is case-insensitive, so ModA.Taux owns the local `Taux` call.

        The lowercase twin in ModB must not capture it, and the function's own
        `Taux = 1` return assignment is not a call to anything.
        """
        graph = self._graph(
            lineage_excel,
            monkeypatch,
            {
                "ModA": (
                    "Sub Run()\n    Taux\nEnd Sub\n"
                    "Function Taux()\n    Taux = 1\nEnd Function\n"
                ),
                "ModB": "Function taux()\n    taux = 2\nEnd Function\n",
            },
        )
        calls = sorted(
            (e["source"], e["target"]) for e in graph["edges"] if e["kind"] == "call"
        )
        assert calls == [("vp:ModA.Run", "vp:ModA.Taux")]

    def test_member_access_is_not_a_call(self, lineage_excel, monkeypatch):
        """`.Value` is member access even when a procedure is named Value."""
        graph = self._graph(
            lineage_excel,
            monkeypatch,
            {
                "Helpers": (
                    "Public Function Value() As Double\n    Value = 0\nEnd Function\n"
                ),
                "Main": (
                    "Sub Refresh()\n"
                    '    total = Worksheets("S").Range("A1").Value\n'
                    "End Sub\n"
                ),
            },
        )
        assert [e for e in graph["edges"] if e["kind"] == "call"] == []

    def test_module_qualified_call_resolves_exactly(self, lineage_excel, monkeypatch):
        graph = self._graph(
            lineage_excel,
            monkeypatch,
            {
                "ModA": "Sub Run()\n    ModB.Helper\nEnd Sub\n",
                "ModB": "Sub Helper()\n    x = 1\nEnd Sub\n",
            },
        )
        calls = [
            (e["source"], e["target"]) for e in graph["edges"] if e["kind"] == "call"
        ]
        assert calls == [("vp:ModA.Run", "vp:ModB.Helper")]
