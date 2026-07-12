"""Tests for the lineage module: references, grouping, graph, VBA, API."""

import io

from linexcel import LineageResult, analyze
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
        assert rect == Rect("S", 2, 2, 2, 2)

    def test_parse_range_with_sheet(self):
        rect = parse_ref("'Ma Feuille'!A1:C10")
        assert rect.sheet == "Ma Feuille"
        assert (rect.r1, rect.c1, rect.r2, rect.c2) == (1, 1, 10, 3)

    def test_parse_whole_column(self):
        rect = parse_ref("A:B")
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
        rect = stretch_ref(detail, 2, 4, (2, 101), (4, 4))
        assert (rect.r1, rect.r2) == (2, 101)
        assert rect.c1 == rect.c2 == 1

    def test_stretch_anchored_stays_fixed(self):
        detail = parse_ref_detailed("$A$2", "S")
        rect = stretch_ref(detail, 2, 4, (2, 101), (4, 4))
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
        input_labels = {
            n["label"] for n in graph["nodes"] if n["kind"] == "input"
        }
        assert "Ventes!B2:B101" in input_labels
        assert "Ventes!C2:C101" in input_labels

    def test_defined_name_resolved(self, lineage_excel):
        graph = analyze_workbook(lineage_excel, "test.xlsx")["graph"]
        names = [n for n in graph["nodes"] if n["kind"] == "name"]
        assert names and names[0]["label"] == "TauxCible"
        # name is fed by Params!A1 and feeds Synthese!B3
        edges = graph["edges"]
        assert any(
            e["target"] == names[0]["id"] and "Params" in e["source"]
            for e in edges
        )
        assert any(
            e["source"] == names[0]["id"]
            and e["target"].endswith("Synthese!B3")
            for e in edges
        )

    def test_composed_formula_steps_evaluated(self, lineage_excel):
        graph = analyze_workbook(lineage_excel, "test.xlsx")["graph"]
        node = next(
            n for n in graph["nodes"] if n["id"].endswith("Synthese!B3")
        )
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
        assert result.node(node_id)["formula"].startswith("=IF")

    def test_to_json_roundtrip(self, lineage_excel):
        import json

        result = analyze(lineage_excel)
        data = json.loads(result.to_json())
        assert data["meta"]["stats"]["totalFormulas"] == 103

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
        assert sheets["Ventes"]["dimensions"]["columns"] == 4
        assert dossier["defined_names"] == [
            {"name": "TauxCible", "targets": ["Params!A1"]}
        ]
        assert dossier["formula_patterns"][0]["cells"] == 100

    def test_repr_html_wraps_in_data_iframe(self, lineage_excel):
        result = analyze(lineage_excel)
        frame = result._repr_html_()
        assert frame.startswith('<iframe src="data:text/html;base64,')

    def test_save_html(self, tmp_path, lineage_excel):
        result = analyze(lineage_excel)
        out = result.save_html(tmp_path / "graph.html")
        assert out.exists() and out.stat().st_size > 100_000

    def test_document_without_key_raises_aidocerror(
        self, lineage_excel, monkeypatch
    ):
        from linexcel.aidoc import AiDocError

        monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
        monkeypatch.delenv("GEMINI_API_KEY", raising=False)
        result = analyze(lineage_excel)
        try:
            result.document()
        except AiDocError:
            pass
        else:  # pragma: no cover
            raise AssertionError("AiDocError expected when no key is provided")


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
