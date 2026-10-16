"""A per-sheet A1 address cannot establish one scalar for a 3D sheet span."""

import copy

import pytest

from linexcel.aidoc import _compact_steps, build_dossier, build_workbook_dossier
from linexcel.doc_evidence import (
    formula_evidence,
    three_dimensional_reference_evidence,
)
from linexcel.web_ai import evidence


@pytest.mark.parametrize(
    "reference,first,last,address,external",
    [
        ("'Sales Data:Lookups'!A1", "Sales Data", "Lookups", "A1", False),
        ("Sheet1:Sheet3!$A$1:$B$2", "Sheet1", "Sheet3", "$A$1:$B$2", False),
        ("'O''Brien:End'!A1", "O'Brien", "End", "A1", False),
        ("'Start:End'!A:A", "Start", "End", "A:A", False),
        ("Sheet1:Sheet1!A1", "Sheet1", "Sheet1", "A1", False),
        ("'[Budget:2026.xlsx]Start:End'!A1", "Start", "End", "A1", True),
        ("'C:\\Book\\[B.xlsx]Start:End'!A1", "Start", "End", "A1", True),
    ],
)
def test_ast_reference_has_exact_3d_scope_without_sheet_count_or_result(
    reference, first, last, address, external
):
    facts = formula_evidence("=SUM(" + reference + ")")
    (span,) = facts["three_dimensional_references"]
    assert span["reference"] == reference
    assert span["first_sheet_token"] == first
    assert span["last_sheet_token"] == last
    assert span["address_per_sheet"] == address
    assert span["spans_sheets"] is True
    assert span["external_workbook_qualifier"] is external
    assert span["sheet_count"] is span["total_cells"] is None
    assert "not proof of one scalar" in span["scope"]
    assert "evaluated values are not established" in span["scope"]
    assert facts["references"] == [reference]


@pytest.mark.parametrize(
    "formula",
    [
        '=IF(A1="Start:End!A1",1,0)',
        '=INDIRECT("Start:End!A1")',
        "=SUM(A1:B3)",
        "=SUM('My Sheet'!A1:B2)",
        "=SUM('[Budget:2026.xlsx]Single'!A1)",
        "=SUM('C:\\Book\\[B.xlsx]Single'!A1)",
    ],
)
def test_text_literals_address_colons_and_external_path_are_not_3d_spans(formula):
    assert "three_dimensional_references" not in formula_evidence(formula)


def test_3d_fact_list_is_bounded_without_losing_omission_count():
    facts = formula_evidence(
        "=SUM(" + ",".join(f"S{i}:E{i}!A1" for i in range(23)) + ")"
    )
    assert len(facts["three_dimensional_references"]) == 20
    assert facts["three_dimensional_references_omitted"] == 3


def test_reported_3d_count_is_not_promoted_to_total_cell_count():
    step = {
        "expr": "SUM(S1:S3!A1)",
        "evaluated": False,
        "inputs": [{"ref": "S1:S3!A1", "value": {"range": "S1:S3!A1", "n": 1}}],
    }
    before = copy.deepcopy(step)
    (item,) = _compact_steps(step)["inputs"]
    assert item["value"]["n"] == 1  # Preserve the supplied observation unchanged.
    assert item["value"]["n_unit"] == "not_determined"
    assert item["value"]["dimensions"] is None
    assert (
        item["value"]["count_scope"]
        == "reported_count_not_established_as_total_3d_cells"
    )
    assert item["value"]["three_dimensional_reference"]["total_cells"] is None
    assert step == before


def test_unavailable_real_3d_input_stays_unavailable():
    step = {
        "expr": "SUM(S1:S3!A1)",
        "evaluated": True,
        "value": "#REF!",
        "inputs": [{"ref": "S1:S3!A1", "value": None}],
    }
    assert _compact_steps(step)["inputs"] == step["inputs"]


def test_source_facts_reach_batch_and_lazy_card_and_overview():
    node = {
        "id": "c:Cross Refs!B2",
        "kind": "cell",
        "sheet": "Cross Refs",
        "addr": "B2",
        "formula": "=SUM('Sales Data:Lookups'!A1)",
        "sourceFormula": "=SUM('Sales Data:Lookups'!A1)",
        "value": 0,
        "cachedValue": 0,
        "valueSource": "file",
    }
    graph = {"nodes": [node], "edges": []}
    before = copy.deepcopy(graph)
    facts = [
        build_dossier(graph, node["id"])["formula_facts"],
        build_workbook_dossier(graph)["formula_patterns"][0]["formula_facts"],
        evidence(graph, node["id"], [])["formula_facts"],
        evidence(graph, None, [])["formula_patterns"][0]["formula_facts"],
    ]
    assert all(
        fact["three_dimensional_references"][0]["reference"]
        == "'Sales Data:Lookups'!A1"
        for fact in facts
    )
    assert graph == before


@pytest.mark.parametrize(
    "invalid", [None, "A1", "A1:B2", "'Only'!A1", "'[Bad]Bad]Start:End'!A1"]
)
def test_3d_helper_does_not_guess_from_invalid_or_single_sheet_token(invalid):
    assert three_dimensional_reference_evidence(invalid) is None
