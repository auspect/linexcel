"""Frozen lazy dossiers share formula facts and exact proofs with batch text."""

import json

import pytest
from test_web_ai import CONFIG, FakeProvider, graph

from linexcel import aidoc
from linexcel.web_ai import evidence, generate


@pytest.mark.parametrize("formula,self_reference", [("=B1+1", True), ("=A1*7", False)])
def test_lazy_focal_and_overview_preserve_known_formula_target(formula, self_reference):
    source = graph()
    source["nodes"][1]["formula"] = formula
    focal = evidence(source, "S!B1", [])
    assert focal["formula_facts"]["literal_self_reference"] is self_reference
    overview = evidence(source, None, [])
    pattern = overview["formula_patterns"][0]
    assert pattern["node_id"] == "S!B1"
    assert pattern["formula_facts"]["literal_self_reference"] is self_reference
    assert "formula_facts" not in focal["node"]  # Same facts are not duplicated.


def test_lazy_overview_qualified_tokens_insert_the_selected_retained_source(
    monkeypatch,
):
    source = graph()
    source["nodes"][1].update(
        sourceFormula="=(A1*7)",
        canonicalFormula="=A1 * 7",
        formulaProvenance={"xmlText": "(A1*7)"},
    )
    dossier = evidence(source, None, [])
    assert "sample" not in dossier
    assert len(dossier["formula_patterns"]) == len(dossier["value_samples"]) == 1
    pattern = dossier["formula_patterns"][0]
    assert "sourceFormula" not in pattern and "id" not in pattern
    assert pattern["source_formula"] == "=(A1*7)"
    raw = (
        "{{SOURCE_FORMULA:S!B1}}\n{{STORED_XML_FORMULA:S!B1}}\n"
        "{{CANONICAL_FORMULA:S!B1}}"
    )
    provider = FakeProvider(raw)
    monkeypatch.setattr(aidoc, "_resolve_provider", lambda **kwargs: provider)
    output = generate(dossier, CONFIG, {"language": "en"}, 16000)
    insertions = output["validation"]["deterministic_formula_insertions"]
    assert [record["value"] for record in insertions] == [
        "=(A1*7)",
        "(A1*7)",
        "=A1 * 7",
    ]
    assert insertions[0]["pointer"] == "/formula_patterns/0/source_formula"
    assert insertions[1]["pointer"] == "/formula_patterns/0/formula_provenance/xmlText"
    assert output["rawResponse"] == raw
    assert "```text\n=(A1*7)\n```" in output["markdown"]
    assert len(provider.calls) == 2
    for _, user, _ in provider.calls:
        assert json.dumps(dossier, ensure_ascii=False) in user


@pytest.mark.parametrize(
    "value,task_status,result_status,revision,target,formula,present,expected",
    [
        (0, "succeeded", "completed", "abc", "S!B1", "=A1*7", True, True),
        (False, "succeeded", "completed", "abc", "S!B1", "=A1*7", True, True),
        ("#REF!", "succeeded", "completed", "abc", "S!B1", "=A1*7", True, True),
        (None, "succeeded", "completed", "abc", "S!B1", "=A1*7", True, False),
        (42, "succeeded", "completed", "abc", "S!B1", "=A1*7", False, False),
        (42, "failed", "completed", "abc", "S!B1", "=A1*7", True, False),
        (42, "succeeded", "unsupported", "abc", "S!B1", "=A1*7", True, False),
        (42, "succeeded", "completed", "stale", "S!B1", "=A1*7", True, False),
        (42, "succeeded", "completed", "abc", "S!A1", "=A1*7", True, False),
        (42, "succeeded", "completed", "abc", "S!B1", "", True, False),
    ],
)
def test_lazy_engine_result_evidence_requires_the_actual_completed_formula_target(
    value, task_status, result_status, revision, target, formula, present, expected
):
    source = graph()
    source["nodes"][1]["formula"] = formula
    result = {
        "nodeId": target,
        "status": result_status,
        "provenance": "targeted_recalculation",
    }
    if present:
        result["value"] = value
    task = {
        "id": "calculation-1",
        "nodeId": target,
        "status": task_status,
        "revision": revision,
        "finishedAt": 123,
        "result": result,
    }
    snapshot = evidence(source, "S!B1", [task], revision="abc")
    proof = snapshot["latestCalculation"]["result_evidence"]
    assert proof["recorded_engine_formula_result"] is expected
    assert proof["saved_cache_is_engine_result"] is False
    assert proof["error_cause"] == "not_established"
    assert snapshot["savedCache"]["value"] == 40
