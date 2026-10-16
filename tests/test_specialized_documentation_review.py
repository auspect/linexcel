"""A real second source-led pass, not a label on the initial generation."""

import hashlib
import json

import pytest

from linexcel.ai_validation import validate_documentation
from linexcel.aidoc import AiDocError, TokenUsage, document_nodes, document_workbook
from linexcel.doc_evidence import formula_evidence
from linexcel.documentation_proof import render_formula_tokens

GRAPH = {
    "nodes": [
        {
            "id": "c:S!C7",
            "sheet": "S",
            "addr": "C7",
            "kind": "cell",
            "formula": "=C5>5000",
            "sourceFormula": "=C5>5000",
            "value": True,
            "valueSource": "engine",
        }
    ],
    "edges": [],
}


class ReviewProvider:
    def __init__(self, responses):
        self.responses = iter(responses)
        self.requests = []

    def generate(self, *args, **kwargs):
        raise AssertionError("usage method expected")

    def generate_with_usage(self, system, user, **kwargs):
        self.requests.append((system, user))
        response = next(self.responses)
        if isinstance(response, Exception):
            raise response
        return response, TokenUsage(input_tokens=10, output_tokens=5, requests=1)


@pytest.mark.parametrize("workbook", [False, True])
def test_default_runs_draft_then_review_and_preserves_both_sources_and_cost(workbook):
    provider = ReviewProvider(
        [
            "Incorrect draft: saved cache is engine.",
            "Revised: source origin is engine; saved cache is separate.",
        ]
    )
    reports, usage = {}, TokenUsage()
    if workbook:
        result = document_workbook(
            GRAPH, provider=provider, validation_results=reports, usage=usage
        )
        report = reports["workbook"]
    else:
        result = document_nodes(
            GRAPH,
            ["c:S!C7"],
            provider=provider,
            validation_results=reports,
            usage=usage,
        )["c:S!C7"]
        report = reports["c:S!C7"]
    assert len(provider.requests) == 2
    assert "SPECIALIZED REVIEW PASS" not in provider.requests[0][0]
    assert "SPECIALIZED REVIEW PASS" in provider.requests[1][0]
    assert "Incorrect draft" in provider.requests[1][1]
    assert report["sent_dossier"]["json_text"] in provider.requests[1][1]
    assert usage.requests == 2 and usage.total == 30
    stages = report["specialized_review"]["stages"]
    assert [stage["phase"] for stage in stages] == ["draft", "review"]
    assert report["raw_response"] == stages[-1]["raw_response"]
    for stage, (system, user) in zip(stages, provider.requests):
        assert (
            stage["system_prompt_sha256"] == hashlib.sha256(system.encode()).hexdigest()
        )
        assert stage["user_prompt_sha256"] == hashlib.sha256(user.encode()).hexdigest()
        assert (
            stage["content_sha256"]
            == hashlib.sha256(stage["raw_response"].encode()).hexdigest()
        )
    assert "2 requests" in result and "not factually certified" in result
    assert result.endswith("Revised: source origin is engine; saved cache is separate.")
    assert not report["specialized_review"]["factual_certification"]


def test_opt_out_is_explicit_and_preserved_as_disabled_not_reviewed():
    provider = ReviewProvider(["One draft."])
    reports = {}
    result = document_workbook(
        GRAPH, provider=provider, validation_results=reports, specialized_review=False
    )
    assert len(provider.requests) == 1
    assert reports["workbook"]["specialized_review"]["status"] == "disabled"
    assert "2 requests" not in result


@pytest.mark.parametrize(
    "failure",
    [
        "",
        "```markdown\n\n```",
        AiDocError("truncated", usage=TokenUsage(input_tokens=7, requests=1)),
    ],
)
def test_empty_or_failed_review_never_falls_back_to_reviewed_draft(failure):
    provider = ReviewProvider(["Draft is not the final review.", failure])
    reports, usage = {}, TokenUsage()
    with pytest.raises(AiDocError):
        document_workbook(
            GRAPH, provider=provider, validation_results=reports, usage=usage
        )
    assert usage.requests == 2
    assert reports["workbook"]["status"] == "failed"
    assert reports["workbook"]["specialized_review"]["stages"][-1]["status"] == "failed"
    assert (
        reports["workbook"]["specialized_review"]["stages"][0]["raw_response"]
        == "Draft is not the final review."
    )


def test_budget_is_checked_between_actual_draft_and_review_requests():
    provider = ReviewProvider(["Draft.", "Must never be sent."])
    reports, usage = {}, TokenUsage()
    with pytest.raises(AiDocError, match="before specialized review"):
        document_workbook(
            GRAPH,
            provider=provider,
            validation_results=reports,
            usage=usage,
            token_budget=12,
        )
    assert len(provider.requests) == 1 and usage.total == 15
    assert reports["workbook"]["specialized_review"]["stages"][-1]["status"] == "failed"


def test_complete_final_request_can_overshoot_api_budget_with_exact_cost():
    provider = ReviewProvider(["Draft.", "Complete final above approximate budget."])
    reports, usage = {}, TokenUsage()
    result = document_workbook(
        GRAPH,
        provider=provider,
        validation_results=reports,
        usage=usage,
        token_budget=25,
    )
    assert usage.total == 30 and usage.requests == 2
    assert len(provider.requests) == 2
    assert result.endswith("Complete final above approximate budget.")
    assert reports["workbook"]["specialized_review"]["status"] == "completed"
    assert (
        reports["workbook"]["raw_response"]
        == "Complete final above approximate budget."
    )


def test_complete_final_overshoot_stops_the_next_queued_node():
    provider = ReviewProvider(["Draft.", "Complete reviewed first node."])
    second = {**GRAPH["nodes"][0], "id": "c:S!C8", "addr": "C8"}
    graph = {"nodes": [*GRAPH["nodes"], second], "edges": []}
    reports, usage = {}, TokenUsage()
    with pytest.warns(UserWarning, match="1 never sent"):
        result = document_nodes(
            graph,
            ["c:S!C7", "c:S!C8"],
            provider=provider,
            validation_results=reports,
            usage=usage,
            max_workers=1,
            token_budget=25,
        )
    assert usage.total == 30 and usage.requests == 2
    assert list(result) == ["c:S!C7"]
    assert reports["c:S!C7"]["specialized_review"]["status"] == "completed"
    assert "c:S!C8" not in reports


def test_space_intersection_union_and_literal_self_reference_are_separate_facts():
    intersection = formula_evidence("=SUM('Sales Data'!D2:D10 'Sales Data'!D5:D20)")
    assert intersection["reference_operators"][0]["meaning"] == "intersection"
    union = formula_evidence("=SUM((A1:A9,A5:A20))")
    assert union["reference_operators"][0]["meaning"] == "union"
    assert (
        formula_evidence("=C5+1", sheet="S", address="C7")["literal_self_reference"]
        is False
    )
    assert (
        formula_evidence("='S'!$C$7+1", sheet="S", address="C7")[
            "literal_self_reference"
        ]
        is True
    )
    assert (
        formula_evidence("=Other!C7+1", sheet="S", address="C7")[
            "literal_self_reference"
        ]
        is False
    )
    assert formula_evidence("=C7+1")["literal_self_reference"] is None


def test_overview_proof_selects_exact_unique_retained_id_not_representative_guess():
    dossier = {
        "formula_patterns": [
            {
                "node_id": "g:Sales Data!F2#250",
                "source_formula": "=D2*E2",
                "formula_provenance": {"xmlText": "D2*E2"},
            },
            {
                "node_id": "c:S!C7",
                "source_formula": "=C5>5000",
                "formula_provenance": {"xmlText": "C5>5000"},
            },
        ]
    }
    text, insertions = render_formula_tokens(
        "{{SOURCE_FORMULA:c:S!C7}}\n{{STORED_XML_FORMULA:g:Sales Data!F2#250}}", dossier
    )
    assert "=C5>5000" in text and "D2*E2" in text
    assert insertions[0]["pointer"] == "/formula_patterns/1/source_formula"
    assert insertions[1]["pointer"] == "/formula_patterns/0/formula_provenance/xmlText"
    dossier["formula_patterns"].append(dossier["formula_patterns"][1])
    _, bad = render_formula_tokens("{{SOURCE_FORMULA:c:S!C7}}", dossier)
    assert bad[0]["status"] == "not_supplied"


@pytest.mark.parametrize(
    "left,result,expected",
    [
        (6724.5, True, "supported"),
        (6724.5, False, "contradicted"),
        (672.45, True, "contradicted"),
    ],
)
def test_explicit_arithmetic_claim_checks_bound_operands_and_result(
    left, result, expected
):
    claim = {
        "statement": "Comparison claim.",
        "classification": "fact",
        "evidence": [{"pointer": "/total", "value": 6724.5}],
        "calculation": {
            "operator": ">",
            "operands": [
                {"pointer": "/total", "value": left},
                {"pointer": "/limit", "value": 5000},
            ],
            "result": result,
        },
    }
    raw = "Comparison claim.\n```json_claims\n" + json.dumps([claim]) + "\n```"
    report = validate_documentation(raw, {"total": 6724.5, "limit": 5000})
    assert report["factual_review"]["checks"][0]["calculation"]["verdict"] == expected
    assert report["raw_response"] == raw
    if expected == "contradicted":
        assert report["status"] == "qualified"


@pytest.mark.parametrize("value", ["1,234", True, None, [1]])
def test_scalar_arithmetic_never_parses_strings_or_coerces_other_json_types(value):
    from linexcel.ai_validation import _calculation_check

    check = _calculation_check(
        {
            "operator": "+",
            "operands": [
                {"pointer": "/a", "value": value},
                {"pointer": "/b", "value": 1},
            ],
            "result": 2,
        },
        {"a": value, "b": 1},
    )
    assert check["verdict"] == "indeterminate"


def test_scalar_division_is_bounded_and_independent_of_ambient_decimal_context():
    from decimal import localcontext

    from linexcel.ai_validation import _calculation_check

    def calculation(left, right, result):
        return _calculation_check(
            {
                "operator": "/",
                "operands": [
                    {"pointer": "/a", "value": left},
                    {"pointer": "/b", "value": right},
                ],
                "result": result,
            },
            {"a": left, "b": right},
        )

    with localcontext() as context:
        context.prec = 1
        assert calculation(6724.5, 1000, 6.7245)["verdict"] == "supported"
    assert calculation(1, 3, 0.3333)["reason"] == "nonterminating_decimal_division"
    assert calculation(1, 0, 0)["verdict"] == "indeterminate"


def test_group_neighbor_identity_does_not_exclude_a_nonrepresentative_member():
    from linexcel.aidoc import build_dossier

    group = {
        "id": "g:Hostile Text!C2#8",
        "kind": "group",
        "sheet": "Hostile Text",
        "addr": "C2",
        "bbox": "C2:C9",
        "count": 8,
        "sourceFormula": '=B2&" copied"',
    }
    focal = {
        "id": "c:Hostile Text!B9",
        "kind": "cell",
        "sheet": "Hostile Text",
        "addr": "B9",
        "formula": "=broken",
        "valueSource": "file",
        "value": "#VALUE!",
    }
    dossier = build_dossier(
        {
            "nodes": [focal, group],
            "edges": [{"source": focal["id"], "target": group["id"]}],
        },
        focal["id"],
    )
    neighbor = dossier["dependents"][0]
    assert (neighbor["address"], neighbor["extent"]) == ("C2", "C2:C9")
    assert neighbor["edge_source"] == focal["id"]
    assert neighbor["edge_target"] == group["id"]
    assert (
        neighbor["value_role"] == "representative cell, not an aggregate of the group"
    )
    assert not dossier["result_evidence"]["recorded_engine_formula_result"]
    assert dossier["computed_value"] is None


@pytest.mark.parametrize("patterns", [None, {"omitted": "size limit"}, [None]])
def test_omitted_or_invalid_overview_examples_cannot_supply_a_proof(patterns):
    text, records = render_formula_tokens(
        "{{SOURCE_FORMULA:c:S!A1}}", {"formula_patterns": patterns}
    )
    assert "not_supplied" in text and records[0]["status"] == "not_supplied"
