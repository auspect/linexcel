"""Replay observed provenance/offset/topology mistakes against sent evidence."""

import copy
import json

import pytest

from linexcel.ai_validation import (
    MAX_CLAIMS,
    REVIEW_VERSION,
    review_protocol,
    validate_documentation,
)
from linexcel.aidoc import (
    _SYSTEM,
    _VISION_SYSTEM,
    _WORKBOOK_SYSTEM,
    MAX_WORKBOOK_DOSSIER_CHARS,
    _fit_node_dossier,
    _fit_workbook_dossier,
    build_dossier,
    build_workbook_dossier,
    document_nodes,
)
from linexcel.doc_evidence import formula_evidence
from linexcel.i18n import LANGUAGES
from linexcel.web_ai import evidence


def response(statement, pointer, value, classification="fact"):
    return (
        statement
        + "\n```json_claims\n"
        + json.dumps(
            [
                {
                    "statement": statement,
                    "classification": classification,
                    "evidence": [{"pointer": pointer, "value": value}],
                }
            ]
        )
        + "\n```"
    )


@pytest.mark.parametrize(
    "pointer,claimed,actual",
    [
        ("/value_origin_kind", "external_file_read", "name_resolution"),
        ("/value_source", "engine", "volatile"),
        ("/cached_value", "OK:6.7k", "OK:6,7k"),
        ("/lookup_offsets/0/offset_from_first", 5, 4),
        ("/graph_connections/branching_nodes/0", "C5", "E4"),
        ("/source_formula", "=1+2", "=" + "(" * 40 + "1+2" + ")" * 40),
    ],
)
def test_observed_mutations_are_contradicted_without_rewriting(
    pointer, claimed, actual
):
    dossier = {}
    # These cases use the same field names/typed distinctions as the review pack.
    if pointer.startswith("/lookup_offsets"):
        dossier["lookup_offsets"] = [{"offset_from_first": actual}]
    elif pointer.startswith("/graph_connections"):
        dossier["graph_connections"] = {"branching_nodes": [actual]}
    else:
        dossier[pointer[1:]] = actual
    raw = response("An explicit claim for review.", pointer, claimed)
    report = validate_documentation(raw, dossier)
    assert report["status"] == "qualified"
    assert report["raw_response"] == raw
    assert report["factual_review"]["counts"]["contradicted"] == 1
    assert (
        report["factual_review"]["checks"][0]["bindings"][0]["evidence_value"] == actual
    )


@pytest.mark.parametrize(
    "actual,claimed", [(False, 0), (1, "1"), (None, "null"), ([True], [1])]
)
def test_types_are_part_of_evidence(actual, claimed):
    review = validate_documentation(
        response("Typed claim.", "/v", claimed), {"v": actual}
    )
    assert review["factual_review"]["counts"]["contradicted"] == 1


@pytest.mark.parametrize(
    "pointer",
    ["/missing", "/omitted/value", "/review_protocol/version", "/x~2", "/list/01"],
)
def test_missing_omitted_or_protocol_evidence_cannot_support_facts(pointer):
    dossier = {
        "omitted": {"omitted": "size limit", "value": 42},
        "review_protocol": {"version": 42},
        "list": [42],
    }
    report = validate_documentation(response("Claim.", pointer, 42), dossier)
    assert report["status"] == "qualified"
    assert report["factual_review"]["counts"]["indeterminate"] == 1


def test_supported_binding_does_not_certify_false_prose_or_inferences():
    # Deliberately false prose: checking a value's presence is not semantic review.
    raw = response("The engine recalculated this saved cache.", "/value_source", "file")
    report = validate_documentation(raw, {"value_source": "file"})
    assert report["status"] == "unverified"
    assert report["factual_review"]["counts"]["supported"] == 1
    assert "prose" in " ".join(report["factual_review"]["limitations"])
    report = validate_documentation(
        response("It may be stale.", "/v", 2, "inference"), {"v": 2}
    )
    assert report["factual_review"]["counts"]["indeterminate"] == 1
    assert (
        validate_documentation("No ledger.", {})["factual_review"]["contract"]
        == "missing"
    )


def test_required_review_contract_cannot_silently_pass_without_a_ledger():
    dossier = {"review_protocol": review_protocol(), "v": 2}
    for raw in ["Plain prose.", "Plain prose.\n```json\n[]\n```"]:
        report = validate_documentation(raw, dossier)
        assert report["status"] == "qualified"
        assert {"kind": "missing_required_claim_block"} in report["issues"]
        assert report["raw_response"] == raw
    assert validate_documentation("Legacy prose.", {})["status"] == "unverified"


def test_required_review_rejects_empty_or_multiple_ledgers_without_certifying_prose():
    dossier = {"review_protocol": review_protocol(), "v": 2}
    empty = validate_documentation("Prose.\n```json_claims\n[]\n```", dossier)
    assert {"kind": "empty_required_claim_ledger"} in empty["issues"]
    raw = response("Claim.", "/v", 2)
    report = validate_documentation(raw + "\n" + raw, dossier)
    assert {"kind": "multiple_required_claim_blocks"} in report["issues"]
    assert report["status"] == "qualified"
    valid = validate_documentation(raw, dossier)
    assert valid["status"] == "unverified"
    assert valid["factual_review"]["counts"]["supported"] == 1


def test_report_formula_roles_do_not_claim_presentation_is_stored_xml():
    roles = validate_documentation("Prose.", {})["formula_field_roles"]
    assert roles["source_formula"] == "presented_original_or_derived_shared_translation"
    assert roles["sourceFormula"] == roles["source_formula"]
    assert (
        roles["formula_provenance.xmlText"] == "decoded_stored_xml_without_added_equals"
    )


def test_claims_must_quote_visible_prose_and_have_bounded_valid_json():
    raw = response("Actual claim.", "/a~1b/~0x/0", 4)
    assert (
        validate_documentation(raw, {"a/b": {"~x": [4]}})["factual_review"]["counts"][
            "supported"
        ]
        == 1
    )
    changed = raw.replace("Actual claim.", "Absent claim.", 1)
    assert (
        validate_documentation(changed, {})["factual_review"]["checks"][0]["reason"]
        == "statement_not_in_response"
    )
    assert (
        validate_documentation("```json_claims\n{bad}\n```", {})["status"]
        == "qualified"
    )
    payload = [{"statement": "Claim.", "classification": "unknown", "evidence": []}] * (
        MAX_CLAIMS + 1
    )
    report = validate_documentation(
        "Claim.\n```json_claims\n" + json.dumps(payload) + "\n```", {}
    )
    assert len(report["factual_review"]["checks"]) == MAX_CLAIMS
    assert {"kind": "claim_count_limit"} in report["issues"]


@pytest.mark.parametrize("classification", [[], {}, 2, None])
def test_untrusted_claim_types_are_qualified_instead_of_crashing(classification):
    raw = response("Claim.", "/v", 1, classification)
    report = validate_documentation(raw, {"v": 1})
    assert report["status"] == "qualified"
    assert report["factual_review"]["checks"][0]["reason"] == "invalid_claim_schema"


def test_incomplete_claim_block_cannot_leak_review_scaffolding_into_prose():
    from linexcel.aidoc import _insert_tables

    raw = 'Actual prose.\n```json_claims\n[{"statement":'
    report = validate_documentation(raw, {})
    assert {"kind": "incomplete_claim_block"} in report["issues"]
    assert report["raw_response"] == raw
    assert _insert_tables(raw) == "Actual prose."


@pytest.mark.parametrize(
    "label", ["json_claims.", "json_claims_extra", "json_claims []"]
)
def test_claim_fence_mentions_and_partial_labels_do_not_create_a_ledger(label):
    from linexcel.ai_validation import strip_claims

    raw = "Actual prose.\nMention ```" + label + '\n```json\n{"v": 1}\n```'
    report = validate_documentation(raw, {"review_protocol": {"claims_required": True}})
    assert {"kind": "missing_required_claim_block"} in report["issues"]
    assert {"kind": "incomplete_claim_block"} not in report["issues"]
    assert strip_claims(raw) == raw


def test_inline_table_fence_mention_cannot_hide_actual_prose_evidence():
    raw = (
        'Mention a block ```json_tables.\nVisible proof.\n```json\n{"other": 1}\n```\n'
        + response("Visible proof.", "/v", 1).split("\n", 1)[1]
    )
    report = validate_documentation(raw, {"v": 1})
    assert report["factual_review"]["checks"][0]["verdict"] == "supported"


def test_claim_fences_accept_crlf_and_markdown_indentation():
    raw = (
        response("Visible proof.", "/v", 1)
        .replace("```json_claims", "  ```json_claims ")
        .replace("\n", "\r\n")
    )
    report = validate_documentation(raw, {"v": 1})
    assert report["factual_review"]["checks"][0]["verdict"] == "supported"


def test_private_table_json_is_not_visible_statement_evidence():
    raw = (
        'Actual prose.\n```json_tables\n[{"caption":"Hidden claim."}]\n```\n'
        + response("Hidden claim.", "/v", 2).split("\n", 1)[1]
    )
    report = validate_documentation(raw, {"v": 2})
    assert report["status"] == "qualified"
    assert (
        report["factual_review"]["checks"][0]["reason"] == "statement_not_in_response"
    )


def test_generation_retains_ledger_raw_but_renders_only_prose_and_qualification():
    raw = response("This cache is from the engine.", "/value_source", "engine")
    reports = {}
    docs = document_nodes(
        {"nodes": [{"id": "x", "valueSource": "file"}], "edges": []},
        ["x"],
        provider=lambda *a, **kw: raw,
        validation_results=reports,
    )
    assert "json_claims" not in docs["x"]
    assert docs["x"].endswith("This cache is from the engine.")
    assert "AI evidence review" in docs["x"]
    assert reports["x"]["raw_response"] == raw


def test_literal_lookup_offset_and_function_inventory_are_not_support_verdicts():
    facts = formula_evidence(
        "=IFERROR(VLOOKUP(A1,B2:F20,5,FALSE),XLOOKUP(A1,B2:B20,F2:F20))"
    )
    assert facts["functions_observed"] == ["IFERROR", "VLOOKUP", "XLOOKUP"]
    assert facts["function_selectors"] == []
    assert facts["lookup_offsets"][0]["offset_from_first"] == 4
    assert facts["lookup_offsets"][0]["table_reference"] == "B2:F20"
    assert (
        formula_evidence("=VLOOKUP(A1,B2:F20,ColIndex,FALSE)")["lookup_offsets"][0][
            "offset_from_first"
        ]
        is None
    )
    assert "not engine support" in facts["function_inventory_scope"]
    assert (
        formula_evidence("=HLOOKUP(A1,B2:F2001,2000,FALSE)")["lookup_offsets"][0][
            "offset_from_first"
        ]
        == 1999
    )


def test_original_canonical_and_shared_translation_survive_batch_and_web_dossiers():
    node = {
        "id": "S!B22",
        "kind": "cell",
        "sheet": "S",
        "addr": "B22",
        "cell": "B22",
        "formula": "=1 + 2",
        "sourceFormula": "=" + "(" * 40 + "1+2" + ")" * 40,
        "canonicalFormula": "=1 + 2",
        "valueSource": "engine",
        "value": 3,
        "formulaProvenance": {"status": "stored_xml", "canonicalMatchesSource": False},
        "sourceValueProvenance": {
            "role": "formula_cache",
            "rawXmlValue": "3",
            "status": "present",
        },
    }
    graph = {"nodes": [node], "edges": [], "meta": {}}
    dossier = json.loads(_fit_node_dossier(build_dossier(graph, node["id"])))
    assert dossier["source_formula"] == node["sourceFormula"]
    assert dossier["canonical_formula"] == "=1 + 2"
    assert dossier["formula_provenance"]["canonicalMatchesSource"] is False
    assert (
        evidence(graph, node["id"], [])["node"]["sourceFormula"]
        == node["sourceFormula"]
    )
    node.update(sourceFormula=None, formulaProvenance={"status": "shared_translation"})
    assert build_dossier(graph, node["id"])["source_formula"] is None


def test_exact_sample_formulas_support_quotations_without_inventing_from_values():
    dossier = {
        "source_formula": "=C4*D4",
        "value_samples": [
            {"addr": "E29", "sourceFormula": "=C29*D29", "canonicalFormula": None},
            {"addr": "E54", "sourceFormula": "=C54*D54", "value": 75.0},
            {
                "addr": "E78",
                "sourceFormula": {"omitted": "dossier size limit"},
                "value": "=C78*D78",
                "cachedValue": "=C78*D78",
            },
        ],
    }
    report = validate_documentation("Stored `=C29*D29` and `=C54*D54`.", dossier)
    assert report["status"] == "unverified"
    assert report["counts"]["source_supported"] == 2
    assert report["checks"][0]["evidence"] == [
        {"source_id": "E29", "field": "sourceFormula"}
    ]
    report = validate_documentation("Invented `=C78*D78` or `=C103*D103`.", dossier)
    assert report["status"] == "qualified"
    assert report["counts"]["not_source_supported"] == 2


@pytest.mark.parametrize("samples", [None, {"omitted": "size limit"}, False, [None]])
def test_absent_or_omitted_samples_cannot_support_formula_quotes(samples):
    report = validate_documentation("Invented `=C29*D29`.", {"value_samples": samples})
    assert report["status"] == "qualified"
    assert report["counts"]["not_source_supported"] == 1


def test_overview_branches_and_merges_identify_different_nodes():
    graph = {
        "nodes": [],
        "edges": [{"source": "E4", "target": name} for name in ("C5", "C6", "C7")]
        + [{"source": name, "target": "E4"} for name in ("C4", "D4")],
        "meta": {},
    }
    connections = build_workbook_dossier(graph)["graph_connections"]
    assert connections["branching_nodes"] == ["E4"]
    assert connections["merging_nodes"] == ["E4"]
    assert evidence(graph, None, [])["graph_connections"] == {
        **connections,
        "semantics": (
            "Source dependency/static call relationships, never execution order."
        ),
    }


def test_static_vba_source_is_evidence_not_execution():
    node = {
        "id": "v:Module.Main",
        "kind": "vba",
        "module": "Module",
        "proc": "Main",
        "procKind": "Sub",
        "code": 'Sub Main()\nRange("A1").Value = 2\nEnd Sub' * 100,
        "vbaProvenance": {"execution": "not_executed", "extractor": "oletools.olevba"},
        "lines": [1, 3],
    }
    graph = {"nodes": [node], "edges": [], "meta": {}}
    batch = build_dossier(graph, node["id"])["vba"]
    web = evidence(graph, node["id"], [])["node"]
    assert batch["execution"] == web["execution"] == "not_executed"
    assert batch["code"] == web["code"] == node["code"][:2500]
    assert batch["code_omitted_characters"] == web["codeOmittedCharacters"] > 0


def test_specialized_review_contract_is_available_in_all_localized_prompts():
    assert (
        set(_SYSTEM) == set(_WORKBOOK_SYSTEM) == set(_VISION_SYSTEM) == set(LANGUAGES)
    )
    for language in LANGUAGES:
        for prompt in (_SYSTEM[language], _WORKBOOK_SYSTEM[language]):
            assert REVIEW_VERSION in prompt
            assert "json_claims" in prompt
            assert "source_formula/sourceFormula" in prompt
            assert "lookup_offsets" in prompt
            assert "formula_provenance.xmlText" in prompt
            assert "cached_agreement=differ" in prompt
            assert "preview_range" in prompt
        assert "hidden state" in _VISION_SYSTEM[language]
        assert "LOCAL crop pixels" in _VISION_SYSTEM[language]
        assert "Return ONLY one JSON object" in _VISION_SYSTEM[language]


def test_large_overview_examples_fit_without_erasing_retained_source():
    # Provenance repeats the original XML text beside a canonical expression.
    # Keeping five examples used to exhaust the bound even after dropping
    # presentation and names; reducing examples must keep the retained proof.
    source = "=" + "(" * 1500 + "1+2" + ")" * 1500
    graph = {
        "nodes": [
            {
                "id": f"S!B{i}",
                "kind": "cell",
                "sheet": "S",
                "addr": f"B{i}",
                "formula": "=1 + 2",
                "canonicalFormula": "=1 + 2",
                "sourceFormula": source,
                "formulaProvenance": {
                    "status": "stored_xml",
                    "xmlText": source[1:],
                    "canonicalMatchesSource": False,
                },
            }
            for i in range(10)
        ],
        "edges": [{"source": "S!B0", "target": "S!B1"}],
        "meta": {
            "definedNameEvidence": {
                "status": "complete",
                "definitions": [
                    {"name": "Threshold", "expression": "5000", "scope": "workbook"}
                ],
            }
        },
    }
    context = {
        "sheets": [
            {
                "name": "S",
                "preview": [
                    {"row": i, "values": ["Label " + "x" * 50]} for i in range(9)
                ],
            }
        ]
    }
    original_graph = copy.deepcopy(graph)
    original = build_workbook_dossier(graph, context=context)
    blob = _fit_workbook_dossier(original)
    fitted = json.loads(blob)
    assert len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS
    assert len(fitted["formula_patterns"]) == 1
    retained = fitted["formula_patterns"][0]
    assert retained["source_formula"] == source
    assert retained["canonical_formula"] == "=1 + 2"
    assert retained["formula_provenance"]["xmlText"] == source[1:]
    assert fitted["formula_pattern_coverage"] == {
        "scope": "graph_formula_nodes",
        "total": 10,
        "shown": 1,
        "omitted": 9,
    }
    assert fitted["source_defined_names"]["definitions"][0]["name"] == "Threshold"
    assert len(fitted["sheets"][0]["preview"]) == 4
    assert fitted["sheets"][0]["preview_rows_omitted_from_dossier"] == 5
    assert fitted["graph_connections"]["edges"] == graph["edges"]
    assert graph == original_graph
    report = validate_documentation(
        response("An omitted formula.", "/formula_patterns/1/source_formula", source),
        fitted,
    )
    assert report["factual_review"]["counts"]["indeterminate"] == 1


def test_overview_size_fallback_counts_omissions_without_substitute_evidence():
    graph = {
        "nodes": [
            {"id": "S!A1", "kind": "cell", "sourceFormula": "=" + "1+" * 10000 + "1"},
            *[
                {"id": f"v:{i}", "kind": "vba", "module": "M", "proc": f"P{i}"}
                for i in range(25)
            ],
        ],
        "edges": [],
        "meta": {
            "definedNameEvidence": {
                "status": "complete",
                "definitions": [
                    {"name": f"N{i}", "expression": "x" * 1000, "scope": "workbook"}
                    for i in range(30)
                ],
            }
        },
    }
    dossier = build_workbook_dossier(graph)
    dossier["sheets"] = [
        {
            "name": "S",
            "preview": [{"row": i, "values": [i]} for i in range(9)],
            "comments": [{"cell": "A1", "text": "x" * 500} for _ in range(30)],
        }
    ]
    blob = _fit_workbook_dossier(dossier)
    fitted = json.loads(blob)
    assert len(blob) <= MAX_WORKBOOK_DOSSIER_CHARS
    assert fitted["formula_patterns"] == []
    assert fitted["formula_pattern_coverage"]["omitted"] == 1
    assert fitted["sheets"][0]["preview_rows_omitted_from_dossier"] == 9
    assert fitted["sheets"][0]["comments_omitted_from_dossier"] == 30
    assert "preview" not in fitted["sheets"][0]
    assert (
        len(fitted["vba_procedures"]) + fitted["vba_procedures_omitted_from_dossier"]
        == 25
    )
    names = fitted["source_defined_names"]
    assert len(names["definitions"]) + names["omitted_for_size"] == 30
    assert names["scope_counts_in_supplied_definitions"]["workbook"] == len(
        names["definitions"]
    )
    assert fitted["review_protocol"]["version"] == REVIEW_VERSION
