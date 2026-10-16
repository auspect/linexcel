"""Replay literal-text and full/source-preview boundary mistakes."""

import copy
import hashlib
import json
from xml.etree import ElementTree as ET

import pytest

from linexcel.ai_validation import REVIEW_VERSION
from linexcel.aidoc import (
    _fit_node_dossier,
    build_dossier,
    build_workbook_dossier,
    document_nodes,
    document_workbook,
)
from linexcel.doc_evidence import (
    formula_text_evidence,
    value_comparison_evidence,
    vba_excerpt_evidence,
)
from linexcel.formula_source import cell_source_fields
from linexcel.i18n import LANGUAGES
from linexcel.web_ai import evidence


def test_decoded_xml_and_added_equals_are_measured_without_raw_bytes_claim():
    element = ET.fromstring(
        '<c r="B22"><f>  (((IF(A1=&quot;a&amp;b&quot;,1,2))))  </f><v>1</v></c>'
    )
    fields = cell_source_fields(
        element,
        workbook_sha256="workbook-hash",
        worksheet_part="xl/worksheets/sheet1.xml",
        sheet="S",
        cell="B22",
        canonical_formula='=IF(A1="a&b",1,2)',
        canonical_origin="engine",
    )
    stored = '  (((IF(A1="a&b",1,2))))  '
    assert fields["formulaProvenance"]["xmlText"] == stored
    assert fields["formulaProvenance"]["xmlTextSemantics"] == (
        "decoded_xml_character_data"
    )
    assert fields["formulaProvenance"]["sourceFormulaPrefixOrigin"] == (
        "presentation_added"
    )
    assert fields["sourceFormula"] == "=" + stored
    proof = formula_text_evidence(fields)
    assert proof["stored_xml_character_count"] == len(stored)
    assert proof["presented_source_character_count"] == len(stored) + 1
    assert proof["source_to_xml_relation"] == (
        "presentation_equals_prefix_plus_decoded_xml"
    )
    assert proof["added_prefix"] == "="
    assert proof["canonical_to_presented_source"] == "literal_different"


def test_missing_or_shared_slave_text_is_not_inferred_from_canonical_formula():
    for node in (
        {"formula": "=A2+2"},
        {
            "sourceFormula": None,
            "canonicalFormula": "=A2+2",
            "formulaProvenance": {"status": "shared_translation", "xmlText": ""},
        },
    ):
        proof = formula_text_evidence(node)
        assert proof["source_to_xml_relation"] == "not_comparable"
        assert proof["added_prefix"] is None
        assert proof["presented_source_character_count"] is None
        assert proof["canonical_to_presented_source"] == "not_comparable"
    proof = formula_text_evidence(
        {"sourceFormula": "=A1+3", "formulaProvenance": {"xmlText": "A1+2"}}
    )
    assert proof["source_to_xml_relation"] == "inconsistent"
    assert proof["added_prefix"] is None


@pytest.mark.parametrize("comparison", ["format", "differ", "match", None])
@pytest.mark.parametrize(
    "computed,cached", [("OK:6.7k", "OK:6,7k"), ("1,234", "1.234")]
)
def test_different_strings_never_gain_numeric_equivalence_from_a_label(
    comparison, computed, cached
):
    node = {
        "value": computed,
        "cachedValue": cached,
        "cachedAgreement": comparison,
    }
    before = copy.deepcopy(node)
    proof = value_comparison_evidence(node)
    assert proof["literal_text_equal"] is False
    assert proof["numeric_equivalence"] == "not_established_from_strings"
    assert proof["displayed_json_type"] == proof["cached_json_type"] == "string"
    assert node == before


def test_no_numeric_parse_or_type_coercion_is_claimed_for_missing_or_typed_values():
    proof = value_comparison_evidence({"value": True, "cachedValue": 1})
    assert proof["displayed_json_type"] == "boolean"
    assert proof["cached_json_type"] == "number"
    assert proof["literal_text_equal"] is None
    assert proof["numeric_equivalence"] == "not_assessed"
    proof = value_comparison_evidence({"value": 1234, "cachedValue": "1,234"})
    assert proof["numeric_equivalence"] == "not_established_from_strings"
    proof = value_comparison_evidence({"value": "1.00"})
    assert proof["cached_json_type"] == "unknown"
    assert proof["literal_text_equal"] is None


def test_all_localized_text_prompts_require_actual_boundaries_and_literal_ledger():
    from linexcel.aidoc import _SYSTEM, _WORKBOOK_SYSTEM

    for language in LANGUAGES:
        for prompt in (_SYSTEM[language], _WORKBOOK_SYSTEM[language]):
            assert REVIEW_VERSION in prompt
            assert "value_comparison.literal_text_equal=false" in prompt
            assert "source_to_xml_relation" in prompt
            assert "omitted_full_source_to_graph" in prompt
            assert "omitted_graph_to_dossier" in prompt
            assert "omitted_full_source_to_dossier" in prompt
            assert "copy a short contiguous excerpt" in prompt
            assert "no other delimiters" not in prompt


def macro_node():
    full = "é" * 10351
    return full, {
        "id": "v:Module.Macro",
        "kind": "vba",
        "module": "Module",
        "proc": "Macro",
        "procKind": "Sub",
        "lines": [2, 300],
        "code": full[:6000],
        "vbaProvenance": {
            "totalCodeCharacters": len(full),
            "codeTruncated": True,
            "procedureSha256": hashlib.sha256(full.encode()).hexdigest(),
            "execution": "not_executed",
        },
    }


def test_full_source_graph_and_sent_excerpt_omissions_are_separate_character_counts():
    full, node = macro_node()
    proof = vba_excerpt_evidence(node)
    assert proof["code"] == full[:2500]
    assert proof["code_status"] == "supplied_excerpt"
    assert proof["code_bounds"] == {
        "unit": "decoded_source_characters",
        "full_source_characters": 10351,
        "graph_characters": 6000,
        "dossier_characters": 2500,
        "omitted_full_source_to_graph": 4351,
        "omitted_graph_to_dossier": 3500,
        "omitted_full_source_to_dossier": 7851,
        "source_and_graph_size_consistent": True,
    }
    assert proof["code_omitted_characters"] == 3500
    assert proof["code_hashes"] == {
        "full_source_sha256": hashlib.sha256(full.encode()).hexdigest(),
        "graph_sha256": hashlib.sha256(full[:6000].encode()).hexdigest(),
        "dossier_sha256": hashlib.sha256(full[:2500].encode()).hexdigest(),
    }
    assert len(proof["code"].encode()) == 5000  # Characters are not UTF-8 bytes.


@pytest.mark.parametrize("full_count", [None, -1, True, "10351"])
def test_unknown_full_source_count_does_not_become_graph_length(full_count):
    _, node = macro_node()
    node["vbaProvenance"]["totalCodeCharacters"] = full_count
    proof = vba_excerpt_evidence(node)
    assert proof["code_bounds"]["full_source_characters"] is None
    assert proof["code_bounds"]["omitted_full_source_to_graph"] is None
    assert proof["code_bounds"]["omitted_full_source_to_dossier"] is None
    assert proof["code_bounds"]["source_and_graph_size_consistent"] is None
    assert proof["code_omitted_characters"] == 3500


@pytest.mark.parametrize("full_count,flag", [(5999, True), (10351, False)])
def test_inconsistent_source_counts_never_produce_negative_or_guessed_omissions(
    full_count, flag
):
    _, node = macro_node()
    node["vbaProvenance"].update(totalCodeCharacters=full_count, codeTruncated=flag)
    proof = vba_excerpt_evidence(node)
    assert proof["code_bounds"]["source_and_graph_size_consistent"] is False
    assert proof["code_bounds"]["omitted_full_source_to_graph"] is None
    assert proof["code_bounds"]["omitted_full_source_to_dossier"] is None


def test_size_omission_updates_the_sent_stage_without_claiming_excerpt_hash():
    _, node = macro_node()
    node["vbaProvenance"]["extra_metadata"] = "x" * 12000
    graph = {"nodes": [node], "edges": []}
    before = copy.deepcopy(graph)
    sent = json.loads(_fit_node_dossier(build_dossier(graph, node["id"])))
    macro = sent["vba"]
    assert macro["code"]["omitted"] == "dossier size limit"
    assert macro["code_status"] == "omitted_by_dossier_size_limit"
    assert macro["code_bounds"]["dossier_characters"] == 0
    assert macro["code_bounds"]["omitted_graph_to_dossier"] == 6000
    assert macro["code_bounds"]["omitted_full_source_to_dossier"] == 10351
    assert macro["code_omitted_characters"] == 6000
    assert macro["code_hashes"]["dossier_sha256"] is None
    assert graph == before


def test_batch_and_web_reuse_the_same_bounds_and_distinguish_latest_from_saved():
    _, node = macro_node()
    graph = {"nodes": [node], "edges": []}
    batch = build_dossier(graph, node["id"])["vba"]
    web = evidence(graph, node["id"], [])["node"]
    for field in ("code", "code_status", "code_bounds", "code_hashes"):
        assert batch[field] == web[field]
    assert build_workbook_dossier(graph)["vba_source_scope"]["code_text"] == (
        "not_supplied_in_overview"
    )
    formula = {
        "id": "S!C7",
        "kind": "cell",
        "value": "OK:6,7k",
        "cachedValue": "OK:6,7k",
        "cachedAgreement": "match",
        "formula": "=A1",
        "dependencies": [],
    }
    calculation = {
        "id": "request",
        "result": {"value": "OK:6.7k", "comparison": "format"},
    }
    sent = evidence({"nodes": [formula], "edges": []}, formula["id"], [calculation])
    assert sent["node"]["value_comparison"]["literal_text_equal"] is True
    current = sent["latestCalculation"]["value_comparison"]
    assert current["literal_text_equal"] is False
    assert current["numeric_equivalence"] == "not_established_from_strings"


def test_node_and_overview_reports_preserve_exact_sent_json_not_a_reconstruction():
    received = []

    def provider(system, user, **kwargs):
        received.append(user.split("\n", 1)[1])
        return "A source-grounded statement."

    node = {
        "id": "S!A1",
        "kind": "cell",
        "formula": "=1+2",
        "sourceFormula": "=1+2",
        "formulaProvenance": {"xmlText": "1+2"},
        "value": "é",
        "cachedValue": "e",
    }
    graph = {"nodes": [node], "edges": []}
    reports = {}
    document_nodes(graph, [node["id"]], provider=provider, validation_results=reports)
    document_workbook(graph, provider=provider, validation_results=reports)
    assert len(received) == 4  # Both documents actually receive a second review.
    for key, exact, critic in zip(
        (node["id"], "workbook"), received[::2], received[1::2]
    ):
        snapshot = reports[key]["sent_dossier"]
        assert snapshot["json_text"] == exact
        assert (
            snapshot["json_text_sha256"] == hashlib.sha256(exact.encode()).hexdigest()
        )
        assert snapshot["characters"] == len(exact)
        assert snapshot["utf8_bytes"] == len(exact.encode())
        assert critic.startswith(exact + "\n\nCandidate to review")
        assert reports[key]["raw_response"] == "A source-grounded statement."
        assert (
            reports[key]["factual_review"]["dossier_sha256"]
            == hashlib.sha256(
                json.dumps(
                    json.loads(exact), ensure_ascii=False, sort_keys=True, default=str
                ).encode()
            ).hexdigest()
        )
