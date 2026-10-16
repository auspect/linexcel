"""Exact source proofs do not depend on model transcription."""

import hashlib
import json

import pytest

from linexcel.aidoc import document_nodes, document_workbook
from linexcel.documentation_proof import render_formula_tokens


@pytest.mark.parametrize("web", [False, True])
def test_native_40_parenthesis_pairs_and_presentation_prefix_remain_exact(web):
    xml = "(" * 40 + "1+2" + ")" * 40
    source = "=" + xml
    node = (
        {
            "sourceFormula": source,
            "canonicalFormula": "=1 + 2",
            "formulaProvenance": {"xmlText": xml},
        }
        if web
        else {
            "source_formula": source,
            "canonical_formula": "=1 + 2",
            "formula_provenance": {"xmlText": xml},
        }
    )
    dossier = {"node": node} if web else node
    text, records = render_formula_tokens(
        "{{SOURCE_FORMULA}}\n{{STORED_XML_FORMULA}}\n{{CANONICAL_FORMULA}}", dossier
    )
    assert "```text\n" + source + "\n```" in text
    assert "```text\n" + xml + "\n```" in text
    assert [r["characters"] for r in records] == [84, 83, 6]
    assert records[1]["value_sha256"] == hashlib.sha256(xml.encode()).hexdigest()
    assert records[1]["pointer"] == (
        "/node/formulaProvenance/xmlText" if web else "/formula_provenance/xmlText"
    )
    assert records[1]["role"] == "decoded_stored_xml_without_added_equals"


def test_missing_original_or_omitted_field_never_falls_back_to_canonical_or_neighbor():
    dossier = {
        "source_formula": {"omitted": "size limit"},
        "canonical_formula": "=A1+1",
        "precedents": [{"source_formula": "=Z1+2"}],
    }
    text, records = render_formula_tokens(
        "{{SOURCE_FORMULA}}\n{{STORED_XML_FORMULA}}", dossier
    )
    assert "=A1+1" not in text and "=Z1+2" not in text
    assert all(r["status"] == "not_supplied" for r in records)


def test_hostile_delimiters_use_labelled_json_string_with_exact_decodable_text():
    source = '=IF(A1,"```\n</details><img onerror=alert(1)>\n`",0)'
    text, records = render_formula_tokens(
        "{{SOURCE_FORMULA}}", {"source_formula": source}
    )
    assert records[0]["representation"] == "json_encoded_string"
    encoded = text.split("```text\n", 1)[1].rsplit("\n```", 1)[0]
    assert json.loads(encoded) == source
    assert records[0]["value"] == source


def test_generation_retains_raw_tokens_and_separate_insertion_ledger():
    source = "=" + "(" * 40 + "1+2" + ")" * 40
    graph = {
        "nodes": [
            {
                "id": "c:S!B22",
                "kind": "cell",
                "formula": "=1 + 2",
                "sourceFormula": source,
                "canonicalFormula": "=1 + 2",
                "formulaProvenance": {"xmlText": source[1:], "status": "stored_xml"},
            }
        ],
        "edges": [],
    }
    raw = "Proof:\n{{SOURCE_FORMULA}}\n{{STORED_XML_FORMULA}}"
    reports = {}
    docs = document_nodes(
        graph, ["c:S!B22"], provider=lambda *a, **kw: raw, validation_results=reports
    )
    assert source in docs["c:S!B22"]
    assert reports["c:S!B22"]["raw_response"] == raw
    assert len(reports["c:S!B22"]["deterministic_formula_insertions"]) == 2


def test_overview_tokens_report_missing_focal_formula_without_using_sample():
    graph = {
        "nodes": [{"id": "c:S!A1", "kind": "cell", "formula": "=1+2"}],
        "edges": [],
    }
    raw = "Proof:\n{{SOURCE_FORMULA}}\n{{STORED_XML_FORMULA}}\n{{CANONICAL_FORMULA}}"
    reports = {}
    text = document_workbook(
        graph, provider=lambda *a, **kw: raw, validation_results=reports
    )
    assert "{{" not in text and "=1+2" not in text
    assert text.count("not_supplied_in_sent_dossier") == 3
    assert reports["workbook"]["raw_response"] == raw
    records = reports["workbook"]["deterministic_formula_insertions"]
    assert [r["status"] for r in records] == ["not_supplied"] * 3


@pytest.mark.parametrize(
    "raw",
    [
        "`{{SOURCE_FORMULA}}`",
        "``{{SOURCE_FORMULA}}``",
        "```text\n{{SOURCE_FORMULA}}\n```",
        "```text\n`{{SOURCE_FORMULA}}`\n{{STORED_XML_FORMULA}}\n```",
        "Before `{{SOURCE_FORMULA}}` after.",
    ],
)
def test_code_wrapped_tokens_insert_standalone_blocks_with_exact_source(raw):
    source = "=" + "(" * 40 + "1+2" + ")" * 40
    text, records = render_formula_tokens(
        raw, {"source_formula": source, "formula_provenance": {"xmlText": source[1:]}}
    )
    assert "{{" not in text and "````" not in text
    assert "```text\n" + source + "\n```\n\n" in text
    assert records[0]["value"] == source
    fences = [line for line in text.splitlines() if line.startswith("```")]
    assert len(fences) == 2 * len(records)
    if raw.startswith("Before"):
        assert text.startswith("Before \n\n") and text.endswith("\n\n after.")
