"""Deterministic quotation evidence is narrower than factual validation."""

import pytest

from linexcel.ai_validation import MAX_QUOTES, validate_documentation
from linexcel.aidoc import _qualify_documentation, document_nodes, document_workbook
from linexcel.i18n import AI_VALIDATION_NOTICES, LANGUAGES

DOSSIER = {
    "node_id": "c:Summary!C5",
    "formula": "=SUM(Sales!E4:E103)",
    "dependents": [
        {"id": "c:Summary!C7", "formula": '=IF(C5>0,"OK: "&ROUND(C5/1000,1),"No")'}
    ],
}


def test_real_neighbor_transcription_defect_is_qualified_not_rewritten():
    raw = "Aval : `ROUND(C5/100,1)` et `ROUND(C5/1000,1)`."
    report = validate_documentation(raw, DOSSIER)
    assert report["status"] == "qualified"
    assert report["raw_response"] == raw
    assert [c["status"] for c in report["checks"]] == [
        "not_source_supported",
        "source_supported",
    ]
    assert report["checks"][1]["source_ids"] == ["c:Summary!C7"]


@pytest.mark.parametrize(
    "quote",
    ['=IF(C5>0,"ok: "&ROUND(C5/1000,1),"No")', '=IF(C5>0,"OK:"&ROUND(C5/1000,1),"No")'],
)
def test_literal_case_and_whitespace_are_not_normalized(quote):
    report = validate_documentation(f"`{quote}`", DOSSIER)
    assert report["checks"][0]["status"] == "not_source_supported"


def test_cosmetic_spaces_and_function_case_have_source_support():
    report = validate_documentation("`round(C5 / 1000, 1)`", DOSSIER)
    assert report["checks"][0]["status"] == "source_supported"
    assert report["status"] == "unverified"


@pytest.mark.parametrize(
    "source,quote",
    [("=A1 B1", "=A1B1"), ("=A1 'S'!B1", "=A1'S'!B1"), ('="A B"', '="AB"')],
)
def test_intersections_and_string_contents_are_not_erased(source, quote):
    assert (
        validate_documentation(f"`{quote}`", {"formula": source})["checks"][0]["status"]
        != "source_supported"
    )


@pytest.mark.parametrize(
    "quote", ["IF(C5>0,…)", "IF(C5>0,...)", "ROUND(C5/1000,1", '=IF(A1,"open)']
)
def test_abbreviated_or_incomplete_expression_is_distinct(quote):
    assert (
        validate_documentation(f"`{quote}`", DOSSIER)["checks"][0]["status"]
        == "not_checkable"
    )


def test_literal_ellipsis_is_not_an_abbreviated_expression():
    report = validate_documentation('`="..."`', {"formula": '="..."'})
    assert report["checks"][0]["status"] == "source_supported"


def test_math_paraphrases_and_bare_references_do_not_pretend_to_be_checked():
    raw = (
        "`6724.5/1000=6.7245`; `Summary!C5`; linear topology; table "
        "absent; eight members have gaps."
    )
    report = validate_documentation(raw, DOSSIER)
    assert report["status"] == "unverified"
    assert report["checks"] == []
    assert report["limitations"]


def test_hypothetical_complete_formula_is_only_qualified():
    report = validate_documentation("For example, suppose `=SUM(A1:A2)`.", DOSSIER)
    assert report["checks"][0]["status"] == "not_source_supported"
    assert (
        report["status"] == "qualified"
    )  # Never an assertion that the example is false.


def test_does_not_treat_values_comments_or_decomposition_as_source_formulas():
    report = validate_documentation(
        "`=SUM(A1:A2)`",
        {
            "comments": [{"formula": "=SUM(A1:A2)"}],
            "decomposition": {"formula": "=SUM(A1:A2)"},
            "displayed_value": "=SUM(A1:A2)",
        },
    )
    assert report["checks"][0]["status"] == "not_source_supported"


def test_match_does_not_claim_correct_reference_attribution():
    report = validate_documentation("WrongSheet!Z9 is `=SUM(Sales!E4:E103)`.", DOSSIER)
    assert report["checks"][0]["status"] == "source_supported"
    assert report["status"] == "unverified"
    assert "attribution" in " ".join(report["limitations"])


def test_fenced_formulas_and_whole_response_markdown_wrapper():
    report = validate_documentation("```excel\n=SUM(Sales!E4:E103)\n```", DOSSIER)
    assert report["checks"][0]["status"] == "source_supported"
    raw = "```markdown\nBad: `ROUND(C5/100,1)`\n```"
    report = validate_documentation(raw, DOSSIER)
    assert report["status"] == "qualified"
    assert raw[report["checks"][0]["offset"] :].startswith("ROUND")


def test_limits_are_explicit():
    report = validate_documentation("`=1` " * (MAX_QUOTES + 1), {"formula": "=1"})
    assert len(report["checks"]) == MAX_QUOTES
    assert {"kind": "quotation_limit"} in report["issues"]


def test_r1c1_support_retains_the_evidence_kind():
    report = validate_documentation(
        "`SUM(R[-1]C)`",
        {
            "node_id": "S!B2",
            "formula": "=SUM(B1)",
            "r1c1_form": "SUM(R[-1]C)",
        },
    )
    assert report["checks"][0]["evidence"] == [
        {"source_id": "S!B2", "field": "r1c1_form"}
    ]


def test_argument_expression_keeps_strings_and_binding_attribution_unverified():
    report = validate_documentation(
        "`SUM(A1:A2)>Limit`",
        {
            "formula": '=IF(SUM(A1:A2)>Limit,"ROUND(B1/100,1)",0)',
        },
    )
    assert report["checks"][0]["status"] == "source_supported"
    report = validate_documentation(
        "`ROUND(B1/100,1)`",
        {
            "formula": '=IF(SUM(A1:A2)>Limit,"ROUND(B1/100,1)",0)',
        },
    )
    assert report["checks"][0]["status"] == "not_source_supported"


def test_illustrations_and_ellipses_are_counted_without_alerts():
    report = validate_documentation(
        "`ROUND(6.7245,1)`; `ROUND(…)`; `=SUM(A1:A2)\n=12\n=12`; `formualizer (Rust)`",
        DOSSIER,
    )
    assert report["status"] == "unverified"
    assert report["counts"]["not_checkable"] == 3
    assert report["issues"] == []


@pytest.mark.parametrize(
    "source,quote",
    [
        ("=SUM(RC[-1] RC[-2])", "=SUM(RC[-1]RC[-2])"),
        ("=SUM(Table1[A] Table1[B])", "=SUM(Table1[A]Table1[B])"),
        ("=SUM(Table1[A] [B])", "=SUM(Table1[A][B])"),
    ],
)
def test_bracket_reference_intersections_remain_semantic(source, quote):
    assert (
        validate_documentation(f"`{quote}`", {"formula": source})["status"]
        == "qualified"
    )
    assert (
        validate_documentation(f"`{source}`", {"formula": source})["checks"][0][
            "status"
        ]
        == "source_supported"
    )


@pytest.mark.parametrize(
    "source,quote",
    [
        ("=SUM(Table1[ A ])", "=SUM(Table1[A])"),
        ("=SUM(Table1[Unit Price])", "=SUM(Table1[Unit  Price])"),
        ("=SUM(Table1[IF(A1,2,3)])", "IF(A1,2,3)"),
        ("=Table1[ROUND(A1)]", "ROUND(A1)"),
    ],
)
def test_structured_column_headers_are_not_formula_syntax(source, quote):
    report = validate_documentation(f"`{quote}`", {"formula": source})
    assert report["checks"][0]["status"] != "source_supported"
    assert (
        validate_documentation(f"`{source}`", {"formula": source})["checks"][0][
            "status"
        ]
        == "source_supported"
    )


def test_notice_quotes_only_three_bounded_excerpts_without_rewriting_response():
    raw = "Raw <text> is preserved."
    quote = '=IF(A1,"`</blockquote><img src=x>`",B1)'
    report = {
        "status": "qualified",
        "checks": [
            {"status": "not_source_supported", "quotation": quote},
            {"status": "not_source_supported", "quotation": "=" + "A1+" * 200},
            {"status": "not_source_supported", "quotation": "=A3"},
            {"status": "not_source_supported", "quotation": "=A4"},
        ],
    }
    rendered = _qualify_documentation(raw, report, "en")
    assert "Quotation 1: excerpt omitted (Markdown delimiters)." in rendered
    assert quote not in rendered
    assert "=A4" not in rendered
    assert "(+1)" in rendered
    assert "A1+" * 100 not in rendered
    assert rendered.endswith("\n\n" + raw)


@pytest.mark.parametrize("language", LANGUAGES)
def test_api_qualification_visible_by_default_with_optional_raw_reports(language):
    graph = {
        "nodes": [{"id": "a", "kind": "cell", "formula": "=SUM(A1:A2)"}],
        "edges": [],
    }
    raw = "Source: `=SUM(A1:A3)`"
    reports = {}
    docs = document_nodes(
        graph,
        ["a"],
        provider=lambda *a, **kw: raw,
        language=language,
        validation_results=reports,
    )
    assert docs["a"].startswith("> " + AI_VALIDATION_NOTICES[language])
    assert docs["a"].endswith("\n\n" + raw)
    assert "> ` =SUM(A1:A3) `" in docs["a"]
    assert reports["a"]["raw_response"] == raw
    overview = document_workbook(
        graph,
        provider=lambda *a, **kw: raw,
        language=language,
        validation_results=reports,
    )
    assert overview.startswith("> " + AI_VALIDATION_NOTICES[language])
    assert reports["workbook"]["status"] == "qualified"
    assert document_nodes(graph, ["a"], provider=lambda *a, **kw: raw)["a"].startswith(
        "> "
    )
