"""Named sources narrow the authorized selection before invoking any model."""

import pytest

from linexcel.question_sources import SourceSelectionError, select_question_sources


def rev(project, name):
    return {"projectId": project, "id": f"{project}:sha", "filename": name}


@pytest.fixture
def sources():
    return [rev("a", "Budget FY26.xlsx"), rev("b", "other.xlsx"), rev("c", "Café.xlsm")]


@pytest.mark.parametrize(
    "question",
    [
        'Dans "Budget FY26.xlsx", quelle est la marge ?',
        "Dans BUDGET FY26.XLSX, quelle est la marge ?",
        "Quelle marge dans [Budget FY26.xlsx] ?",
    ],
)
def test_literal_filename_freezes_one_source(sources, question):
    result, selection = select_question_sources(sources, question)
    assert result == sources[:1]
    assert selection == {"kind": "question_filename", "filenames": ["Budget FY26.xlsx"]}


def test_multiple_named_sources_are_a_whitelist(sources):
    result, _ = select_question_sources(sources, "Compare other.xlsx et Café.xlsm")
    assert {item["projectId"] for item in result} == {"b", "c"}


def test_unicode_and_case_are_normalized_but_accents_remain_significant(sources):
    assert select_question_sources(sources, "CAFE\u0301.XLSM")[0] == sources[2:]
    with pytest.raises(SourceSelectionError):
        select_question_sources(sources, "Cafe.xlsm")


def test_conflicting_known_name_never_broadens(sources):
    with pytest.raises(SourceSelectionError):
        select_question_sources(sources, "other.xlsx", project="a")


@pytest.mark.parametrize(
    "question",
    [
        'source: "Missing Budget.xlsx"',
        "old-other.xlsx",
        "Quelles cellules lisent [Prix 2026.xlsx]Tarifs!B2 ?",
        'Qui lit "Missing other.xlsx" ?',
    ],
)
def test_unknown_name_keeps_the_explicit_selection(sources, question):
    # An external link target that is not imported is text, not a new source.
    for selection in ({"project": "a"}, {"revision_id": "a:sha"}):
        result, chosen = select_question_sources(sources, question, **selection)
        assert result == sources[:1]
        assert chosen == {"kind": "explicit", "filenames": ["Budget FY26.xlsx"]}


@pytest.mark.parametrize(
    "question", ['source: "Missing Budget.xlsx"', "old-other.xlsx"]
)
def test_unknown_name_without_selection_still_stops(sources, question):
    with pytest.raises(SourceSelectionError):
        select_question_sources(sources, question)


def test_duplicate_name_needs_project_or_revision():
    sources = [rev("a", "same.xlsx"), rev("b", "SAME.xlsx")]
    with pytest.raises(SourceSelectionError):
        select_question_sources(sources, "Dans same.xlsx, quelle valeur ?")
    assert select_question_sources(sources, "same.xlsx", project="b")[0] == sources[1:]
    assert (
        select_question_sources(sources, "same.xlsx", revision_id="a:sha")[0]
        == sources[:1]
    )


def test_longer_literal_name_does_not_select_embedded_name():
    sources = [rev("a", "the budget.xlsx"), rev("b", "budget.xlsx")]
    assert select_question_sources(sources, "the budget.xlsx")[0] == sources[:1]


def test_no_named_source_preserves_structured_selection(sources):
    assert select_question_sources(sources, "Quelle marge ?")[0] == sources
    result, selection = select_question_sources(sources, "Quelle marge ?", project="a")
    assert result == sources[:1] and selection["kind"] == "explicit"


def test_unknown_project_does_not_fall_back_to_library(sources):
    with pytest.raises(SourceSelectionError):
        select_question_sources(sources, "Quelle marge ?", project="missing")


@pytest.mark.parametrize(
    "opening, closing", [('"', '"'), ("“", "”"), ("«", "»"), ("'", "'"), ("‘", "’")]
)
def test_unknown_quoted_full_name_cannot_match_known_suffix(opening, closing):
    sources = [rev("a", "Budget.xlsx"), rev("b", "other.xlsx")]
    with pytest.raises(SourceSelectionError):
        select_question_sources(
            sources, f"Dans {opening}Missing Budget.xlsx{closing}, quelle valeur ?"
        )


@pytest.mark.parametrize("opening, closing", [('"', '"'), ("'", "'"), ("«", "»")])
def test_known_full_quoted_name_wins_over_known_suffix(opening, closing):
    sources = [rev("a", "Missing Budget.xlsx"), rev("b", "Budget.xlsx")]
    assert (
        select_question_sources(sources, f"{opening}Missing Budget.xlsx{closing}")[0]
        == sources[:1]
    )


def test_quoted_filename_with_apostrophe_is_one_literal_source():
    sources = [rev("a", "O'Brien.xlsx"), rev("b", "Brien.xlsx")]
    assert select_question_sources(sources, 'Dans "O\'Brien.xlsx"')[0] == sources[:1]
    assert select_question_sources(sources, "Dans 'O'Brien.xlsx'")[0] == sources[:1]


@pytest.mark.parametrize(
    "opening, closing",
    [('"', '"'), ("“", "”"), ("«", "»"), ("'", "'"), ("‘", "’"), ("[", "]")],
)
def test_quote_pairs_do_not_cross_other_quoted_phrases(opening, closing):
    sources = [rev("a", "Budget.xlsx")]
    question = (
        f"Quel est le {opening}revenu net{closing} dans {opening}Budget.xlsx{closing} ?"
    )
    assert select_question_sources(sources, question)[0] == sources
    with pytest.raises(SourceSelectionError):
        select_question_sources(sources, f"Dans {opening}Missing Budget.xlsx{closing}")


def test_bracketed_workbook_reference_checks_full_name_before_sheet_suffix():
    sources = [rev("a", "Budget.xlsx")]
    with pytest.raises(SourceSelectionError):
        select_question_sources(sources, "[Missing Budget.xlsx]Summary!D3")
    assert select_question_sources(sources, "'[Budget.xlsx]Summary'!D3")[0] == sources


def test_markdown_literal_filename_is_resolved_as_a_whole():
    sources = [rev("a", "Budget.xlsx")]
    with pytest.raises(SourceSelectionError):
        select_question_sources(sources, "Dans `Missing Budget.xlsx`")
    assert select_question_sources(sources, "Dans `Budget.xlsx`")[0] == sources


def test_attached_reference_names_are_text_not_sources():
    # A report imported with its references: the question may name them.
    sources = [rev("r", "rapport.xlsx"), rev("x", "autre.xlsx")]
    linked = ["consolidation.xlsx", "budget.xlsx", "Ventes Région été.xlsx"]
    question = (
        "Dans rapport.xlsx, expliquez Synthèse!B5 via consolidation.xlsx, "
        "Ventes Région été.xlsx et « budget.xlsx »."
    )
    result, selection = select_question_sources(sources, question, linked=linked)
    assert result == sources[:1]
    assert selection == {"kind": "question_filename", "filenames": ["rapport.xlsx"]}
    # Linked names alone never select or widen: the whole library stays in scope.
    assert select_question_sources(sources, "via budget.xlsx", linked=linked)[0] == (
        sources
    )


def test_linked_names_are_only_read_when_needed():
    calls = []

    def linked():
        calls.append(1)
        return ["budget.xlsx"]

    sources = [rev("r", "rapport.xlsx")]
    select_question_sources(sources, "Quelle marge dans rapport.xlsx ?", linked=linked)
    assert calls == []
    select_question_sources(sources, "Quelle marge via budget.xlsx ?", linked=linked)
    assert calls == [1]


def test_missing_name_is_quoted_in_the_message():
    with pytest.raises(SourceSelectionError, match="« inconnu.xlsx »"):
        select_question_sources(
            [rev("r", "rapport.xlsx")], "via inconnu.xlsx", linked=[]
        )
