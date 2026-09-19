"""Documentation must preserve uncertainty rather than manufacture proof."""

import json

import pytest

from linexcel.aidoc import (
    TokenUsage,
    _unwrap_markdown,
    build_dossier,
    describe_images,
    document_nodes,
)


@pytest.mark.parametrize("source", ["file", "volatile", "external", None])
def test_cached_or_unknown_value_is_not_called_computed(source):
    node = {"id": "S!A1", "value": 123, "cachedValue": 123}
    if source:
        node["valueSource"] = source
    dossier = build_dossier({"nodes": [node], "edges": []}, node["id"])
    assert dossier is not None
    assert dossier["computed_value"] is None
    assert dossier["displayed_value"] == 123
    assert dossier["value_source"] == (source or "unknown")


def test_provider_receives_disagreement_and_partial_neighbor_evidence():
    node = {
        "id": "S!A1",
        "value": 0,
        "valueSource": "engine",
        "cachedValue": 42,
        "cachedAgreement": "differ",
        "groupCachedAgreement": "differ",
        "steps": {"expr": "=1-1", "evaluated": False, "value": 0},
    }
    neighbors = [
        {"id": f"S!B{i}", "value": i, "valueSource": "file"} for i in range(1, 36)
    ]
    graph = {
        "meta": {"engine": "formualizer (Rust)"},
        "nodes": [node, *neighbors],
        "edges": [{"source": n["id"], "target": node["id"]} for n in neighbors],
    }
    received = []

    def provider(system, user, **kwargs):
        received.append(json.loads(user.split("\n", 1)[1]))
        return "A value differs from its cache; no evaluated breakdown available."

    document_nodes(graph, [node["id"]], provider=provider)
    dossier = received[0]
    assert dossier["computed_value"] == 0
    assert dossier["recalculation_engine"] == "formualizer (Rust)"
    assert dossier["cached_value"] == 42
    assert dossier["cached_agreement"] == "differ"
    assert dossier["group_cached_agreement"] == "differ"
    assert dossier["neighbor_coverage"]["precedents_total"] == 35
    assert dossier["neighbor_coverage"]["precedents_omitted"] == (
        35 - len(dossier["precedents"])
    )
    assert all(n["value_source"] == "file" for n in dossier["precedents"])
    # The dossier can omit an oversized breakdown, but cannot label it evaluated.
    decomposition = dossier["decomposition"]
    assert not isinstance(decomposition, dict) or not decomposition["evaluated"]


def test_batch_reads_graph_edges_once():
    class CountingEdges(list):
        reads = 0

        def __iter__(self):
            self.reads += 1
            return super().__iter__()

    nodes = [{"id": str(i), "value": i, "valueSource": "engine"} for i in range(60)]
    edges = CountingEdges({"source": str(i), "target": str(i + 1)} for i in range(59))
    docs = document_nodes(
        {"nodes": nodes, "edges": edges},
        [str(i) for i in range(60)],
        provider=lambda *args, **kwargs: "Documented.",
    )
    assert len(docs) == len(nodes)
    assert edges.reads == 1


def test_self_loop_is_both_precedent_and_dependent():
    dossier = build_dossier(
        {
            "nodes": [{"id": "S!A1", "formula": "=A1+1"}],
            "edges": [{"source": "S!A1", "target": "S!A1"}],
        },
        "S!A1",
    )
    assert dossier is not None
    assert len(dossier["precedents"]) == len(dossier["dependents"]) == 1


@pytest.mark.parametrize("fence", ["markdown", "md", "Markdown"])
def test_whole_markdown_wrapper_is_removed_from_node_cards(fence):
    docs = document_nodes(
        {"nodes": [{"id": "S!A1"}], "edges": []},
        ["S!A1"],
        provider=lambda *args, **kwargs: f"```{fence}\n**Description**\n```",
    )
    assert docs == {"S!A1": "**Description**"}


def test_real_code_fences_and_surrounding_prose_are_preserved():
    for text in (
        "```python\nprint('hello')\n```",
        "Example:\n```markdown\n**bold**\n```",
        "```markdown\n**bold**\n```\nEnd.",
    ):
        assert _unwrap_markdown(text) == text


def test_vision_card_unwraps_only_the_complete_markdown_envelope():
    class Provider:
        def generate(self, system_prompt, user_prompt, **kwargs):
            return "Unused"

        def generate_with_image(self, system_prompt, user_prompt, image, **kwargs):
            return "```markdown\n**Layout**: two columns.\n```", TokenUsage()

    assert describe_images({"Sheet": b"synthetic image"}, provider=Provider()) == {
        "Sheet": "**Layout**: two columns."
    }


def test_no_matching_names_does_not_mean_no_workbook_names():
    from linexcel.doc_evidence import relevant_names

    evidence = {
        "status": "complete",
        "total": 1,
        "definitions": [{"name": "Threshold", "scope": "workbook"}],
    }
    facts = relevant_names(evidence, ["A1"], "S")
    assert facts["definitions"] == []
    assert facts["selection_scope"] == "names_referenced_by_this_formula_only"
    assert facts["workbook_definition_count"] == 1


def test_overview_retains_recalculation_evidence_without_live_api_details():
    from linexcel.aidoc import build_workbook_dossier
    from linexcel.execution import add_coverage

    graph = {
        "meta": {"execution": {"status": "completed", "engineAvailable": False}},
        "nodes": [
            {
                "id": "g:S!A1#100",
                "kind": "group",
                "formula": "=1+2",
                "sheet": "S",
                "addr": "A1",
                "count": 100,
                "value": 3,
                "valueSource": "engine",
            }
        ],
        "edges": [],
    }
    add_coverage(graph)
    dossier = build_workbook_dossier(graph)
    assert "engineAvailable" not in dossier["execution"]
    assert dossier["value_coverage"]["counts"]["engine"] == 1
    (pattern,) = dossier["formula_patterns"]
    assert pattern["displayed_value"] == 3
    assert pattern["value_source"] == "engine"
    assert pattern["representative_cell"] == {"sheet": "S", "address": "A1"}
    assert graph["meta"]["execution"]["engineAvailable"] is False


def test_name_metadata_reader_status_is_not_a_failed_engine_evaluation():
    from linexcel.aidoc import build_dossier, build_workbook_dossier

    definition = {
        "name": "Threshold",
        "scope": "workbook",
        "expression": "S!$B$1",
        "expression_evaluated": False,
    }
    graph = {
        "meta": {
            "definedNameEvidence": {"status": "complete", "definitions": [definition]}
        },
        "nodes": [{"id": "c:S!A1", "sheet": "S", "formula": "=Threshold"}],
        "edges": [],
    }
    for dossier in (build_dossier(graph, "c:S!A1"), build_workbook_dossier(graph)):
        facts = dossier["source_defined_names"]
        assert facts["definition_provenance"] == "source_workbook_metadata"
        (name,) = facts["definitions"]
        assert name["expression"] == "S!$B$1"
        assert "expression_evaluated" not in name
    assert definition["expression_evaluated"] is False
