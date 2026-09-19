"""Source facts must distinguish dimensions, membership, provenance and scope."""

import copy
import json

import pytest

from linexcel import aidoc


def group(**overrides):
    return {
        "id": "g",
        "kind": "group",
        "addr": "E4",
        "bbox": "E4:E103",
        "count": 100,
        "formula": "=C4*D4",
        "value": 72.5,
        "valueSource": "engine",
        "samples": [
            {"addr": address, "value": value}
            for address, value in [
                ("E4", 72.5),
                ("E29", 27),
                ("E54", 75),
                ("E78", 21),
                ("E103", 135),
            ]
        ],
        **overrides,
    }


def test_complete_group_is_not_a_five_cell_sample():
    node = group()
    graph = {
        "nodes": [node, {"id": "out", "formula": "=SUM(E4:E103)"}],
        "edges": [{"source": "g", "target": "out"}],
    }
    own = aidoc.build_dossier(graph, "g")["group_coverage"]
    neighbor = aidoc.build_dossier(graph, "out")["precedents"][0]["group_coverage"]
    assert own["membership"] == "complete_bbox"
    assert own["member_cells"] == own["bbox_dimensions"]["cells"] == 100
    assert own["sampled_cells_in_graph"] == 5
    assert own["other_sampled_cells_in_graph"] == 4
    assert own["representative_in_samples"] is True
    assert own["value_samples_in_dossier"] == 5
    assert neighbor["value_samples_in_dossier"] == 0
    assert neighbor["sampled_cells_in_graph"] == 5


def test_sparse_membership_does_not_depend_on_how_many_values_are_sampled():
    facts = aidoc._group_coverage(
        group(count=8, bbox="E4:E13", samples=[{"addr": "E4"}])
    )
    assert facts["membership"] == "partial_bbox"
    assert facts["member_cells"] == 8
    assert facts["bbox_dimensions"]["cells"] == 10


@pytest.mark.parametrize(
    "overrides",
    [
        {"addr": "Z99"},
        {"count": 101},
        {"count": None},
        {"count": True},
        {"bbox": "E103:E4"},
        {"bbox": "Sheet1:Sheet3!E4:E103"},
        {"samples": [{"addr": ""}]},
        {"samples": [{"addr": "E104"}]},
        {"samples": [{"addr": "E4"}, {"addr": "E4"}]},
    ],
)
def test_missing_or_contradictory_group_evidence_never_claims_complete(overrides):
    assert aidoc._group_coverage(group(**overrides))["membership"] == "unknown"


def test_more_samples_than_members_cannot_be_certified():
    facts = aidoc._group_coverage(
        group(
            count=2,
            bbox="E4:E5",
            samples=[{"addr": "E4"}, {"addr": "E5"}, {"addr": "E6"}],
        )
    )
    assert facts["membership"] == "unknown"
    assert facts["sampled_cells_in_graph"] is None


@pytest.mark.parametrize("reference", ["'Sales Data'!B1:F1048576", "B:F"])
def test_full_columns_have_a_cell_count_not_five_million_rows(reference):
    source = {
        "expr": "VLOOKUP(...) ",
        "evaluated": True,
        "value": "#N/A",
        "inputs": [{"ref": reference, "value": {"range": reference, "n": 5242880}}],
    }
    before = copy.deepcopy(source)
    value = aidoc._compact_steps(source)["inputs"][0]["value"]
    assert value["n_unit"] == "cells"
    assert value["dimensions"] == {"rows": 1048576, "columns": 5, "cells": 5242880}
    assert source == before


@pytest.mark.parametrize(
    "reference",
    [
        "Sheet1:Sheet3!A1:B2",
        "'Sheet1:Sheet3'!A1:B2",
        "Sheet!A1!B2",
        "A1:B2!C3",
        "SUM(A1)!B2",
        "[1]Sheet!A1",
        "'S'!XFE1",
        "A0:A3",
        "A3:A1",
        "Name",
        "A1 B2",
        "A$$1",
        "$$A1",
        "A1$",
    ],
)
def test_dimensions_reject_unresolved_or_nonrectangular_reference(reference):
    assert aidoc._range_dimensions(reference) is None


def test_quoted_sheet_and_absolute_rectangles_are_supported():
    assert aidoc._range_dimensions("'O''Brien''s Café'!$B$2:$D$5") == {
        "rows": 4,
        "columns": 3,
        "cells": 12,
    }


@pytest.mark.parametrize(
    "kind,formula,expected",
    [
        ("input", None, "input_read"),
        ("name", None, "name_resolution"),
        ("cell", "=B2*2", "formula_result"),
        (None, None, "engine_value_without_formula"),
    ],
)
def test_engine_provenance_is_not_always_formula_recalculation(kind, formula, expected):
    facts = aidoc._value_evidence(
        {"kind": kind, "formula": formula, "valueSource": "engine", "value": 42}
    )
    assert facts["value_origin_kind"] == expected
    assert facts["value_source"] == "engine"


def test_false_evaluation_flag_does_not_assert_no_attempt():
    result = aidoc._compact_steps(
        {"expr": "SUM(A1:A2)", "evaluated": False, "value": 123}
    )
    assert result["value"] == "no evaluated value recorded"
    assert result["evaluation_attempt"] == "unknown"
    assert result["evaluated"] is False


def test_sales_topology_cannot_be_read_as_a_linear_list():
    edges = [
        {"source": a, "target": b}
        for a, b in [("Sales", "C5"), ("Sales", "C6"), ("Sales", "C7"), ("C5", "C7")]
    ]
    facts = aidoc.build_workbook_dossier({"nodes": [], "edges": edges})[
        "graph_connections"
    ]
    assert facts["edges"] == edges
    assert facts["branching_observed"] and facts["merging_observed"]
    assert facts["omitted_from_dossier"] == 0


def test_workbook_trim_preserves_edges_and_counts_omitted_patterns():
    graph = {
        "nodes": [
            {"id": f"n{i}", "kind": "cell", "formula": "=1", "value": "x" * 500}
            for i in range(40)
        ],
        "edges": [{"source": "n0", "target": f"n{i}"} for i in range(1, 40)],
    }
    dossier = aidoc.build_workbook_dossier(graph)
    assert dossier["formula_pattern_coverage"]["omitted"] == 20
    connections = copy.deepcopy(dossier["graph_connections"])
    # Force the existing preview-first/pattern-tail trimming stages.
    dossier["sheets"] = [
        {"name": "S", "preview": [{"values": ["x" * 700]} for _ in range(20)]}
    ]
    fitted = json.loads(aidoc._fit_workbook_dossier(dossier))
    assert fitted["graph_connections"] == connections
    assert len(connections["edges"]) == aidoc.MAX_DOCUMENT_CONNECTIONS
    assert connections["omitted_from_dossier"] == 39 - aidoc.MAX_DOCUMENT_CONNECTIONS
    assert fitted["formula_pattern_coverage"] == {
        "scope": "graph_formula_nodes",
        "total": 40,
        "shown": 5,
        "omitted": 35,
    }


def test_node_trim_keeps_graph_sample_counts_but_updates_shown_samples():
    node = group(value="v" * 10000, samples=[{"addr": "E4", "value": "s" * 3000}])
    dossier = json.loads(
        aidoc._fit_node_dossier(
            aidoc.build_dossier({"nodes": [node], "edges": []}, "g")
        )
    )
    assert dossier["value_samples"] == []
    assert dossier["group_coverage"]["sampled_cells_in_graph"] == 1
    assert dossier["group_coverage"]["value_samples_in_dossier"] == 0
    assert dossier["group_coverage"]["membership"] == "complete_bbox"
    assert dossier["neighbor_coverage"]["omissions_scope"] == "dossier_only"


@pytest.mark.parametrize("language", aidoc._LANGUAGES)
def test_all_languages_request_the_explicit_facts(language):
    node, workbook = aidoc._SYSTEM[language], aidoc._WORKBOOK_SYSTEM[language]
    assert all(
        key in node
        for key in [
            "group_coverage.membership",
            "value_origin_kind",
            "dimensions.rows/columns/cells",
            "evaluation_attempt",
        ]
    )
    assert all(
        key in workbook
        for key in [
            "graph_connections.edges",
            "branching_observed",
            "omitted_from_dossier",
            "group_coverage.membership",
        ]
    )
