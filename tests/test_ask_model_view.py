"""What the model receives is compact and bounded; the reader's evidence is not."""

import json
import re

from test_ask import call, read, run, select_handle, tools  # noqa: F401

from linexcel.aidoc import TokenUsage
from linexcel.ask import ANSWER_BYTES, AskLimits, encoded
from linexcel.ask_provider import ToolResponse
from linexcel.ask_view import RESULT_BYTES, model_view

SHA = "5" * 64
REVISION = {"id": f"{'a' * 32}:{SHA}", "projectId": "a" * 32, "sha256": SHA}


def source_cell(index, **extra):
    return {
        "id": f"lx1.node{index}",
        "handle": f"lx1.node{index}",
        "sourceNode": f"S!A{index}",
        "revision": REVISION,
        "filename": "book.xlsx",
        "sheet": "S",
        "address": f"A{index}",
        "kind": "cell",
        "labels": [
            {
                "text": "Revenue",
                "sheet": "S",
                "cell": "A1",
                "method": "above_text",
                "role": "column",
                "truncated": False,
                "distance": 1,
                "blank_gap": 0,
            }
        ],
        "value": index,
        "valueProvenance": "saved_cache",
        "cacheStatus": "saved",
        "formula": f"=A{index + 1}",
        "sourceValueProvenance": {"workbookSha256": SHA, "rawXmlValue": str(index)},
        **extra,
    }


def lineage(count):
    nodes = [source_cell(index) for index in range(1, count + 1)]
    return {
        "revision": REVISION,
        "filename": "book.xlsx",
        "direction": "upstream",
        "nodes": nodes,
        "edges": [
            {"source": nodes[i + 1]["id"], "target": nodes[i]["id"], "kind": "formula"}
            for i in range(count - 1)
        ],
        "boundaries": [
            {"node": nodes[-1]["id"], "reason": "input_constant"},
            {
                "node": nodes[0]["id"],
                "reason": "unresolved_external",
                "linkId": "9" * 64,
                "relationshipPath": "/Users/someone/dir/Ventes%20%C3%A9t%C3%A9.xlsx",
                "revision": REVISION,
            },
            {
                "node": nodes[0]["id"],
                "reason": "cache_only",
                "cached": [{"row": 2, "column": 2, "value": 6600.000000000001}],
            },
        ],
        "stronglyConnectedComponents": [],
        "revisions": [REVISION],
        "coverage": {
            "complete": False,
            "exhaustive": True,
            "scope": "loaded owner revisions only",
            "nodesReturned": count,
            "edgesReturned": count - 1,
            "depth": 20,
            "capsHit": [],
            "boundaries": ["cache_only", "input_constant", "unresolved_external"],
        },
        "trace": {"generation": "g" * 32, "revisions": [REVISION]},
        "untrustedContent": True,
    }


def test_a_large_lineage_is_cut_to_its_nearest_nodes_and_says_so():
    full = lineage(172)
    view, omitted = model_view("trace_lineage", full)
    blob = encoded(view)
    assert len(blob.encode()) <= RESULT_BYTES
    assert SHA not in blob and "sourceValueProvenance" not in blob
    assert omitted["nodes"] == 172 - len(view["nodes"]) > 0
    assert view["omittedForModel"] == omitted and "Narrow" in view["hint"]
    # Breadth-first order: the root and its nearest precedents are what remains.
    assert [node["address"] for node in view["nodes"][:3]] == ["A1", "A2", "A3"]
    kept = {node["handle"] for node in view["nodes"]}
    assert view["edges"] and all(
        edge["source"] in kept and edge["target"] in kept for edge in view["edges"]
    )
    # A constant input is not a limit. A missing workbook is: the model gets its
    # file name and the cached values, never the folder on the author's machine.
    missing, cache = view["boundaries"]
    assert missing == {
        "node": "lx1.node1",
        "reason": "unresolved_external",
        "filename": "Ventes été.xlsx",
    }
    assert cache["cachedInCitingWorkbook"][0]["value"] == 6600.000000000001
    assert "someone" not in blob and "9" * 64 not in blob
    assert view["coverage"]["complete"] is False
    # The reader's evidence is the untouched full result.
    assert len(full["nodes"]) == 172 and full["nodes"][0]["revision"] == REVISION


def test_a_small_result_is_delivered_whole():
    view, omitted = model_view("trace_lineage", lineage(3))
    assert omitted == {} and "omittedForModel" not in view
    assert len(view["nodes"]) == 3 and len(view["edges"]) == 2
    assert view["nodes"][0]["labels"] == [
        {"text": "Revenue", "role": "column", "cell": "A1"}
    ]


def test_search_view_drops_per_sheet_coverage_and_keeps_the_verdict():
    sheets = [
        {"revision": REVISION, "sheet": f"S{i}", "byKind": {}} for i in range(120)
    ]
    data = {
        "generation": "g" * 32,
        "results": [
            {
                **source_cell(1),
                "occurrenceCount": 3,
                "occurrences": [
                    {"address": f"A{i}", "handle": f"lx1.node{i}"} for i in range(1, 4)
                ],
                "name": None,
                "objectId": "S!A1",
            }
        ],
        "pagination": {"offset": 0, "limit": 20, "matched": 1, "hasMore": False},
        "coverage": {
            "complete": False,
            "capsHit": ["build_seconds"],
            "nodesTotal": 125365,
            "nodesScanned": 4000,
            "nodesUnscanned": 121365,
            "numericPolicy": "Labelled finite numeric cells only",
            "labelMethods": "x" * 200,
            "sheets": sheets,
        },
        "trace": {"terms": ["revenue"], "revisions": [REVISION]},
        "untrustedContent": True,
    }
    view, omitted = model_view("search_catalog", data)
    assert not omitted and len(encoded(view).encode()) < 2048
    assert view["coverage"] == {
        "complete": False,
        "capsHit": ["build_seconds"],
        "nodesTotal": 125365,
        "nodesScanned": 4000,
        "nodesUnscanned": 121365,
        "numericPolicy": "Labelled finite numeric cells only",
    }
    assert view["matched"] == 1 and view["hasMore"] is False
    assert view["results"][0]["occurrenceCount"] == 3
    assert len(view["results"][0]["occurrences"]) == 3


def test_link_view_keeps_cached_values_and_drops_the_author_machine_path():
    data = {
        "links": [
            {
                "bookKey": "2",
                "relationshipPath": "/Users/someone/private/Sources/budget.xlsx",
                "filename": "budget.xlsx",
                "reference": "[2]Budget!$B$4",
                "sheet": "Budget",
                "address": "B4",
                "cached": [{"row": 4, "column": 2, "value": 6000.0}],
                "sourceRevision": REVISION,
                "sourceNode": "Conso!B2",
                "status": "unresolved",
                "targetRevision": None,
                "candidates": [],
                "pin": None,
            }
        ],
        "pagination": {"matched": 1, "hasMore": False},
        "coverage": {"complete": True, "capsHit": []},
    }
    view, _ = model_view("list_external_links", data)
    blob = encoded(view)
    assert "someone" not in blob and SHA not in blob
    link = view["links"][0]
    assert link["filename"] == "budget.xlsx" and link["status"] == "unresolved"
    assert link["cachedInCitingWorkbook"] == [{"row": 4, "column": 2, "value": 6000.0}]
    assert link["sourceProject"] == "a" * 32 and link["pinned"] is False


def test_long_text_is_clipped_for_the_model_only():
    cell = source_cell(1, value="x" * 5000, formula="=" + "A1+" * 900)
    view, _ = model_view("get_cell", cell)
    assert len(view["value"]) == 301 and view["value"].endswith("…")
    assert len(view["formula"]) == 601
    assert len(cell["value"]) == 5000


def tool_payloads(provider):
    return [
        message["content"]
        for message in provider.messages[-1]
        if message["role"] == "tool"
    ]


def test_run_delivers_the_compact_view_and_counts_what_it_sent(tools):  # noqa: F811
    engine, source = tools
    result, provider = run(
        engine,
        [
            read(source),
            lambda messages: [
                call(
                    "submit_answer",
                    {"status": "proposed", "handles": [select_handle(messages, "B2")]},
                    "end",
                )
            ],
        ],
    )
    assert result["status"] == "answer_proposed"
    (payload,) = tool_payloads(provider)
    assert "sha256" not in payload and "sourceValueProvenance" not in payload
    assert not re.search(r"[0-9a-f]{64}", payload)
    read_step = next(
        step for step in result["trace"]["steps"] if step["tool"] == "read_range"
    )
    assert read_step["budgetUse"]["resultBytes"] == len(payload.encode())
    # The answer keeps the full evidence the reader needs.
    cell = result["answer"]["cells"][0]
    assert cell["revision"]["sha256"] and cell["handle"].startswith("lx1.")


def test_a_verified_answer_is_not_discarded_for_its_own_size(tools):  # noqa: F811
    engine, source = tools
    limits = AskLimits(result_bytes=4096)
    result, _ = run(
        engine,
        [
            read(source),
            lambda messages: [
                call(
                    "submit_answer",
                    {
                        "status": "proposed",
                        "handles": [
                            select_handle(messages, "B2"),
                            select_handle(messages, "A2"),
                            select_handle(messages, "A3"),
                        ],
                    },
                    "end",
                )
            ],
        ],
        limits=limits,
    )
    assert result["status"] == "answer_proposed"
    answer_bytes = len(encoded(result["answer"]).encode())
    assert limits.result_bytes < answer_bytes < ANSWER_BYTES
    assert result["trace"]["budgetUse"]["resultBytes"] < limits.result_bytes


def test_a_misnamed_sheet_or_oversized_range_can_be_corrected(tools):  # noqa: F811
    engine, source = tools

    def wrong(sheet, cells, cid):
        return call(
            "read_range",
            {"project": source.project_id, "sheet": sheet, "range": cells},
            cid,
        )

    def corrected(messages):
        error = json.loads(messages[-1]["content"])
        assert error["trust"] == "linexcel"
        assert error["error"]["code"] == "invalid_range"
        assert "25 cells" in error["error"]["hint"]
        return read(source, "third")

    result, provider = run(
        engine,
        [
            [wrong("TEE_COMPLET", "A1:B3", "first")],
            lambda messages: (
                json.loads(messages[-1]["content"])["error"]["code"] == "unknown_sheet"
                and [wrong("S", "A1:Z60", "second")]
            ),
            corrected,
            lambda messages: [
                call(
                    "submit_answer",
                    {"status": "proposed", "handles": [select_handle(messages)]},
                    "end",
                )
            ],
        ],
    )
    assert result["status"] == "answer_proposed"
    verdicts = [
        (step["tool"], step["verdict"], step["summary"].get("rejected"))
        for step in result["trace"]["steps"]
        if step["tool"] == "read_range"
    ]
    assert verdicts == [
        ("read_range", "rejected", "unknown_sheet"),
        ("read_range", "rejected", "invalid_range"),
        ("read_range", "completed", None),
    ]
    assert result["trace"]["budgetUse"]["steps"] == 4
    # No workbook text and no exception text reached the model with the error.
    assert "TEE_COMPLET" not in provider.messages[1][-1]["content"]


def test_repeated_invalid_arguments_still_end_the_run(tools):  # noqa: F811
    engine, source = tools
    responses = [
        [
            call(
                "read_range",
                {"project": source.project_id, "sheet": "Nope", "range": "A1"},
                f"bad{index}",
            )
        ]
        for index in range(4)
    ]
    result, provider = run(engine, responses)
    assert result["status"] == "failed" and result["answer"] is None
    assert result["trace"]["reason"] == "invalid_or_unavailable_tool_provider"
    assert len(provider.messages) == 4


def test_an_oversized_request_is_stopped_before_it_is_sent(tools):  # noqa: F811
    engine, _ = tools
    # The first request fits. Its tool result then makes the second one too
    # large for what is left, so that request is never made. The provider's own
    # count for the first request calibrates the estimate of the second.
    engine.scope["revisions"][0]["sheets"] = [f"Sheet {i}" for i in range(900)]
    first = ToolResponse(
        [call("list_workbooks")],
        "tool_calls",
        TokenUsage(input_tokens=1300, output_tokens=20, requests=1),
    )
    result, provider = run(
        engine,
        [first, [call("submit_answer", {"status": "not_found"})]],
        limits=AskLimits(tokens=4000),
    )
    assert result["status"] == "not_found"
    assert result["trace"]["reason"] == "token_budget"
    assert len(provider.messages) == 1
    assert result["trace"]["budgetUse"]["tokens"] == 1320


def test_a_rejected_proposal_can_be_resubmitted_with_cells(tools):  # noqa: F811
    engine, source = tools
    result, provider = run(
        engine,
        [
            read(source),
            [call("submit_answer", {"status": "proposed", "handles": ["c999"]}, "a")],
            lambda messages: (
                json.loads(messages[-1]["content"])["error"]["code"] == "invalid_handle"
                and [
                    call(
                        "submit_answer",
                        {
                            "status": "proposed",
                            "handles": [select_handle(messages[:-2], "B2")],
                        },
                        "b",
                    )
                ]
            ),
        ],
    )
    assert result["status"] == "answer_proposed"
    assert result["answer"]["cells"][0]["address"] == "B2"
    assert (
        "cells only"
        in json.loads(provider.messages[-1][-1]["content"])["error"]["hint"]
    )
