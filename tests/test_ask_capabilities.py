"""Run-local model capabilities and executable deterministic computation evidence."""

import hashlib
import json
import re

import pytest
from test_ask import call, read, run, select_handle
from test_ask import tools as tools

from linexcel.ask import AskLimits, QuestionRun
from linexcel.ask_tools import calculate, replay_computation


def test_model_uses_short_aliases_but_answer_and_trace_keep_authority(tools):
    engine, source = tools
    aliases = []

    def reread(messages):
        alias = select_handle(messages)
        assert re.fullmatch(r"c[1-9][0-9]*", alias)
        aliases.append(alias)
        return [call("get_cell", {"handle": alias}, "reread")]

    def trace(messages):
        cell = json.loads(messages[-1]["content"])["data"]
        assert cell["handle"] == aliases[0] and cell["value"] == 7
        return [
            call(
                "trace_lineage",
                {
                    "handle": cell["handle"],
                    "direction": "upstream",
                    "depth": 2,
                },
                "trace",
            )
        ]

    def finish(messages):
        lineage = json.loads(messages[-1]["content"])["data"]
        assert all(not n["id"].startswith("lx1.") for n in lineage["nodes"])
        return [
            call(
                "submit_answer",
                {
                    "status": "proposed",
                    "handles": [aliases[0]],
                },
                "finish",
            )
        ]

    result, provider = run(engine, [read(source), reread, trace, finish])
    assert result["status"] == "answer_proposed"
    assert result["answer"]["cells"][0]["handle"].startswith("lx1.")
    assert all(
        h.startswith("lx1.")
        for s in result["trace"]["steps"]
        for h in s.get("summary", {}).get("handles", [])
    )
    tool_payloads = [
        m["content"]
        for messages in provider.messages
        for m in messages
        if m["role"] == "tool"
    ]
    assert all("lx1." not in blob for blob in tool_payloads)


def test_alias_translation_never_rewrites_workbook_values_or_formulas(tools):
    engine, _ = tools
    from test_ask import FakeProvider

    question = QuestionRun(FakeProvider([]), engine, "Find", AskLimits())
    question._model_aliases["lx1.source"] = "c1"
    result = question.model_result(
        {
            "handle": "lx1.source",
            "id": "lx1.source",
            "source": "lx1.source",
            "stronglyConnectedComponents": [["lx1.source"]],
            "value": "lx1.source",
            "formula": "lx1.source",
            "labels": [{"text": "lx1.source"}],
        }
    )
    assert result["handle"] == result["id"] == result["source"] == "c1"
    assert result["stronglyConnectedComponents"] == [["c1"]]
    assert result["value"] == result["formula"] == "lx1.source"
    assert result["labels"][0]["text"] == "lx1.source"


def test_graph_revealed_source_alias_can_be_reread_computed_and_submitted(tools):
    engine, source = tools
    first_alias = []

    def trace_root(messages):
        root = select_handle(messages, "B2")
        first_alias.append(root)
        return [
            call(
                "trace_lineage",
                {
                    "handle": root,
                    "direction": "upstream",
                    "depth": 2,
                },
                "trace",
            )
        ]

    def read_graph_source(messages):
        graph = json.loads(messages[-1]["content"])["data"]
        source_cell = next(n for n in graph["nodes"] if n["address"] == "A2")
        assert source_cell["handle"] == source_cell["id"]
        assert source_cell["handle"] != first_alias[0]
        return [call("get_cell", {"handle": source_cell["handle"]}, "read-source")]

    result, _ = run(
        engine,
        [
            [
                call(
                    "read_range",
                    {"project": source.project_id, "sheet": "S", "range": "B2"},
                    "root",
                )
            ],
            trace_root,
            read_graph_source,
            lambda messages: [
                call(
                    "compute",
                    {
                        "op": "sum",
                        "operands": [
                            json.loads(messages[-1]["content"])["data"]["handle"]
                        ],
                    },
                    "sum",
                )
            ],
            lambda messages: [
                call(
                    "submit_answer",
                    {
                        "status": "proposed",
                        "compute": json.loads(messages[-1]["content"])["data"][
                            "computeHandle"
                        ],
                    },
                    "finish",
                )
            ],
        ],
    )
    assert result["status"] == "answer_proposed"
    assert [(c["address"], c["value"]) for c in result["answer"]["cells"]] == [
        ("A2", 7)
    ]
    assert result["answer"]["computation"]["value"] == 7
    assert result["answer"]["cells"][0]["handle"].startswith("lx1.")


@pytest.mark.parametrize(
    "alter", [lambda h: h + "=", lambda h: h.upper(), lambda h: "c9999"]
)
def test_altered_or_unissued_alias_never_becomes_authority(tools, alter):
    engine, source = tools
    result, _ = run(
        engine,
        [
            read(source),
            lambda messages: [
                call(
                    "get_cell",
                    {"handle": alter(select_handle(messages))},
                    "bad",
                )
            ],
        ],
    )
    assert result["status"] == "failed" and result["answer"] is None
    assert result["trace"]["reason"] == "invalid_source_handle"


def test_reused_tools_cannot_inherit_source_or_computation_capabilities(tools):
    engine, source = tools
    first, _ = run(
        engine,
        [
            read(source),
            lambda messages: [
                call(
                    "submit_answer",
                    {"status": "proposed", "handles": [select_handle(messages)]},
                    "end",
                )
            ],
        ],
    )
    assert first["status"] == "answer_proposed"
    old_full = first["answer"]["cells"][0]["handle"]
    engine.computations["lc1.1"] = {"op": "sum", "operands": [old_full]}
    second, _ = run(engine, [[call("get_cell", {"handle": "c2"})]])
    assert second["trace"]["reason"] == "invalid_source_handle"
    assert not engine.issued and not engine.computations
    third, _ = run(
        engine,
        [
            [
                call(
                    "submit_answer",
                    {
                        "status": "proposed",
                        "handles": [old_full],
                        "compute": "lc1.1",
                    },
                )
            ]
        ],
    )
    assert third["status"] == "not_found" and third["answer"] is None


@pytest.mark.parametrize(
    "op,operands,value,error",
    [
        ("sum", [{"value": 7}, {"value": None}, {"value": 3}], 10, None),
        ("count", [{"value": 7}, {"value": None}], 1, None),
        ("ratio", [{"value": 7}, {"value": 0}], None, "division_by_zero"),
        ("sum", [{"value": 7}, {"value": "#REF!"}], None, "source_error"),
        (
            "sum",
            [
                {"value": 7},
                {"value": None, "formula": "=A1", "valueProvenance": "saved_cache"},
            ],
            None,
            "saved_value_unavailable",
        ),
        ("sum", [{"value": float("nan")}], None, "nonfinite_operand"),
        (
            "sum",
            [{"value": 7}, {"value": '__import__("os").system("unsafe")'}],
            None,
            "nonnumeric_operand",
        ),
    ],
)
def test_server_recipe_executes_same_allowlisted_result(
    op, operands, value, error, capsys
):
    computed = calculate(op, operands)
    receipt = computed["receipt"]
    assert receipt["engineVersion"] == 1
    assert hashlib.sha256(receipt["code"].encode()).hexdigest() == receipt["codeSha256"]
    # This test executes the server-authored downloadable recipe. Production
    # passes only typed operations to calculate; no model code is evaluated.
    exec(compile(receipt["code"], "<server-computation-recipe>", "exec"), {})
    replay = json.loads(capsys.readouterr().out)
    assert replay == replay_computation(receipt)
    assert replay["value"] == value and replay["error"] == error
    assert replay["blanks"] == computed["blanks"]
    assert replay["missingSavedValues"] == computed["missingSavedValues"]


def test_recipe_rejects_model_code_and_changed_engine_without_execution():
    receipt = calculate("sum", [{"value": 7}])["receipt"]
    for altered in [
        {**receipt, "code": "raise AssertionError('must not execute')"},
        {**receipt, "engineVersion": 2},
        {**receipt, "operation": "eval"},
    ]:
        with pytest.raises(ValueError):
            replay_computation(altered)


def test_compute_tool_rejects_model_code_before_calculation(tools):
    engine, _ = tools
    with pytest.raises(ValueError, match="Unknown or missing"):
        engine.execute("compute", {"op": "sum", "operands": [], "code": "1+1"})


def test_computation_receipt_uses_reread_canonical_operands_but_model_does_not(tools):
    engine, source = tools
    seen = []

    def finish(messages):
        data = json.loads(messages[-1]["content"])["data"]
        seen.append(data)
        assert "code" not in data["receipt"] and "operands" not in data["receipt"]
        assert all(re.fullmatch(r"c[1-9][0-9]*", c["handle"]) for c in data["operands"])
        return [
            call(
                "submit_answer",
                {"status": "proposed", "compute": data["computeHandle"]},
                "end",
            )
        ]

    result, _ = run(
        engine,
        [
            read(source),
            lambda messages: [
                call(
                    "compute",
                    {
                        "op": "sum",
                        "operands": [
                            select_handle(messages, "A2"),
                            select_handle(messages, "A3"),
                        ],
                    },
                    "sum",
                )
            ],
            finish,
        ],
    )
    assert result["status"] == "answer_proposed"
    computation = result["answer"]["computation"]
    assert computation["value"] == 10
    receipt = computation["receipt"]
    assert receipt["operation"] == "sum" and len(receipt["operands"]) == 2
    assert all(
        c["handle"].startswith("lx1.") and c["revision"]["sha256"]
        for c in receipt["operands"]
    )
    assert [c["value"] for c in receipt["operands"]] == [7, 3]
    assert replay_computation(receipt)["value"] == 10
    assert all(c["handle"] in engine.issued for c in receipt["operands"])
