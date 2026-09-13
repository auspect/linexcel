"""Incomplete model calls must not become successful documentation."""

import json
from types import SimpleNamespace

import pytest

from linexcel.aidoc import (
    MAX_DOSSIER_CHARS,
    AiDocError,
    TokenUsage,
    _fit_node_dossier,
    _OpenAICompatProvider,
    build_dossier,
    describe_images,
    document_nodes,
    document_workbook,
)


@pytest.mark.parametrize("vision", [False, True])
@pytest.mark.parametrize(
    "reason,text", [("length", "Cut off"), ("content_filter", "Blocked"), ("stop", " ")]
)
def test_provider_rejects_incomplete_calls_and_preserves_usage(vision, reason, text):
    response = SimpleNamespace(
        choices=[
            SimpleNamespace(message=SimpleNamespace(content=text), finish_reason=reason)
        ],
        usage=SimpleNamespace(prompt_tokens=10, completion_tokens=20),
    )
    provider = object.__new__(_OpenAICompatProvider)
    provider._model = "test"
    provider._client = SimpleNamespace(
        chat=SimpleNamespace(
            completions=SimpleNamespace(create=lambda **kwargs: response)
        )
    )
    with pytest.raises(AiDocError) as caught:
        if vision:
            provider.generate_with_image("system", "user", b"png")
        else:
            provider.generate_with_usage("system", "user")
    assert caught.value.usage.total == 30


@pytest.mark.parametrize("text", ["", " ", "```markdown\n\n```", "{{T1}}"])
def test_empty_documents_fail_after_markdown_cleanup(text):
    graph = {"nodes": [{"id": "S!A1"}], "edges": []}

    def provider(*args, **kwargs):
        return text

    with pytest.raises(AiDocError, match="empty response"):
        document_nodes(graph, ["S!A1"], provider=provider)
    with pytest.raises(AiDocError, match="empty response"):
        document_workbook(graph, provider=provider)


def test_failed_calls_count_towards_budget_before_more_nodes_are_sent():
    calls = []

    class Provider:
        def generate(self, *args, **kwargs):
            raise AssertionError("usage method expected")

        def generate_with_usage(self, *args, **kwargs):
            calls.append(1)
            raise AiDocError(
                "truncated",
                usage=TokenUsage(input_tokens=10, output_tokens=20, requests=1),
            )

    usage = TokenUsage()
    with pytest.warns(UserWarning, match="never sent"), pytest.raises(AiDocError):
        document_nodes(
            {"nodes": [{"id": "a"}, {"id": "b"}], "edges": []},
            ["a", "b"],
            provider=Provider(),
            usage=usage,
            token_budget=25,
            max_workers=1,
        )
    assert len(calls) == 1
    assert usage.total == 30


def test_empty_vision_is_failure_and_keeps_consumed_tokens():
    class Provider:
        def generate(self, *args, **kwargs):
            return "unused"

        def generate_with_image(self, *args, **kwargs):
            return " ", TokenUsage(input_tokens=10, output_tokens=2, requests=1)

    usage = TokenUsage()
    with pytest.raises(AiDocError, match="empty response"):
        describe_images({"sheet": b"image"}, provider=Provider(), usage=usage)
    assert usage.total == 12


def test_giant_source_formula_is_explicitly_omitted_without_partial_formula():
    dossier = build_dossier(
        {
            "nodes": [{"id": "S!A1", "formula": "=1+" * 10000, "value": "v" * 20000}],
            "edges": [],
        },
        "S!A1",
    )
    blob = _fit_node_dossier(dossier)
    assert len(blob) <= MAX_DOSSIER_CHARS
    decoded = json.loads(blob)
    assert decoded["formula"]["omitted"] == "dossier size limit"
    assert decoded["formula"]["original_characters"] == 30000
