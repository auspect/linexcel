"""Actual OpenAI-compatible SDK with offline HTTP, including M0 negotiation."""

import json

import pytest

from linexcel.ask_provider import OpenAIToolProvider
from linexcel.ask_tools import TOOL_SCHEMAS
from linexcel.web import WebConfig

pytestmark = pytest.mark.ai


def test_tool_provider_reuses_negotiated_token_parameter(tmp_path):
    httpx = pytest.importorskip("httpx")
    sent = []

    def respond(request):
        data = json.loads(request.content)
        sent.append(data)
        if "max_tokens" in data:
            return httpx.Response(
                400,
                json={
                    "error": {
                        "message": "Unsupported parameter: 'max_tokens'. "
                        "Use 'max_completion_tokens' instead.",
                        "param": "max_tokens",
                        "type": "invalid_request_error",
                    }
                },
            )
        return httpx.Response(
            200,
            json={
                "id": "offline",
                "object": "chat.completion",
                "created": 0,
                "model": "tools-test",
                "choices": [
                    {
                        "index": 0,
                        "finish_reason": "tool_calls",
                        "message": {
                            "role": "assistant",
                            "content": None,
                            "tool_calls": [
                                {
                                    "id": "end",
                                    "type": "function",
                                    "function": {
                                        "name": "submit_answer",
                                        "arguments": '{"status":"not_found"}',
                                    },
                                }
                            ],
                        },
                    }
                ],
                "usage": {
                    "prompt_tokens": 10,
                    "completion_tokens": 5,
                    "total_tokens": 15,
                },
            },
        )

    provider = OpenAIToolProvider(
        WebConfig(
            tmp_path,
            ai_enabled=True,
            ai_base_url="http://offline-tools.test/v1",
            ai_model="tools-test",
        )
    )
    original = provider.client._client
    import openai

    provider.client._client = openai.OpenAI(
        api_key="offline",
        base_url="http://offline-tools.test/v1",
        http_client=httpx.Client(transport=httpx.MockTransport(respond)),
        max_retries=0,
    )
    try:
        response = provider.complete(
            [{"role": "user", "content": "question"}],
            TOOL_SCHEMAS,
            tokens=200,
            seconds=2,
        )
        assert response.calls[0]["function"]["name"] == "submit_answer"
        assert response.usage.requests == 2
        assert "max_tokens" in sent[0] and sent[1]["max_completion_tokens"] == 200
        assert sent[1]["tools"] == TOOL_SCHEMAS
        assert (
            provider.client.request_parameters["token_parameter"]
            == "max_completion_tokens"
        )
    finally:
        provider.close()
        original.close()


def sdk_http():
    import importlib

    default = pytest.importorskip("openai").DefaultHttpxClient
    return importlib.import_module(default.__mro__[1].__module__.split(".")[0])


def sdk_provider(monkeypatch, tmp_path, respond, *, endpoint="http://127.0.0.1/v1"):
    import openai

    httpx = sdk_http()

    default = openai.DefaultHttpxClient

    def client(**kwargs):
        return default(transport=httpx.MockTransport(respond), **kwargs)

    monkeypatch.setattr(openai, "DefaultHttpxClient", client)
    return OpenAIToolProvider(
        WebConfig(
            tmp_path, ai_enabled=True, ai_base_url=endpoint, ai_model="normalized-tools"
        )
    )


def sdk_response(calls):
    return {
        "id": "offline",
        "object": "chat.completion",
        "created": 0,
        "model": "normalized-tools",
        "choices": [
            {
                "index": 0,
                "finish_reason": "tool_calls",
                "message": {"role": "assistant", "tool_calls": calls},
            }
        ],
        "usage": {"prompt_tokens": 10, "completion_tokens": 10, "total_tokens": 20},
    }


def test_provider_diagnostics_count_content_without_retaining_it(monkeypatch, tmp_path):
    httpx = sdk_http()

    def respond(request):
        data = sdk_response([])
        data["choices"][0]["finish_reason"] = "length"
        data["choices"][0]["message"].update(
            content="SECRET_REPLY", reasoning="SECRET_THOUGHT"
        )
        return httpx.Response(200, json=data)

    with_provider = sdk_provider(monkeypatch, tmp_path, respond)
    try:
        response = with_provider.complete(
            [{"role": "user", "content": "question"}],
            TOOL_SCHEMAS,
            tokens=128,
            seconds=2,
        )
        assert response.finish_reason == "length"
        assert response.diagnostics == {
            "contentBytes": 12,
            "reasoningBytes": 14,
            "httpStatus": 200,
        }
        assert "SECRET" not in json.dumps(response.diagnostics)
    finally:
        with_provider.close()


def test_ollama_metadata_is_normalized_and_echoed_without_weakening_loop(
    monkeypatch, tmp_path
):
    from types import SimpleNamespace

    httpx = sdk_http()
    from test_ask import call

    from linexcel.ask import AskLimits, QuestionRun

    sent = []

    def respond(request):
        sent.append(json.loads(request.content))
        item = (
            call("list_workbooks", {}, "first")
            if len(sent) == 1
            else call("submit_answer", {"status": "not_found"}, "end")
        )
        item["index"] = 0
        item["function"]["vendor_metadata"] = "discard"
        return httpx.Response(200, json=sdk_response([item]))

    provider = sdk_provider(monkeypatch, tmp_path, respond)
    tools = SimpleNamespace(
        deadline=__import__("time").monotonic() + 10,
        scope={},
        issued=set(),
        execute=lambda name, args: {
            "data": {"status": "not_found"}
            if name == "submit_answer"
            else {"revisions": []}
        },
    )
    try:
        result = QuestionRun(provider, tools, "Find a cell", AskLimits()).run()
        assert result["status"] == "not_found"
        assert len(sent) == 2
        echoed = next(m for m in sent[1]["messages"] if m["role"] == "assistant")[
            "tool_calls"
        ][0]
        assert set(echoed) == {"id", "type", "function"}
        assert set(echoed["function"]) == {"name", "arguments"}
    finally:
        provider.close()


@pytest.mark.parametrize(
    "endpoint", ["http://127.0.0.1/v1", "https://approved.example/v1"]
)
def test_question_loopback_redirect_cannot_forward_to_remote(
    monkeypatch, tmp_path, endpoint
):
    httpx = sdk_http()

    from linexcel.ask_provider import EndpointRedirectRefusedError
    from linexcel.web_ask import local_endpoint

    received = []
    remote = "https://unconsented.example.test/v1/chat/completions"
    assert not local_endpoint(remote)

    def respond(request):
        received.append((str(request.url), request.content))
        return httpx.Response(307, headers={"Location": remote})

    provider = sdk_provider(monkeypatch, tmp_path, respond, endpoint=endpoint)
    try:
        assert provider.client._client._client.follow_redirects is False
        with pytest.raises(EndpointRedirectRefusedError):
            provider.complete(
                [{"role": "user", "content": "SECRET_CELL_VALUE"}],
                TOOL_SCHEMAS,
                tokens=128,
                seconds=2,
            )
        assert len(received) == 1 and received[0][0].startswith(
            endpoint.removesuffix("v1")
        )
        assert b"SECRET_CELL_VALUE" in received[0][1]
        assert all(url != remote for url, _ in received)
    finally:
        provider.close()


@pytest.mark.parametrize("vision", [False, True])
@pytest.mark.parametrize(
    "endpoint", ["http://127.0.0.1/v1", "https://approved.example/v1"]
)
def test_shared_provider_refuses_documentation_and_vision_redirects(
    monkeypatch, endpoint, vision
):
    import openai

    from linexcel.aidoc import AiDocError, OpenAICompatProvider

    httpx = sdk_http()
    received = []

    def respond(request):
        received.append((str(request.url), request.content))
        return httpx.Response(
            307, headers={"Location": "https://unapproved.example/v1/chat/completions"}
        )

    default = openai.DefaultHttpxClient
    monkeypatch.setattr(
        openai,
        "DefaultHttpxClient",
        lambda **kw: default(transport=httpx.MockTransport(respond), **kw),
    )
    with OpenAICompatProvider(base_url=endpoint, model="offline") as provider:
        with pytest.raises(AiDocError, match="Endpoint redirect refused"):
            provider._complete(
                [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": "SECRET_CELL_VALUE"},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": "data:image/png;base64,SECRET_IMAGE"
                                },
                            },
                        ]
                        if vision
                        else "SECRET_CELL_VALUE",
                    }
                ],
                "SECRET_CELL_VALUE",
                {"max_tokens": 128},
                vision=vision,
            )
    assert len(received) == 1
    assert received[0][0].startswith(endpoint)
    assert b"SECRET_CELL_VALUE" in received[0][1]
    assert (b"SECRET_IMAGE" in received[0][1]) == vision
