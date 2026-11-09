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
