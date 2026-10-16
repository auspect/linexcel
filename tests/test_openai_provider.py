"""Exercise real SDK requests through HTTPX without an external API."""

import json
from types import SimpleNamespace

import pytest

from linexcel.aidoc import AiDocError, OpenAICompatProvider, _resolve_provider

httpx = pytest.importorskip("httpx")
openai = pytest.importorskip("openai")


def completion(request):
    return httpx.Response(
        200,
        json={
            "id": "test",
            "object": "chat.completion",
            "created": 0,
            "model": "test-model",
            "choices": [
                {
                    "index": 0,
                    "message": {"role": "assistant", "content": "# Result"},
                    "finish_reason": "stop",
                }
            ],
            "usage": {"prompt_tokens": 7, "completion_tokens": 3, "total_tokens": 10},
        },
    )


@pytest.mark.parametrize(
    "base,model,url",
    [
        (
            "https://{model}.gateway.example/v1",
            "model-a",
            "https://model-a.gateway.example/v1/chat/completions",
        ),
        (
            "https://gateway.example/{model}/v1",
            "org/model-a",
            "https://gateway.example/org%2Fmodel-a/v1/chat/completions",
        ),
        (
            "https://model-a.gateway.example/v1",
            "body-model",
            "https://model-a.gateway.example/v1/chat/completions",
        ),
    ],
)
def test_model_url_and_json_body_are_independent(base, model, url):
    seen = []

    def respond(request):
        seen.append(request)
        return completion(request)

    with httpx.Client(transport=httpx.MockTransport(respond)) as http:
        with OpenAICompatProvider(
            model=model, base_url=base, http_client=http
        ) as provider:
            text, usage = provider.generate_with_usage("system", "user", max_tokens=42)
            assert text == "# Result"
            assert usage.total == 10 and not usage.estimated
            assert _resolve_provider(provider=provider) is provider
        assert not http.is_closed
    assert str(seen[0].url) == url
    body = json.loads(seen[0].content)
    assert body["model"] == model
    assert body["max_tokens"] == 42
    assert "stream" not in body


def test_http_sdk_and_request_options_reach_text_and_vision():
    seen = []

    def respond(request):
        seen.append(request)
        return completion(request)

    with httpx.Client(transport=httpx.MockTransport(respond)) as http:
        with OpenAICompatProvider(
            model="model-a",
            base_url="https://{model}.example/v1",
            api_key="secret",
            http_client=http,
            client_kwargs={
                "timeout": httpx.Timeout(12, connect=2),
                "max_retries": 0,
                "default_headers": {"X-Org": "team"},
                "default_query": {"version": "test"},
            },
            request_kwargs={
                "temperature": 0.7,
                "top_p": 0.9,
                "extra_body": {"custom": True},
                "extra_headers": {"X-Route": "custom"},
            },
        ) as provider:
            provider.generate("system", "user")
            provider.generate_with_image("system", "user", b"test-image", max_tokens=8)
    for request in seen:
        assert request.headers["X-Org"] == "team"
        assert request.headers["X-Route"] == "custom"
        assert request.url.params["version"] == "test"
        assert request.extensions["timeout"]["connect"] == 2
        body = json.loads(request.content)
        assert body["custom"] is True and body["temperature"] == 0.7
    body = json.loads(seen[1].content)
    assert body["max_tokens"] == 8
    assert body["messages"][1]["content"][1]["image_url"]["url"].startswith(
        "data:image/png;base64,"
    )


def test_proxy_options_are_forwarded_and_owned_client_closed(monkeypatch):
    captured = {}
    original = openai.DefaultHttpxClient

    def build_http(**kwargs):
        captured.update(kwargs)
        return original(**kwargs)

    monkeypatch.setattr(openai, "DefaultHttpxClient", build_http)
    with OpenAICompatProvider(
        model="test",
        base_url="https://test.example/v1",
        proxy="http://proxy.example:8080",
        http_client_kwargs={"trust_env": False, "follow_redirects": False},
        client_kwargs={"timeout": 20},
    ) as provider:
        sdk = provider._client
        assert not sdk.is_closed()
    assert sdk.is_closed()
    assert captured["proxy"] == "http://proxy.example:8080"
    assert captured["trust_env"] is False


def test_preconfigured_openai_client_remains_caller_owned():
    with httpx.Client(transport=httpx.MockTransport(completion)) as http:
        with openai.OpenAI(
            api_key="test", base_url="https://test.example/v1", http_client=http
        ) as sdk:
            with OpenAICompatProvider(model="test", client=sdk) as provider:
                assert provider.generate("system", "user") == "# Result"
            assert not sdk.is_closed()


@pytest.mark.parametrize(
    "options",
    [
        {"proxy": "http://proxy", "http_client": object()},
        {"proxy": "http://proxy", "http_client_kwargs": {"proxy": "http://other"}},
        {"client": object(), "base_url": "http://test"},
        {"client_kwargs": {"base_url": "http://other"}},
        {"request_kwargs": {"model": "other"}},
        {"request_kwargs": {"extra_body": {"max_tokens": 999999}}},
        {"request_kwargs": {"stream": True}},
    ],
)
def test_ambiguous_or_reserved_options_fail_before_requests(options):
    params = {"model": "test", "base_url": "https://test.example/v1", **options}
    with pytest.raises(ValueError):
        OpenAICompatProvider(**params)


@pytest.mark.parametrize(
    "base,model",
    [
        ("https://{model}.example/v1", "org/model"),
        ("https://{model}.example/v1", "model@evil"),
        ("file:///tmp/test", "model"),
        ("https://test.example/{unknown}", "model"),
        ("https://user:password@test.example/v1", "model"),
        ("https://test.example/v1?secret=abc", "model"),
    ],
)
def test_invalid_urls_fail_before_client_creation(base, model):
    with pytest.raises(ValueError):
        OpenAICompatProvider(model=model, base_url=base)


def test_provider_error_does_not_quote_server_payload_or_url():
    def respond(request):
        return httpx.Response(401, json={"error": {"message": "secret-from-gateway"}})

    with httpx.Client(transport=httpx.MockTransport(respond)) as http:
        with OpenAICompatProvider(
            model="test", base_url="https://private.example/v1", http_client=http
        ) as provider:
            with pytest.raises(AiDocError) as caught:
                provider.generate("system", "user")
    assert "AuthenticationError" in str(caught.value)
    assert "HTTP 401" in str(caught.value)
    assert "secret-from-gateway" not in str(caught.value)
    assert "private.example" not in str(caught.value)


def test_vision_billing_failure_is_reported_without_payload_or_false_cause():
    def respond(request):
        return httpx.Response(
            402, json={"error": {"message": "private billing details"}}
        )

    with httpx.Client(transport=httpx.MockTransport(respond)) as http:
        with OpenAICompatProvider(
            model="test", base_url="https://private.example/v1", http_client=http
        ) as provider:
            with pytest.raises(AiDocError) as caught:
                provider.generate_with_image("system", "user", b"image")
    message = str(caught.value)
    assert "HTTP 402" in message
    assert "private billing details" not in message
    assert "private.example" not in message
    assert "text-only" not in message


def sse_chunk(content=None, reason=None, *, usage=None, reasoning=None):
    return {
        "id": "stream-test",
        "object": "chat.completion.chunk",
        "created": 0,
        "model": "resolved-model",
        "usage": usage,
        "choices": []
        if usage
        else [
            {
                "index": 0,
                "delta": {"content": content, "reasoning_content": reasoning},
                "finish_reason": reason,
            }
        ],
    }


@pytest.mark.parametrize("vision", [False, True])
def test_real_sdk_stream_assembles_only_content_and_terminal_usage(vision):
    responses, requests, receipts = [], [], []

    def respond(request):
        requests.append(request)
        chunks = [
            sse_chunk(reasoning="Private thinking is not documentation."),
            sse_chunk("  # "),
            sse_chunk("Result  ", "stop"),
            sse_chunk(
                usage={"prompt_tokens": 7, "completion_tokens": 3, "total_tokens": 10}
            ),
        ]
        body = (
            "".join("data: " + json.dumps(c) + "\n\n" for c in chunks)
            + "data: [DONE]\n\n"
        )
        response = httpx.Response(
            200, headers={"content-type": "text/event-stream"}, content=body
        )
        responses.append(response)
        return response

    class RecordingProvider(OpenAICompatProvider):
        def _checked_response(self, response, prompt, text, *, require_stop=False):
            assert responses[0].is_closed
            receipts.append((response.model, response.id, require_stop))
            return super()._checked_response(
                response, prompt, text, require_stop=require_stop
            )

    with httpx.Client(transport=httpx.MockTransport(respond)) as http:
        with RecordingProvider(
            model="requested-model",
            base_url="https://test.example/v1",
            http_client=http,
            streaming=True,
            request_kwargs={
                "extra_body": {
                    "reasoning_split": True,
                    "stream_options": {"include_usage": False},
                }
            },
        ) as provider:
            if vision:
                text, usage = provider.generate_with_image(
                    "system", "user", b"image", max_tokens=42
                )
            else:
                text, usage = provider.generate_with_usage(
                    "system", "user", max_tokens=42
                )
    assert text == "# Result"
    assert usage.total == 10 and not usage.estimated
    assert receipts == [("resolved-model", "stream-test", True)]
    assert responses[0].is_closed
    body = json.loads(requests[0].content)
    assert body["stream"] is True
    assert body["stream_options"]["include_usage"] is True
    assert body["max_tokens"] == 42 and body["reasoning_split"] is True
    if vision:
        assert body["messages"][1]["content"][1]["image_url"]["url"].startswith(
            "data:image/png;base64,"
        )


class FakeStream:
    def __init__(self, chunks, error=None, close_error=None):
        self.chunks, self.error, self.close_error = chunks, error, close_error
        self.closed = False

    def __iter__(self):
        yield from self.chunks
        if self.error:
            raise self.error

    def close(self):
        self.closed = True
        if self.close_error:
            raise self.close_error


def fake_chunk(
    content=None,
    reason=None,
    *,
    index=0,
    usage=None,
    model="resolved",
    response_id="response",
):
    choices = (
        []
        if usage
        else [
            SimpleNamespace(
                index=index,
                delta=SimpleNamespace(content=content, reasoning_content="ignored"),
                finish_reason=reason,
            )
        ]
    )
    return SimpleNamespace(choices=choices, usage=usage, model=model, id=response_id)


def provider_for_stream(stream):
    calls = []

    def create(**kwargs):
        calls.append(kwargs)
        return stream

    client = SimpleNamespace(
        chat=SimpleNamespace(completions=SimpleNamespace(create=create))
    )
    return OpenAICompatProvider(model="test", client=client, streaming=True), calls


@pytest.mark.parametrize("reason", [None, "length", "content_filter", "tool_calls"])
def test_stream_requires_stop_even_with_nonempty_text_and_preserves_usage(reason):
    reported = SimpleNamespace(prompt_tokens=7, completion_tokens=3)
    stream = FakeStream([fake_chunk("Partial", reason), fake_chunk(usage=reported)])
    provider, _ = provider_for_stream(stream)
    with pytest.raises(AiDocError) as caught:
        provider.generate_with_usage("system", "user")
    assert stream.closed
    assert caught.value.usage.total == 10
    assert not caught.value.usage.estimated


def test_stream_empty_content_never_promotes_reasoning_to_documentation():
    stream = FakeStream([fake_chunk(reason="stop")])
    provider, _ = provider_for_stream(stream)
    with pytest.raises(AiDocError, match="empty response"):
        provider.generate("system", "user")
    assert stream.closed


def test_stream_without_reported_usage_is_explicitly_estimated():
    stream = FakeStream([fake_chunk("Completed response", "stop")])
    provider, _ = provider_for_stream(stream)
    text, usage = provider.generate_with_usage("system", "user")
    assert text == "Completed response"
    assert usage.estimated and usage.requests == 1
    assert stream.closed


@pytest.mark.parametrize("close_error", [None, RuntimeError("private close details")])
def test_interrupted_stream_is_not_retried_and_always_closed(close_error):
    stream = FakeStream(
        [fake_chunk("Partial")],
        error=RuntimeError("private gateway details"),
        close_error=close_error,
    )
    provider, calls = provider_for_stream(stream)
    with pytest.raises(AiDocError) as caught:
        provider.generate("system", "user")
    assert len(calls) == 1 and stream.closed
    assert "private" not in str(caught.value)
    assert caught.value.usage.estimated
    assert caught.value.__cause__ is stream.error
    if close_error:
        assert caught.value.__notes__ == ["Stream close also failed (RuntimeError)"]


def test_close_failure_cannot_first_be_recorded_as_a_complete_response(monkeypatch):
    usage = SimpleNamespace(prompt_tokens=7, completion_tokens=3)
    stream = FakeStream(
        [fake_chunk("Complete text", "stop"), fake_chunk(usage=usage)],
        close_error=RuntimeError("private close details"),
    )
    provider, _ = provider_for_stream(stream)
    receipts = []
    monkeypatch.setattr(
        provider, "_checked_response", lambda *args, **kwargs: receipts.append(args)
    )
    with pytest.raises(AiDocError, match="stream close failed") as caught:
        provider.generate("system", "user")
    assert receipts == [] and stream.closed
    assert caught.value.usage.total == 10


@pytest.mark.parametrize(
    "invalid",
    [
        "multiple_choices",
        "wrong_index",
        "non_text",
        "after_terminal",
        "model_changed",
        "id_changed",
    ],
)
def test_invalid_stream_cannot_mix_choices_or_lose_reported_usage(invalid):
    reported = SimpleNamespace(prompt_tokens=7, completion_tokens=3)
    bad = fake_chunk("Unexpected", "stop")
    if invalid == "multiple_choices":
        bad.choices.append(bad.choices[0])
    elif invalid == "wrong_index":
        bad.choices[0].index = 1
    elif invalid == "non_text":
        bad.choices[0].delta.content = ["unexpected"]
    elif invalid == "model_changed":
        bad.model = "different-model"
    elif invalid == "id_changed":
        bad.id = "different-id"
    chunks = [
        fake_chunk("Valid prefix", "stop" if invalid == "after_terminal" else None),
        fake_chunk(usage=reported),
        bad,
    ]
    stream = FakeStream(chunks)
    provider, _ = provider_for_stream(stream)
    with pytest.raises(AiDocError) as caught:
        provider.generate("system", "user")
    assert caught.value.usage.total == 10 and not caught.value.usage.estimated
    assert stream.closed


@pytest.mark.parametrize(
    "request_options",
    [
        {"n": 2},
        {"n": True},
        {"extra_body": {"n": 0}},
        {"stream": False},
        {"extra_body": {"stream": True}},
    ],
)
def test_streaming_options_are_explicit_and_single_choice(request_options):
    with pytest.raises(ValueError):
        OpenAICompatProvider(
            model="test",
            client=object(),
            streaming=True,
            request_kwargs=request_options,
        )


def test_native_multiimage_sdk_request_and_receipt_bind_exact_prompt():
    import base64
    import hashlib

    seen = []

    def respond(request):
        seen.append(request)
        return completion(request)

    images = [
        ("r0-c0", b"first bytes", "image/png"),
        ("r0-c1", b"second bytes", "image/jpeg"),
    ]
    with httpx.Client(transport=httpx.MockTransport(respond)) as http:
        with OpenAICompatProvider(
            model="requested", base_url="https://test.invalid/v1", http_client=http
        ) as provider:
            provider.generate_with_images("system", "json prompt", images)
            receipt = provider.vision_receipt
            assert receipt["response_id"] == "test"
            assert receipt["resolved_model"] == "test-model"
            assert receipt["status"] == "complete"
            assert receipt["finish_reason"] == "stop"
            assert (
                receipt["prompt_text_sha256"]
                == hashlib.sha256(b"system\n\njson prompt").hexdigest()
            )
            assert receipt["content_sha256"] == hashlib.sha256(b"# Result").hexdigest()
    content = json.loads(seen[0].content)["messages"][1]["content"]
    assert content[0] == {"type": "text", "text": "json prompt"}
    for index, (label, payload, media) in enumerate(images):
        assert content[1 + 2 * index] == {"type": "text", "text": f"Region: {label}"}
        image = content[2 + 2 * index]["image_url"]
        assert image["detail"] == "high"
        assert image["url"] == (
            f"data:{media};base64," + base64.b64encode(payload).decode()
        )


def test_incomplete_multiimage_stream_retains_failed_vision_receipt():
    stream = FakeStream(
        [
            fake_chunk("Partial JSON", "length"),
            fake_chunk(usage=SimpleNamespace(prompt_tokens=7, completion_tokens=3)),
        ]
    )
    provider, _ = provider_for_stream(stream)
    with pytest.raises(AiDocError):
        provider.generate_with_images(
            "system", "regions", [("r0-c0", b"x", "image/png")]
        )
    receipt = provider.vision_receipt
    assert receipt["status"] == "failed"
    assert receipt["finish_reason"] == "length"
    assert receipt["streaming"] is True
    assert receipt["usage"]["input_tokens"] == 7
