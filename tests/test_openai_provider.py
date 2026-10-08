"""Exercise real SDK requests through HTTPX without an external API."""

import json

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
