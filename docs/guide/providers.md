# Choosing an AI provider

AI documentation is optional and vendor-neutral. No provider is named in the
code and none is chosen for you: a call without `base_url=` or `provider=`
raises `AiDocError` listing the options, and nothing is sent anywhere until you
pick one.

There are two ways in.

| Way | What it covers |
| --- | --- |
| `base_url=` + `model=` | Anything speaking the OpenAI-compatible chat API — a local runtime, a hosted gateway, a vendor's own endpoint |
| `provider=` | A configured `OpenAICompatProvider`, your own callable, or any `LLMProvider` object |

Install the client once — it is the client for *every* OpenAI-compatible
endpoint, not a choice of vendor:

```bash
uv add linexcel[ai]      # or: pip install "linexcel[ai]"
```

## OpenAI-compatible endpoints

Same two arguments every time; only the URL and the model name change.

=== "Ollama (100% local)"

    Nothing leaves the machine and nothing is billed. The workbook, its values
    and its comments stay on your disk.

    ```python
    docs = result.document(
        base_url="http://localhost:11434/v1",
        model="qwen3.8",
    )
    ```

    No API key is needed; Ollama ignores the one the client sends.

=== "OpenRouter"

    One endpoint, many models — useful for comparing several against the same
    workbook without changing any code but the model string.

    ```python
    docs = result.document(
        base_url="https://openrouter.ai/api/v1",
        model="<vendor>/<model>",
        api_key="...",              # or set LINEXCEL_AI_API_KEY
    )
    ```

=== "vLLM / LM Studio"

    A self-hosted server, local or on your own infrastructure.

    ```python
    docs = result.document(
        base_url="http://localhost:8000/v1",
        model="<the model you served>",
    )
    ```

=== "A vendor's own endpoint"

    Most vendors expose an OpenAI-compatible URL alongside their native API;
    point `base_url=` at it and pass your key.

    ```python
    docs = result.document(
        base_url="https://api.<vendor>.example/v1",
        model="<their model id>",
        api_key="...",
    )
    ```

### Environment variables

Each argument has an environment equivalent, so a provider can be configured
once outside the code:

| Variable | Argument | Notes |
| --- | --- | --- |
| `LINEXCEL_AI_BASE_URL` | `base_url=` | `OPENAI_BASE_URL` is also read |
| `LINEXCEL_AI_MODEL` | `model=` | `OPENAI_MODEL` is also read |
| `LINEXCEL_AI_API_KEY` | `api_key=` | `OPENAI_API_KEY` is also read |
| `LINEXCEL_AI_TIMEOUT_SECONDS` | `client_kwargs={"timeout": ...}` on `OpenAICompatProvider` | Default 300 seconds; explicit SDK timeout takes precedence |

```bash
export LINEXCEL_AI_BASE_URL=http://localhost:11434/v1
export LINEXCEL_AI_MODEL=qwen3.8
```

```python
docs = result.document()   # configured entirely by the environment
```

!!! note "A model must always be named"

    `base_url=` without a model raises rather than falling back to one.
    Endpoints do not agree on a default — `qwen3.8` means nothing to a hosted
    API and a hosted model id means nothing to Ollama — so linexcel does not
    invent one.

## Model in the hostname, proxies and HTTPX

The SDK treats the endpoint and model as independent values. For a gateway at
`https://model-a.gateway.example/v1`, use that address as `base_url` and its
expected model ID as `model`. The SDK posts to `/v1/chat/completions` and includes
the model in the JSON body. The server must support this chat API.

`OpenAICompatProvider` also accepts a literal `{model}` placeholder in the URL:

```python
import os
from linexcel.aidoc import OpenAICompatProvider

with OpenAICompatProvider(
    model="model-a",
    base_url="https://{model}.gateway.example/v1",
    api_key=os.environ["LINEXCEL_AI_API_KEY"],
    proxy="http://proxy.example:8080",  # omit when unnecessary
    http_client_kwargs={"trust_env": False, "follow_redirects": False},
    client_kwargs={
        "timeout": 60.0,
        "max_retries": 0,
        "default_headers": {"X-Team": "finance"},
        "default_query": {"api-version": "2026-01-01"},
    },
    request_kwargs={"extra_body": {"custom_option": True}},
) as provider:
    docs = result.document(provider=provider, max_tokens=1_000)
    overview = result.document_workbook(provider=provider)
    # result.describe_screenshots(shots, provider=provider) for a vision model
```

This expands to `https://model-a.gateway.example/v1/chat/completions`.
Hostnames require a DNS-compatible model ID. For an ID such as `org/model-a`,
use a fixed hostname, or `https://gateway.example/{model}/v1`: a path placeholder
is percent-encoded as `org%2Fmodel-a`. A full URL passed to `base_url` must stop
before `/chat/completions`. Query parameters belong in `default_query`, and
credentials belong in `api_key` or headers.

The same placeholder works with `result.document(base_url=..., model=...)` and
`LINEXCEL_AI_BASE_URL`. Advanced options are configured through `provider=` in
the Python API; the command line and browser settings do not expose them.

| Argument | Configures |
| --- | --- |
| `proxy` | Explicit HTTP proxy; omit to use transport defaults |
| `http_client_kwargs` | Default SDK transport: TLS `verify`, certificate context, connection `limits`, `transport`, `mounts`, `event_hooks`, `trust_env`, redirects |
| `http_client` | An existing synchronous `httpx.Client`, or the HTTPX2 client supported by your SDK |
| `client_kwargs` | OpenAI constructor options such as timeout, retries, headers, query, organization and project |
| `request_kwargs` | Chat completion options, including `top_p`, `temperature`, `extra_headers`, `extra_query`, `extra_body` |
| `client` | An already configured synchronous `OpenAI` instance, used with `model` and optional `request_kwargs` |

The locked OpenAI SDK 3.13.0 defaults to HTTPX2. Older supported SDKs use HTTPX;
the minimum supported SDK is 1.55.3.
linexcel uses the SDK's `DefaultHttpxClient` for `http_client_kwargs`; pass
objects such as transports or timeout instances from the matching HTTP library.
The AI extra also installs HTTPX for explicit injection:

```python
import httpx
from linexcel.aidoc import OpenAICompatProvider

with httpx.Client(
    proxy="http://proxy.example:8080",
    verify=True,
    trust_env=False,
    follow_redirects=False,
) as http:
    with OpenAICompatProvider(
        model="model-a",
        base_url="https://{model}.gateway.example/v1",
        http_client=http,
        client_kwargs={"timeout": httpx.Timeout(60, connect=5)},
    ) as provider:
        docs = result.document(provider=provider)
```

Injected HTTP and SDK clients remain caller-owned. Close them after all requests
finish. The provider context manager closes connections it creates itself.
Do not combine an injected HTTP client with `proxy` or `http_client_kwargs`;
configure those on the injected client. Do not combine `client` with endpoint,
credential or SDK configuration.

Requests are non-streaming. `request_kwargs` and `extra_body` cannot override
`model`, `messages` or `max_tokens`. Set the output ceiling through the
documentation method. linexcel defaults to zero SDK retries; enabling retries
can increase request count and duration outside the token tally.

Provider errors report the HTTP status without quoting the server payload.
For a hosted gateway, an unbounded request can be refused because its potential
output exceeds available credit. Pass `max_tokens=` to the documentation method
to bound that request. A truncated response still fails validation; inspect the
finish reason before lowering the ceiling further.

For OpenRouter, the model's web page is separate from the API endpoint. This
configuration was tested with text and images:

```python
with OpenAICompatProvider(
    base_url="https://openrouter.ai/api/v1",
    model="qwen/qwen3.8-flash",
    api_key=os.environ["OPENROUTER_API_KEY"],
    request_kwargs={"extra_body": {"reasoning": {"enabled": False}}},
) as provider:
    docs = result.document(provider=provider, max_tokens=2_048, max_workers=1)
```

The output ceiling and concurrency should suit your workbook and account.
OpenRouter distinguishes account credits, key spending caps and reservations
for concurrent requests; see its [credit limits documentation](https://openrouter.ai/docs/api_reference/limits).

See the [official OpenAI Python reference](https://developers.openai.com/api/reference/python)
and [HTTPX proxy guide](https://www.python-httpx.org/advanced/proxies/) for
transport details. Options still depend on the SDK and gateway versions.

## Custom provider

Anything else: a native SDK, an internal gateway, a queue, a stub for testing.
Any callable with this signature works, and so does any object exposing a
`generate` method with the same one:

```python
def my_llm(system_prompt: str, user_prompt: str, *, temperature: float = 0.2) -> str:
    # call whatever you like here
    return response_text


docs = result.document(provider=my_llm)
```

A provider may optionally report what each call consumed, so the token tally
comes from the API rather than an approximation. Implement `generate_with_usage`
alongside `generate`:

```python
from linexcel.aidoc import TokenUsage


class MyProvider:
    def generate(self, system_prompt, user_prompt, *, temperature=0.2, max_tokens=None):
        return self.generate_with_usage(
            system_prompt, user_prompt, temperature=temperature, max_tokens=max_tokens
        )[0]

    def generate_with_usage(
        self, system_prompt, user_prompt, *, temperature=0.2, max_tokens=None
    ):
        response = my_sdk.complete(...)
        return response.text, TokenUsage(
            input_tokens=response.usage.input,
            output_tokens=response.usage.output,
            requests=1,
            model="<model id>",
            provider="<label of your choosing>",
        )
```

Without it, tokens are estimated and [`token_usage.estimated`](ai.md#token-usage)
is `True`.

## Where the data goes

| Configuration | Destination |
| --- | --- |
| Nothing configured | Nowhere — `AiDocError` is raised and the message lists the options |
| `base_url=` pointing at a local runtime | Your own machine |
| `base_url=` pointing at a hosted endpoint | That endpoint's operator, under their terms |
| `provider=` | Wherever your callable sends it |

[Data handling](data-handling.md) details exactly what each call puts in the
payload.
