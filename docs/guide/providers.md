# Choosing an AI provider

AI documentation is optional and vendor-neutral. No provider is named in the
code and none is chosen for you: a call without `base_url=` or `provider=`
raises `AiDocError` listing the options, and nothing is sent anywhere until you
pick one.

There are two ways in.

| Way | What it covers |
| --- | --- |
| `base_url=` + `model=` | Anything speaking the OpenAI-compatible chat API — a local runtime, a hosted gateway, a vendor's own endpoint |
| `provider=` | Your own callable or `LLMProvider` object, for an API that speaks something else |

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
        model="llama3.1",
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

```bash
export LINEXCEL_AI_BASE_URL=http://localhost:11434/v1
export LINEXCEL_AI_MODEL=llama3.1
```

```python
docs = result.document()   # configured entirely by the environment
```

!!! note "A model must always be named"

    `base_url=` without a model raises rather than falling back to one.
    Endpoints do not agree on a default — `llama3.1` means nothing to a hosted
    API and a hosted model id means nothing to Ollama — so linexcel does not
    invent one.

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
