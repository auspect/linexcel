# AI documentation (optional, multi-provider)

AI documentation is opt-in and supports any LLM provider.

**There is no default provider.** A call without `provider=`, `base_url=` or
`model=` raises `AiDocError` and lists the options; nothing is sent anywhere
until you pick one. Three ways to choose:

| Way | Provider |
| --- | --- |
| `model=` (or `GEMINI_MODEL`) + Google key | Google Gemini |
| `base_url=` (or `LINEXCEL_AI_BASE_URL`) | Any OpenAI-compatible endpoint (Ollama, vLLM, LM Studio, OpenAI…) |
| `provider=` | Your own callable / `LLMProvider` instance |

## Google Gemini (explicit opt-in)

Gemini requires naming a model — passing only `api_key=` no longer selects it:

```python
docs = result.document(model="gemini-3.1-flash-lite", api_key="...", language="en")
result.save_html("out.html", docs=docs, language="en")
# or: export GOOGLE_API_KEY=... and pass only model=
```

Requires `google-genai` (`uv add linexcel[ai]`).

## OpenAI-compatible (Ollama, vLLM, LM Studio, OpenAI)

```python
docs = result.document(
    base_url="http://localhost:11434/v1",
    model="llama3.1",
    language="en",
)
```

Requires `openai` (`uv add linexcel[openai]`).

## Custom provider

Any callable with this signature works, and so does any object exposing a
`generate` method with the same one:

```python
def my_llm(system_prompt: str, user_prompt: str, *, temperature: float = 0.2) -> str:
    # call your model here
    return response_text


docs = result.document(provider=my_llm)
```

## Workbook overview

```python
workbook_doc = result.document_workbook(model="gemini-3.1-flash-lite", language="en")
result.save_html("out.html", docs=docs, workbook_doc=workbook_doc, language="en")
```

## Concurrency and partial failures

`document()` issues `max_workers` requests in parallel (4 by default):

```python
docs = result.document(model="gemini-3.1-flash-lite", max_workers=8)
```

Documenting a large workbook is long and often billed, so a node that fails does
not discard the ones that succeeded. The successful cards are returned and a
`UserWarning` reports how many nodes were dropped; `AiDocError` is raised only
when every node failed.

## Token usage

Every AI call is tallied on the result:

```python
docs = result.document(model="gemini-3.1-flash-lite")  # Gemini opt-in; use base_url= or provider= for other providers
overview = result.document_workbook(model="gemini-3.1-flash-lite")

print(result.token_usage)
# 48,210 tokens (44,900 in + 3,310 out) over 5 request(s) [gemini/gemini-3.1-flash-lite]

usage = result.token_usage
usage.input_tokens, usage.output_tokens, usage.total, usage.requests
```

Counts come from the provider when it reports them — Gemini's `usage_metadata`
and the OpenAI-compatible `usage` block are read directly, so the figure is the
one you are billed on. When a provider reports nothing (a custom callable, or a
local runtime that omits the block), the tokens are approximated instead and
`usage.estimated` is `True`; `str(usage)` then prefixes the numbers with `~`.

```python
if result.token_usage.estimated:
    print("approximate — provider did not report usage")
```

Tokens already spent are counted even when a later node fails, since they are
billed regardless. The tally accumulates across every call made on the result;
build a fresh `TokenUsage` and pass it as `usage=` to
`aidoc.document_nodes()` if you need to scope it more finely.

!!! warning "Tokens, not currency"

    `TokenUsage` deliberately carries no price. Rates differ per provider,
    model and region, and a table baked into the package would silently go
    stale. Multiply by your own current rate.

## Languages

`language=` drives both the system prompt sent to the model and the viewer
interface. Nine are available: `en` (default), `fr`, `es`, `de`, `it`, `pt`,
`nl`, `ja`, `zh`.

```python
docs = result.document(model="gemini-3.1-flash-lite", language="ja")
result.save_html("out.html", docs=docs, language="ja")
```

It is a closed allowlist rather than free-form text — the value selects a stored
prompt and reaches the generated JavaScript, so an arbitrary string would be a
prompt-injection and interpolation vector. Any other value raises `ValueError`.

Adding one means extending `linexcel.i18n.UI_STRINGS` and both prompt registries
in `linexcel.aidoc`; the test suite asserts the three stay in sync, so a partial
addition fails rather than surfacing as raw interface keys.

!!! note "Translation provenance"

    English and French were written by hand. The other seven languages — the
    interface strings *and* the AI system prompts — were produced with AI
    assistance and have not been reviewed by native speakers.

    This matters more for the prompts than for the interface: their wording
    steers how the model writes each card, so an awkward phrasing degrades
    output quality rather than just looking odd. Corrections are welcome.

## Data handling

The dossier sent for each node can include formulas, computed values,
precedent/dependent labels, formula decomposition, sheet structure, defined
names, and extracted VBA code. It goes to whichever provider you configure —
and nothing is sent until one is chosen explicitly (no default): Google Gemini
when you pass `model=` with a key, the endpoint behind `base_url` otherwise (a
local Ollama or vLLM runtime keeps it on your machine), or wherever your own
callable sends it.
