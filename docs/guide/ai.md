# AI documentation (optional, multi-provider)

AI documentation is opt-in and supports any LLM provider.

## Google Gemini (default)

```python
docs = result.document(api_key="...", language="en")
result.save_html("out.html", docs=docs, language="en")
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
workbook_doc = result.document_workbook(language="en")
result.save_html("out.html", docs=docs, workbook_doc=workbook_doc, language="en")
```

## Concurrency and partial failures

`document()` issues `max_workers` requests in parallel (4 by default):

```python
docs = result.document(max_workers=8)
```

Documenting a large workbook is long and often billed, so a node that fails does
not discard the ones that succeeded. The successful cards are returned and a
`UserWarning` reports how many nodes were dropped; `AiDocError` is raised only
when every node failed.

## Languages

`language=` drives both the system prompt sent to the model and the viewer
interface. Nine are available: `en` (default), `fr`, `es`, `de`, `it`, `pt`,
`nl`, `ja`, `zh`.

```python
docs = result.document(language="ja")
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
Google Gemini by default, the endpoint behind `base_url` otherwise (a local
Ollama or vLLM runtime keeps it on your machine), or wherever your own callable
sends it.
