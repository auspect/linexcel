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

## Data handling

The dossier sent for each node can include formulas, computed values,
precedent/dependent labels, formula decomposition, sheet structure, defined
names, and extracted VBA code. It goes to whichever provider you configure —
Google Gemini by default, the endpoint behind `base_url` otherwise (a local
Ollama or vLLM runtime keeps it on your machine), or wherever your own callable
sends it.
