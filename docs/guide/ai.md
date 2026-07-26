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
