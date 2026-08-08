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

The dossier behind the overview carries two things. The lineage says how the
workbook *computes*: sheet statistics, the largest formula patterns, defined
names, VBA procedures, unresolved references. The
[workbook context](context.md) says what it *looks like*: the first rows of
each sheet, cell comments, merged ranges, frozen panes, hidden columns — the
same cues the [sheet screenshots](context.md#screenshots) show a human reader.

That second half is what lets the model write about the file rather than about
a graph. A title sitting two rows above a table, a column called `Prix` that
somebody hid, a comment reading *"exported product category"* — no formula
records any of it, and an overview written without them describes a structure
nobody recognises.

The images themselves are never uploaded: `openpyxl` reads the same facts
deterministically, so the "cite only the dossier" rule still holds and no vision
model is required.

```python
# Lineage only — cell contents stay on your machine
workbook_doc = result.document_workbook(base_url=..., include_context=False)
```

The dossier is capped at `aidoc.MAX_WORKBOOK_DOSSIER_CHARS`. A workbook that
exceeds it sheds detail in order — long previews shrink, then the tail of the
pattern and VBA lists, and only as a last resort are previews and comments
dropped — so a small workbook loses nothing.

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

## Capping the bill

`token_usage` reports what a run cost *after* it ran. `token_budget=` decides
what it is allowed to cost before it starts:

```python
docs = result.document(model="gemini-3.1-flash-lite", token_budget=200_000)
print(result.token_usage)
```

The budget is a ceiling on the **total** tokens of the run — input and output
together, every request combined. It is deliberately not a per-node allowance:
one node's card is never the question, and a workbook with 900 formula patterns
is 900 requests whose sum is the only figure that turns up on an invoice. Use
`max_tokens=` when you want to bound an individual response.

Because a request's cost is only known once it has answered, the budget is
enforced *between* requests. Nodes still queued when the tally reaches the
ceiling are never sent, the cards already written are returned, and a
`UserWarning` names how many nodes were left undocumented:

```text
UserWarning: Token budget of 200,000 tokens reached after 201,447:
612 of 900 nodes documented, 288 never sent. Raise token_budget to document the rest.
```

Requests already in flight are allowed to finish, so the final tally can exceed
the ceiling by up to `max_workers` responses — set a budget as an order of
magnitude, not to the token.

The ceiling is counted against `result.token_usage`, which spans the result's
lifetime. One budget covers `document_workbook()` and `document()` together, so
the figure to choose is what the whole workbook is worth to you:

```python
result.document_workbook(model="gemini-3.1-flash-lite", token_budget=200_000)
result.document(model="gemini-3.1-flash-lite", token_budget=200_000)  # same ceiling
```

A budget already spent by earlier calls raises `AiDocError` rather than sending
a request that would breach it. `token_budget=0` — or any non-positive value —
raises `ValueError`; to send nothing at all, do not call `document()`.

!!! tip "Estimating before you spend"

    A dry run costs nothing: point `provider=` at a callable that records its
    prompt and returns `""`. `result.token_usage` then holds the estimated input
    cost of the whole workbook, which is the bulk of the bill for
    documentation-shaped work.

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

Nothing is sent until a provider is chosen explicitly (there is no default):
Google Gemini when you pass `model=` with a key, the endpoint behind `base_url`
otherwise (a local Ollama or vLLM runtime keeps everything on your machine), or
wherever your own callable sends it.

What each call sends:

| Call | Sent |
| --- | --- |
| `document()` | Per node: formula, computed value, precedent/dependent labels, formula decomposition, stretched-group extent, and extracted VBA code |
| `document_workbook()` | Sheet statistics, the largest formula patterns, defined names, VBA procedures, unresolved references, analysis warnings |
| `document_workbook()` with `include_context=True` *(default)* | **Also** the first rows and columns of every sheet, cell comments and their authors, merged ranges, frozen panes, hidden columns |

That last row means cell *contents* leave the machine, not only formulas. It is
on by default because an overview written without them is not worth reading —
but if a workbook holds data that must stay local and your provider is remote,
pass `include_context=False`, or keep the whole run local with `base_url=`.
