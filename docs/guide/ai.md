# AI documentation

Optional and opt-in: linexcel is a lineage analyser first, and every figure in
the report is computed without a model. What AI adds is prose — a card per node
and an overview per workbook, written from the deterministic dossier the
analysis already produced.

Pick a provider first — see [Choosing an AI provider](providers.md). Every
example below uses a local Ollama runtime; substitute your own endpoint.

```python
from linexcel import analyze

result = analyze("workbook.xlsx")
docs = result.document(base_url="http://localhost:11434/v1", model="llama3.1")
result.save_html("out.html", docs=docs)
```

## Why the output is checkable

The model is never asked what a formula does. It is handed the node's dossier —
the exact formula, its step-by-step evaluation, its precedents and their values,
its dependents, the extent of a stretched group, any VBA link — and the system
prompt forbids asserting anything absent from it. Missing information must be
written as "not determined by lineage" rather than guessed.

So every claim in a card traces back to a formula or a value read from the
workbook, and the card sits next to that same evidence in the report. A reader
who doubts a sentence can check it without leaving the page.

## Node cards

```python
# Every calculation node (cells, groups, VBA procedures)
docs = result.document(base_url=..., model=...)

# Or a chosen few
docs = result.document(["c:Synthese!B3"], base_url=..., model=...)
```

`document()` issues `max_workers` requests in parallel (4 by default):

```python
docs = result.document(base_url=..., model=..., max_workers=8)
```

Documenting a large workbook is long and often billed, so a node that fails does
not discard the ones that succeeded. The successful cards are returned and a
`UserWarning` reports how many nodes were dropped; `AiDocError` is raised only
when every node failed.

## Workbook overview

```python
overview = result.document_workbook(base_url=..., model=..., language="en")
result.save_html("out.html", docs=docs, workbook_doc=overview, language="en")
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
overview = result.document_workbook(base_url=..., model=..., include_context=False)
```

The dossier is capped at `aidoc.MAX_WORKBOOK_DOSSIER_CHARS`. A workbook that
exceeds it sheds detail in order — long previews shrink, then the tail of the
pattern and VBA lists, and only as a last resort are previews and comments
dropped — so a small workbook loses nothing.

## Token usage

Every AI call is tallied on the result:

```python
docs = result.document(base_url=..., model="llama3.1")
overview = result.document_workbook(base_url=..., model="llama3.1")

print(result.token_usage)
# 48,210 tokens (44,900 in + 3,310 out) over 5 request(s) [openai-compatible/llama3.1]

usage = result.token_usage
usage.input_tokens, usage.output_tokens, usage.total, usage.requests
```

Counts come from the provider when it reports them — the OpenAI-compatible
`usage` block is read directly, so the figure is the one you are billed on. When
a provider reports nothing (a custom callable, or a local runtime that omits the
block), the tokens are approximated instead and `usage.estimated` is `True`;
`str(usage)` then prefixes the numbers with `~`.

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
    stale. Multiply by your own current rate — or run locally, where the rate
    is zero.

## Capping the bill

`token_usage` reports what a run cost *after* it ran. `token_budget=` decides
what it is allowed to cost before it starts:

```python
docs = result.document(base_url=..., model=..., token_budget=200_000)
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
result.document_workbook(base_url=..., model=..., token_budget=200_000)
result.document(base_url=..., model=..., token_budget=200_000)  # same ceiling
```

A budget already spent by earlier calls raises `AiDocError` rather than sending
a request that would breach it. `token_budget=0` — or any non-positive value —
raises `ValueError`; to send nothing at all, do not call `document()`.

!!! tip "Estimating before you spend"

    A dry run costs nothing: point `provider=` at a callable that records its
    prompt and returns `""`. `result.token_usage` then holds the estimated input
    cost of the whole workbook, which is the bulk of the bill for
    documentation-shaped work.

    ```python
    result.document(provider=lambda system, user, *, temperature=0.2: "")
    print(result.token_usage)  # ~ input cost of documenting this workbook
    ```

## Language

`language=` selects the system prompt sent to the model, and the same value
drives the viewer interface. Nine are available; see
[Languages](languages.md).

```python
docs = result.document(base_url=..., model=..., language="ja")
result.save_html("out.html", docs=docs, language="ja")
```

## What gets sent

Node dossiers and, for the overview, the workbook context. See
[Data handling](data-handling.md) for the payload of each call and where it
goes.
