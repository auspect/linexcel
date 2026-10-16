# AI documentation

Optional and opt-in: linexcel extracts or computes deterministic evidence
without a model, including saved workbook values and calculation qualifications.
What AI adds is prose — a card per node
and an overview per workbook, written from the deterministic dossier the
analysis already produced.

Each requested text document uses two model calls by default: draft, then a
specialized source-led review. Both calls count toward the token budget. The
same provider performs both phases; this is an editorial check, not independent
verification. Empty or failed review responses are not replaced with an
apparently reviewed draft. No model is called until documentation is requested.

Use `specialized_review=False` explicitly on `result.document(...)` or
`result.document_workbook(...)` for a single draft request. The evidence report
then records that review was disabled. The default preserves both prompts,
RAW responses, hashes and usage under `specialized_review.stages`; those records
may contain source data and need the same privacy as the workbook.

Pick a provider first — see [Choosing an AI provider](providers.md). Every
example below uses a local Ollama runtime; substitute your own endpoint.

```python
from linexcel import analyze

result = analyze("workbook.xlsx")
docs = result.document(base_url="http://localhost:11434/v1", model="qwen3.8")
result.save_html("out.html", docs=docs)
```

## Why the output is checkable

The model receives the node's dossier —
stored original XML formula text when available, the separate canonical
expression, its decomposition and recorded evaluation, its precedents and their values,
its dependents, the extent of a stretched group, any VBA link — and the system
prompt forbids asserting anything absent from it. Missing information must be
written as "not determined by lineage" rather than guessed.

The card sits next to that evidence in the report, so a reader can check a
sentence against its formula and values. A prompt is not an accuracy guarantee:
local reviews have found incorrect explanations despite complete responses.

Source name definitions include constants and worksheet scope, independently of
which names become graph nodes. A missing graph edge or a `#NAME?` result does
not establish that the source definition is absent. Metadata too large to
inspect or include is marked incomplete. Group values explicitly identify their
representative cell; they are not group totals or values for unsampled members.

The dossier distinguishes group membership from value sampling. A represented
group can fill its entire bounding rectangle while only a few member values are
included. It also separates an input value read by the engine from a formula
result, and gives range dimensions as rows, columns and cells. Missing evaluated
step values do not by themselves establish whether evaluation was attempted.

Omissions describe the bounded documentation payload. A neighbor removed to
fit that payload can still exist in the graph. Workbook overviews receive
directed connections with their own coverage counts; ordering formula examples
does not establish a calculation chain. These facts help review the explanation,
but do not prevent a model from misinterpreting them.

For literal `AGGREGATE` selectors, the dossier records their
[Excel meanings](https://support.microsoft.com/en-us/excel/functions/aggregate-function).
For example, function 9 selects SUM and option 6 excludes errors, not hidden
rows. Dynamic selectors remain undetermined. These facts describe the formula;
they do not independently verify the engine's implementation or result.

Literal lookup indexes also have explicit facts:
[VLOOKUP's index starts at one](https://support.microsoft.com/en-us/excel/functions/vlookup-function),
so index 5 selects the fifth table column, four columns from the first.
The function inventory describes
syntax rather than engine support. Saved caches remain separate from engine
results; deliberately preserved volatile readings do not prove inability to
calculate. Neither a mismatch nor an error establishes its cause or age.

## Formula quotation checks

Documentation checks complete formula quotations in Markdown against source
formulas in the exact dossier sent to the model. An unsupported quotation adds
a visible notice identifying the quotation; the generated explanation is not
silently corrected. Literal illustrations and incomplete snippets can remain
uncheckable. Processing limits are reported explicitly.

This is a bounded lexical check, not a factual verdict. A supported quotation
does not establish correct attribution, arithmetic, causality, value provenance
or the truth of surrounding prose. Screenshot descriptions require separate
visual review.

Both documentation APIs accept an optional evidence collector:

```python
checks = {}
docs = result.document(base_url=..., model=..., validation_results=checks)
overview = result.document_workbook(
    base_url=..., model=..., validation_results=checks,
)
```

Entries are keyed by node ID, with `workbook` reserved for the overview. Each
entry retains the raw model response, quotation checks, limitations and a
`qualified` or `unverified` status; neither means factually approved. Keep this
evidence private when its source workbook is private. The manual validation
runner saves it alongside the generated documentation for independent review.

Node and overview prompts also request a short `json_claims` ledger of material
statements, each bound to an exact JSON pointer and typed value in the sent
dossier. Its `factual_review` report records supported, contradicted or
indeterminate bindings. The ledger is removed from visible Markdown; raw
responses retain it. Unsupported factual bindings add a visible qualification.
A supported binding proves a value is present, without certifying the meaning
or attribution of the sentence. Missing ledgers remain explicitly unverified.
See the [specialized factual review method](../ai_factual_review.md) for the
independent source/image review and reproducible before/after comparison.

In local review, this check exposed a quoted `/100` where the source used
`/1000`. It also qualified conceptual shorthand such as `=C{row} * D{row}`:
that is not a literal source formula, but is not necessarily a false
explanation. Separately, generated prose miscounted sampled cells and described
a complete group as potentially containing gaps. Those prose errors are
outside this check's scope and still require comparison with source evidence.

A subsequent review using the explicit group, dimension and connection facts
observed correct complete-group descriptions, branching topology, range row
counts and input provenance in the selected cases. It still found false claims
about external caches, ambiguous localized illustrative formulas and invented
visual discrepancies. The six unsupported quotations in that run are lexical
observations, not a count of proven factual errors. Neither the additional
source facts nor a complete generation run certifies the explanation.

An OpenRouter review with `qwen/qwen3.8-flash` on 2026-10-07 also found a
space-separated range intersection described as a union with double counting,
and a cached volatile `TODAY()` value described as a fresh engine calculation.
The deterministic dossier preserved the intersection and marked the volatile
value as an unrecalculated snapshot. These were errors in generated prose;
formula quotation checks did not establish its correctness. Excel's
[reference operators](https://support.microsoft.com/en-us/excel/calculation-operators-and-precedence-in-excel)
distinguish a space (intersection) from a comma (union).

## Node cards

```python
# Every calculation node (cells, groups, VBA procedures)
docs = result.document(base_url=..., model=...)

# Or a chosen few
docs = result.document(["c:Summary!B3"], base_url=..., model=...)
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
a graph. A title sitting two rows above a table, a column called `Price` that
somebody hid, a comment reading *"exported product category"* — no formula
records any of it, and an overview written without them describes a structure
nobody recognises.

The images themselves stay put for this call: `openpyxl` reads the same facts
deterministically, so the "cite only the dossier" rule still holds and no vision
model is required. Sending them is a separate, explicit step — see below.

```python
# Lineage only — cell contents stay on your machine
overview = result.document_workbook(base_url=..., model=..., include_context=False)
```

The dossier is capped at `aidoc.MAX_WORKBOOK_DOSSIER_CHARS`. A workbook that
exceeds it sheds detail in order — long previews shrink, then the tail of the
pattern and VBA lists, and only as a last resort are previews and comments
dropped — so a small workbook loses nothing.

## Describing the screenshots

Everything above is grounded in the graph. This one is not: a screenshot shows
the printed values and labels produced by the renderer. Image analysis returns
short literal quotations from native regions. It does not infer calculations,
column meanings, hidden state or rendering defects from those pixels.

```python
shots = result.save_screenshots("shots/")            # LibreOffice renders them
checks = {}
seen = result.describe_screenshots(
    shots, base_url=..., model="<a vision model>",
    validation_results=checks, evidence_dir="vision-evidence/",
)
result.save_html("out.html", screenshots=shots, screenshot_docs=seen)
```

```bash
linexcel analyze book.xlsx --screenshots shots/ --vision-docs \
    --base-url http://localhost:11434/v1 --vision-model "<a vision model>"
```

Each description appears under the image it describes, in the Sheets tab,
badged *read from the screenshot* rather than as ordinary AI documentation —
the reader can tell a claim about pixels from a claim about the lineage.

Three things worth knowing:

- **The model must accept images.** `model=` here is where a vision model is
  named when it differs from the writing one (`--vision-model` on the command
  line). A text-only endpoint raises `AiDocError` rather than having the
  picture silently dropped from the request.
- **Valid scope does not prove a correct quotation.** Each response must name
  every requested region and return short contiguous quotations in a strict
  JSON contract. Coordinates are computed by the program from each entire
  native crop; the model supplies no geometry. A model can still misread
  characters. Compare its quotations with the saved crop bytes, preserving
  visible punctuation and distinguishing rendered values from saved caches.
- **This is the widest thing linexcel sends.** A picture of a sheet shows every
  row on it, including those no dossier would have quoted — see
  [Data handling](data-handling.md).

Every page is covered by crops of at most 1024×1024 pixels, with 128 pixels of
overlap and no rescaling. Every provider receives exactly one crop per request,
so a quotation cannot borrow pixels from a neighboring crop in the same call.
Custom providers must return the same regional JSON contract. An unreadable
region can abstain explicitly. A sheet is returned only when all its pages and
regions complete. Partial evidence remains in `checks`, and a failed sheet
raises `AiDocError` if none completes, or emits a `UserWarning` if others do.

An invalid JSON contract permits one complete regeneration for the same crop.
Both attempts preserve their full raw response, prompt identity, receipt and
usage; rejected replies are never shortened or repaired. Only the final valid
attempt contributes observations and completed coverage. Transport errors,
empty replies and exhausted budgets fail without this retry.

The deterministic description shows six quotations spread across the image,
with the others in an expandable list. `checks` retains original/crop hashes,
native coordinates, exact JSON prompts, raw responses and provider receipts
when available; `evidence_dir` saves the exact bytes sent. These are review
records, not factual certification. Pillow from `linexcel[screenshots]` is
required for decoding and cropping.

`token_budget` is checked before each call and after reported consumption.
Exceeding it on the last batch also prevents complete coverage; equality is
allowed. The tally remains explicitly estimated if an endpoint supplies no
usage counts. See [local acceptance](../manual_validation.md#native-visual-evidence)
for the complete coverage and independent pixel-review requirements.

## Token usage

Every AI call is tallied on the result:

```python
docs = result.document(base_url=..., model="qwen3.8")
overview = result.document_workbook(base_url=..., model="qwen3.8")

print(result.token_usage)
# 48,210 tokens (44,900 in + 3,310 out) over 5 request(s) [openai-compatible/qwen3.8]

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
one node's card is never the question. With the default specialized review,
900 documented formula patterns require 1,800 requests: draft and review each
consume tokens and are included in the tally. Use
`max_tokens=` when you want to bound an individual response.

Because a request's cost is only known once it has answered, the budget is
enforced *between* requests. Nodes still queued when the tally reaches the
ceiling are never sent, the cards already written are returned, and a
`UserWarning` names how many nodes were left undocumented:

```text
UserWarning: Token budget of 200,000 tokens reached after 201,447:
612 of 900 nodes documented, 288 never sent. Raise token_budget to document the rest.
```

Requests already in flight are allowed to finish. A completed final response is
retained with its exact cost even if that request exceeds the ceiling. A draft
that exhausts its node's remaining budget prevents the review and is never
published as reviewed. Each active node starts with a remaining-budget snapshot;
the tally may therefore include both phases of up to `max_workers` active nodes
before the queue stops. Set the API budget as an order of magnitude, not to the
token. The interactive server separately rejects jobs whose reported final
consumption exceeds their job budget while retaining private failure evidence.

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
    prompt and returns `""`. Empty documents are rejected, so this deliberately
    raises `AiDocError` after processing the nodes. `result.token_usage` still
    holds the estimated input cost of those calls.

    ```python
    from linexcel.aidoc import AiDocError

    try:
        result.document(provider=lambda system, user, *, temperature=0.2: "")
    except AiDocError:
        pass  # Expected: an empty response is not valid documentation.
    print(result.token_usage)  # ~ input cost of documenting this workbook
    ```

Empty replies and responses that the endpoint marks as truncated or filtered
are failures. Reported token usage from rejected replies remains counted toward
the budget. Partial node or image batches still return completed documents with
a warning, so acceptance checks must compare their keys with the requested set.

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

Exact original, stored XML and canonical formula proofs use deterministic
placeholders filled from the actual sent node dossier. Their insertion ledger
records the field, text hash and lengths separately from the raw AI response.
An omitted source stays omitted. The model explains the supplied logic and
records its own claims; it does not need to copy a long original expression.
See [specialized factual review](../ai_factual_review.md) for these boundaries.
