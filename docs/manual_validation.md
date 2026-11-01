# Local acceptance

Run the real pipeline on the current working tree before delivering changes to
analysis, evaluation, documentation, screenshots or the viewer. Automated tests
complement this run; they do not replace it.

```powershell
uv run --extra screenshots validate_manual.py --workbook both --no-ai
```

AI documentation and image analysis are optional. This command validates both
workbooks, all report languages and sheet captures without contacting a model.
The manifest marks AI stages as `skipped`; it does not claim AI coverage.
Use `--no-vision` to retain text documentation but skip image analysis.

To include both AI stages:

```powershell
uv run --extra ai --extra screenshots validate_manual.py --workbook both --vision --model qwen3.8 --vision-model qwen3.8 --token-budget 1000000
```

For AI validation, Ollama must serve the selected model locally.
LibreOffice and Poppler must be available for workbook rendering. The default
endpoint is `http://localhost:11434/v1`; there is no cloud fallback. Both
generated workbooks and image analysis are enabled by default. Calls use one
worker by default to avoid competing for a local model. A token budget is checked
between requests; it includes documentation and image descriptions across
languages for each workbook. Increase it explicitly when the manifest reports
missing responses.

For a local runtime supporting the OpenAI-compatible `reasoning_effort` field,
`--reasoning-effort none` explicitly disables separate reasoning. The option is
omitted by default and is recorded in the run configuration. It applies to both
documentation and vision; it does not change the selected model, response
completeness requirements or factual review requirements.

The manifest separates pipeline generation from factual and visual review in
its `acceptance` fields. A generation success alone leaves delivery not ready.
Review evidence must identify the inspected sources, reports and screenshots.

## Explicitly selected gateway

The user can choose a different provider and model. `--env-file` explicitly
loads an ignored local file containing `OPENAI_API_KEY` and `OPENAI_BASE_URL`;
keep the file and its secret values out of source control and diagnostic output.
Existing environment variables take precedence over values loaded from the file.
For the current review, the selected gateway accepted a simple image probe for
`MiniMax-M3`. Run documentation and vision through that gateway with:

```powershell
uv run --extra ai --extra screenshots validate_manual.py --env-file .env --workbook both --vision --model MiniMax-M3 --vision-model MiniMax-M3 --api-thinking disabled --reasoning-split --api-stream --max-tokens 32768 --api-retries 2 --token-budget 2000000
```

`--reasoning-split` requests a separate reasoning field from providers which
support it, so reasoning blocks do not become displayed documentation. This
option is explicit and is not sent by the default local run.
[MiniMax's API option](https://platform.minimax.io/docs/api-reference/text-openai-api#minimax-m31-flash-preview--minimax-m3-request-parameters).

`--api-thinking disabled` explicitly asks the selected MiniMax-M3 to answer
directly. Its default thinking mode can consume the response limit before final
content appears. `adaptive` enables provider thinking; omitting the option sends
neither setting. Both are recorded in the manifest and apply to documentation
and vision. MiniMax documents `disabled` for M3; Flash Preview rejects it, so
do not copy the option to a different model without checking its support.
[MiniMax's thinking configuration](https://platform.minimax.io/docs/api-reference/text-openai-api).

`--max-tokens` can bound each response, and `--api-retries N` explicitly selects
the SDK's transient-failure retries (default zero). Record both in the manifest.
A response ending at its length limit still fails acceptance. Retries use the
same provider and model; the SDK does not report usage of failed attempts.
Failure traces retain exception types and numeric status/OS codes, never gateway
bodies, headers or credentials.

For the selected MiniMax model, this output ceiling includes reasoning and
final content, even when reasoning is separated from displayed documentation.
The historical `20261010-minimax-m3-lineage-corpus-final` run reached its
16,384-token ceiling on the French `Summary!C5` card and ended with `length`;
that run remains incomplete. The subsequent candidate run uses 32,768 tokens
per response. A higher ceiling does not guarantee completion or factual
accuracy: every response must still contain nonempty final text and end with
`stop`, and source and image review remains required.

`--api-stream` explicitly receives text and image descriptions through the same
provider as streamed completions. Separate reasoning deltas are excluded from
displayed documentation. A stream must contain nonempty text and finish with
`stop`; an interrupted stream or `length` finish fails. The stream is closed
on every exit, and terminal usage is retained when supplied. SDK retries do not
restart a response that fails after streaming has begun. Streaming and final
usage chunks are described by the
[OpenAI protocol](https://developers.openai.com/api/reference/resources/chat/subresources/completions/streaming-events)
and supported by the selected
[MiniMax gateway](https://platform.minimax.io/docs/api-reference/text-openai-api).

An image probe checks request compatibility, not factual accuracy on workbook
screenshots. Inspect the generated descriptions against their original images.
If vision needs a separate explicitly chosen endpoint, use `--vision-base-url`
and its matching `--vision-model`; record both configurations in the run. Do not
silently retry a failed gateway against an unrelated provider. The documentation
and vision endpoints can differ; an explicitly requested stage must complete.

The generated Sales fixture requests French and English; the Stress fixture
requests English. Complete every requested report language and screenshot.
When AI is requested, complete each node card, workbook overview and image
description for the selected stages. A multilingual product does not
make every fixture request all nine supported languages.

## Native visual evidence

Image analysis covers every page of each sheet with regions of at most
1024×1024 pixels, overlapping by 128 pixels. Pixels retain their native scale;
large tables are not reduced to a fitted overview for the model. The built-in
provider sends exactly one region per request, physically separating each
crop from its neighbors. Earlier four-region requests allowed the model to
attribute text from one crop to a blank neighboring crop. Pillow, included in the screenshots extra,
decodes images and creates the regions.

Each response must identify all requested regions in a JSON record with
readability and short contiguous quotations. The program assigns each quotation
the exact native crop as its scope, with coordinates computed from the source;
the model supplies no glyph geometry. Missing regions, invented geometry,
free prose, empty responses and truncated responses fail
the image stage. An unreadable region can explicitly abstain; this records
coverage and a readability limit, and does not establish factual correctness.
Numbers and punctuation must be transcribed as visible, without combining
separate cells or attributing numbers to unseen headers. Clipping, row counts
and interpretations of the workbook are excluded from model output. The crop
coordinates identify physical context limits; a crop boundary is not evidence
of a rendering defect.

Invalid regional JSON can trigger one complete new generation for the same
crop. Both attempts retain their exact prompts, raw responses, diagnostics,
completion receipts and usage. Only a complete valid response is selected;
no observation is silently trimmed or repaired. Transport, empty or truncated
responses are not retried this way, and the budget can prevent a retry.

The rendered description presents six observations spread across the image;
the remaining observations are available in an expandable list. It is a
deterministic presentation of the bounded records, with no additional model
summary. Check the records against the exact saved native regions, including
later columns and rows.

The language evidence JSON includes `vision_checks`: original image SHA-256,
native region coordinates and SHA-256, exact crop files, batch region IDs,
JSON-prompt hashes, raw model JSON and completion receipts. The `vision_regions`
gate verifies complete region processing, matching receipts and stored crop
bytes. Request counts depend on image dimensions and provider capabilities;
one sheet description can require several requests. Budgets accumulate across
all batches and languages. The historical two fixtures and public workbook
images yielded 167 native regions and 56 vision requests with the previous
four-region batches. Those same images require 167 isolated regions to be
resolved: only nonuniform regions need model requests, while exact uniform-pixel
proofs stay local. Invalid-response retries can add requests; rerendered images
can change these counts. Coordinate validity and
generation coverage never certify the truth of an AI observation.

## Replay the public corpus

Acquire the pinned public sources and licenses once into the ignored local
directory. The downloader verifies HTTPS source/redirect hosts, size, file
signature and SHA-256. It refuses changed upstream payloads rather than silently
accepting a different workbook. Its source inventory is
`scripts/public_corpus_sources.json`; downloaded sources and reports stay local.

```powershell
uv run python scripts/fetch_public_corpus.py --output-dir validation_screenshots/public_corpus_20261009
uv run --extra ai --extra screenshots validate_manual.py --env-file .env --workbook both --corpus-manifest validation_screenshots/public_corpus_20261009/manifest.json --vision --model MiniMax-M3 --vision-model MiniMax-M3 --api-thinking disabled --reasoning-split --api-stream --max-tokens 32768 --api-retries 2 --token-budget 2000000
```

The token budget applies to each workbook, across its languages, node cards,
overview and image batches. Size it for the selected model: MiniMax-M3 consumed
1,013,330 reported tokens for the Stress overview and 67 cards in the local
10 October run, leaving no budget for vision under a one-million limit. That
run remains incomplete; increasing the budget requires a new full run and does
not certify its factual accuracy. A completed request can cross the limit,
which then prevents further requests.

Use this manifest on each subsequent full acceptance run. Both generated
fixtures remain included. Every public **workbook** receives complete node
documentation, overview, sheet captures and vision in English, including a
data-only workbook's overview and all sheet images. The input fingerprint is
checked before model requests and again after the case; originals are never
rewritten. A manually selected `--file` does not replace corpus entries.

**Add-in** entries have a separate static-source stage: all VBA modules and
procedures, provenance and a reviewable HTML graph, with no execution or AI/
worksheet coverage claimed. Inspect their tabs and a reproducible procedure
sample independently. Legacy XLA worksheet captures are unavailable by design;
do not substitute an invented screenshot or mark this as a worksheet pass.

For each workbook inspect every sheet and exported/live dashboard tab at
1440×900 and 390×844. Record image dimensions and visible ranges: a fitted
95-column table is not readable at ordinary zoom, and a top-left context preview
does not cover the other columns. Read later columns/rows in the original source
and enlarge the actual capture to verify the relevant values. Check image claims
against pixels, keeping rendered LibreOffice values separate from original
saved caches. Follow the [specialized factual-review protocol](ai_factual_review.md)
to replay previous findings on the same node sample; report unreviewed cards and
remaining discrepancies even when generation coverage succeeds.

## Evidence and review

Each invocation creates a new directory under the ignored
`validation_screenshots/manual/`. `--output-dir PATH` chooses another new
directory, such as a private sibling workspace. Existing directories are
refused. `--file PATH` reads a real workbook without rewriting it. Keep real
workbooks and all reports embedding their contents outside GitHub.

The run saves graphs, PNG files, AI responses, HTML reports and `validation.json`.
The manifest records working-tree source hashes, configuration, input and image
hashes, durations, token usage and coverage by language. A successful process
requires every requested node card, overview, image and image description.
The language evidence files also retain completion receipts: requested and
reported model, response ID when supplied, finish reason, content/prompt-text hashes,
reported or explicitly estimated usage, and complete/failed status. The content
hash covers the stripped final text before documentation qualification; it
excludes separate reasoning. `prompt_text_sha256` covers the system and user
text, excluding request JSON and image bytes; capture hashes record images
separately. Transport failures retain a separate sanitized
trace and unknown usage when the SDK supplies no final counts.
`provider_receipts.jsonl` also appends each received completion and its usage
immediately, including rejected responses, before a language finishes. This
private recovery journal preserves completed calls if a long run is interrupted;
it does not prove coverage or recover usage for requests still in flight.
Missing providers and empty or truncated responses fail when the AI stage is
requested. `--no-ai` and `--no-vision` explicitly skip optional stages and can
succeed without them. The manifest records the scope and per-language stage
status. Incomplete languages, missing captures, `--no-recalc`, `--no-screenshots`
and `--max-nodes` still return a nonzero exit status. A node limit uses
reproducible sampling across sheets; `--seed` controls that sample.

When AI documentation is requested, every expected text card and overview must
complete both drafting and specialized review. The `documentation_phases` gate
binds their exact system/user prompts,
dossier hash, raw responses and usage to two distinct successful provider
receipts. A missing, disabled, truncated, duplicated or unlinked phase fails
acceptance; the final report must match the reviewed response. Both calls count
toward the run's budget. Their successful linkage establishes traceability,
not the factual truth of the document.

Generation coverage does not establish factual correctness. The manifest leaves
manual review pending. Open every dashboard tab, inspect graph connections and
sample nodes across sheets and difficult formulas. Check the original formula,
each evaluated step, cached/recalculated value provenance and the AI's claims.
Check each image description against its source image. Capture the dashboard
at desktop and mobile widths and inspect text, overflow, selection, navigation
and image readability. Record discrepancies with node or sheet identifiers,
the exact run directory and the evidence that supports the finding.

The stress fixture includes intentionally invalid references, circular formulas
and a formula longer than Excel's 8,192-character limit. It tests resilience;
it is not a fully Excel-compatible correctness oracle. LibreOffice is a second
implementation, so disagreement with its saved values requires investigation.
An AI description is another observation to review, not an independent numeric
oracle.
