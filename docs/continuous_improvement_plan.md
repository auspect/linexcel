# Continuous improvement plan

Planning audit: 2026-10-09; archive reconciliation: 2026-10-16.
This page records decisions and acceptance gates;
it does not certify the present working tree. Current delivery rules live in
[`AGENTS.md`](https://github.com/auspect/linexcel/blob/main/AGENTS.md) and
[local acceptance](manual_validation.md).
Historical test counts and merge verdicts apply only to their recorded snapshots.
Private workbooks, prompts, model responses and captures stay under the ignored
`validation_screenshots/` directory or in a private parent directory.

## Priorities and bounded work

| Order | Work | Decision and evidence needed |
| --- | --- | --- |
| P0 | `linexcel serve` navigation and responsive forms | Reproduce sheet selection, left-side overflow and graph readability in the live application; correct observed defects and add browser regressions for the failing interactions. |
| P0 | Current acceptance | Run both generated workbooks with node documentation, overviews, screenshots and image analysis in every language requested by each fixture (Sales: French and English; Stress: English); inspect `validation.json`, then complete independent factual and desktop/mobile review. |
| P1 | Specialized AI review | Version a source-grounded review rubric and compare the same sample before/after prompt changes. Keep lexical quotation checks as one check, with factual review recorded separately. |
| P1 | Exact source lineage | Retain separate original XML formula text, parser canonical formula and decomposition expressions. Attribute an “exact source” quotation only to the immutable original field and workbook hash; expose normalization and shared-formula expansion status. |
| P1 | Excelize formatting trial | Use an optional development oracle; compare raw, typed and formatted readings separately. Adopt a display adapter only after correctness, reproducibility, packaging and performance gates. See [package research](package_research.md#excelize-for-python). |
| P1 | Open issue #65 | Repair the scaling experiment, measure worker memory and actual graph size, and report usable browser interactions rather than completion alone. |
| P1 | Markdown consolidation | Eleven obsolete mandates and scoped reviews are preserved under [Historical records](history/index.md), with their original bodies and search exclusion. Current policy and evidence gates remain here and in local acceptance. |
| P2 | Deterministic contracts | Specify stable source/graph output separately from runtime telemetry and AI prose; publish a versioned additive graph schema and repeatability checks. |
| P2 | Package hardening and performance | Apply the measured adoption sequence in [package research](package_research.md#adoption-sequence-and-evidence-gates); keep Power Query scope and extraction uncertainty explicit. |

P0/P1 describes urgency and experiment order, not evidence of a verified defect
or a commitment to adopt a dependency. Each implementation change needs a
reproduction, a bounded change, meaningful regression checks, updated limitations
and the applicable local acceptance evidence.

## Review the live application as a user

Use `uv run linexcel serve` against the edited package. Walk through opening a
workbook, selecting a sheet, running analysis, inspecting a total, following an
input downstream, opening sheet context and returning to the graph. Test both
generated workbooks, a workbook with no formulas and a dense/disconnected graph.
Record the URL, workbook hash, source snapshot, active filters and viewport for
each finding. An exported standalone viewer and the application's controls need
separate checks where they have different behavior.

For sheet selection, distinguish the analysis request's sheet scope, Graph's
display filter and Sheets' presentation selector. Confirm selected values reach
the relevant request or render state, the visible sheet heading changes, directly
connected cross-sheet context remains truthful, and returning to all sheets
composes with other filters. A selector's visible text alone proves no behavior.

At 1440×900 and 390×844, inspect every dashboard tab, the graph and relevant
forms. Capture screenshots and read them alongside bounding boxes. Verify long
paths/formulas, notices, button groups, focus outlines, modal controls, selected
node panels and keyboard traversal. Check that controls remain inside their
container, that horizontal scrolling is confined to content which needs it, and
that graph-only controls and notices have the intended tab scope.

Assess graph overview by tasks: identify the main output, recognize sheet or
component structure, fit the visible graph, read a label on demand, and focus a
useful neighborhood. Check wide and deep chains, isolated components and high
degree nodes. A graph that fits inside the canvas can still be unreadable.
Reduced views must expose hidden connections and must not imply exhaustive
lineage. Review layout, zoom, pan, filter composition and return navigation with
real clicks, not only CSS assertions.

The 2026-10-09 live review reproduced and corrected sidebar overflow, stale
sheet replies at both neighborhood and hierarchy stages, and clipped graph
framing. Sheet navigation now prioritizes a formula and falls back to source
data; a reader toggle exposes the full canvas. Independent desktop/mobile
captures still show small labels in fitted dense mobile views and a tall
toolbar before the canvas. Treat shorter controls and label readability as a
follow-up, not as proven solved by the fit operation. Sales capture generation
failed with the requested 512 MiB budget and succeeded at 2,048 MiB; the precise
resource failure has not been diagnosed. Generation and factual review results
belong to the separate fingerprinted local acceptance record.

The 2026-10-16 implementation keeps primary controls compact and moves secondary
actions into graph options. Native browser tests reproduced a missed second
card click when a readability notice moved the canvas without invalidating
Cytoscape's hit-test rectangle. The correction resynchronizes geometry after
layout while preserving the camera and node positions. Independent reproduction
at both sizes verifies that case; complete current-corpus visual acceptance is
still a separate gate. Dense fitted graphs still require zoom, local exploration
or the source reader to inspect individual labels.

## Open issues: evidence before closure

The read-only GitHub audit on 2026-10-09 found two open issues. No issue is closed
by this planning document, and no external comment is authorized by this audit.

| Issue | Current decision | Next evidence and closure gate |
| --- | --- | --- |
| [#65: reliability and scaling](https://github.com/auspect/linexcel/issues/65) | Keep open; existing small stress runs do not establish large-workbook or viewer limits. | Parameterize 1M/5M/10M/19M dense-read cases and record the actual worker's peak memory, budgets, versions and per-phase time. Exercise broken references, dead names and unresolved external links across sizes. Build about 100k actual post-grouping graph nodes, record HTML bytes and test usable interaction times and captures. Publish anonymized conclusions only after reviewing provenance and omissions. |
| [#37: selective formula fallback](https://github.com/auspect/linexcel/issues/37) | Keep the current guarded engine path; no new fallback without a present failure. | Produce a minimal current-version failure, an independent verified result and supported semantics; test dependencies, errors, dates, arrays and provenance before adding a bounded optional fallback. If the premise no longer reproduces, document that evidence for a human issue decision. |

The 2026-10-09 audit initially found an unintended self-reference in the
dense-chain stress fixture; its measurements cannot establish recovery scaling.
Use a literal starting input, report formula count and post-grouping graph-node
count separately, and inject a demonstrated recovery-triggering failure. A
missing sheet can now produce a native `#REF!` rather than trigger recovery.
Successful completion with hundreds of unavailable readings is resilience
evidence, not recovery or parity evidence.
The XML formula iterator's shared-slave text is untranslated master text; do
not describe it as independently expanded source formulas.

## Preserve exact formula source separately

The 2026-10-09 generated Stress source check found forty pairs of parentheses
around `1+2` in `Lookups!B22` in the immutable workbook XML, while the graph
reports `=1 + 2`. The saved and engine readings are both `3`; this small
arithmetic check supports that case, but the normalized spelling is not a
verbatim source formula and does not establish general normalization parity.

Retain three distinct fields: original XML formula text and attributes, parser
canonical formula, and each decomposition expression. Record the workbook hash,
sheet/cell identity and normalization method. For shared formulas, preserve the
master text, `si`/`ref` relationship and whether a cell-specific expression was
actually expanded; do not label a reused untranslated master as the slave's
exact source. Keep these source records separate from calculated values and AI
prose.

Acceptance should replay the parenthesis fixture, function case/spelling and
shared-master/slave cases. Require “exact source” quotations to match the original
field; label canonical or expanded expressions explicitly and test their
attribution independently of numerical agreement. An omitted or unavailable
original field must remain a provenance limit rather than an exact-source claim.

The current implementation now carries additive `sourceFormula`,
`canonicalFormula`, `formulaProvenance` and `sourceValueProvenance` fields through
legacy analysis, lazy indexes and steps. Original text is decoded worksheet XML,
with a leading `=` added for presentation; it is not raw XML bytes. Shared-slave
translation stays separate when no text was stored. The same 4 Sales and 27
Stress targets retain their XML formula text, and their new original fields
match that source. This establishes the source-field improvement on that sample;
numerical semantics and new AI prose still require independent review.

## Specialized review for AI documentation

Version the rubric, prompt templates and payload schema. Supply the source
formula, expanded/representative status, typed cached and computed readings,
per-reading provenance, evaluated/unavailable steps, direct graph connections,
group sample bounds, unresolved references and coverage limits. Keep workbook
content as evidence rather than executable instructions.

The current implementation supplies the versioned localized
`linexcel-factual-review-v5` rubric, a typed claim ledger, source/canonical/cache
payloads, literal lookup offsets and observed directed topology. A second
specialized review request receives the same bounded dossier, the draft and
deterministic findings before returning a revised document. Both stages and
their aggregate usage are retained; this same-provider review is not independent
factual certification. See the [method and limits](ai_factual_review.md).
The next evidence gate is the complete before/after replay, including both
request receipts, ledger compliance and independent source/pixel review.
Early MiniMax output often paraphrases ledger statements rather than quoting
exact Markdown, leaving them indeterminate; absence of a repeated old error is
not proof that the underlying defect is corrected.

Use this repeatable review procedure for a reproducible sample spread across
sheets and formula families, including guarded failures, volatile formulas,
dates, groups, dynamic references, external links, Power Query and empty states:

1. Match quoted formulas and cell addresses to the supplied source evidence.
2. Check every value claim against the reading, type and provenance; distinguish
   cached, engine, fallback, volatile and unavailable states.
3. Check dependencies and conditional logic against formulas and graph coverage;
   a static edge or an evaluated representative is not proof of every runtime
   dependency or every grouped cell.
4. Check decomposition claims against completed steps, including local bindings
   and unavailable evaluations. Reject invented causes of disagreement.
5. Check descriptions of screenshot rows, columns, labels, clipping and layout
   against the actual image, not the model's own prose.
6. Record each supported, contradicted or unverifiable claim with a node/sheet,
   exact supporting evidence and severity. An independent reviewer should not
   see the generator's self-verdict as an answer key.
7. Replay the same evidence after a prompt change; compare factual failures,
   abstentions, missing/truncated responses, latency and token usage separately.

Package each reviewed claim as a record with `response_span`, `claim_kind`,
`target_node` or `sheet`, `evidence_pointer`, `verdict`, `severity` and
`correction`. Point to source evidence independently of the generated response.
The dossier should expose binding attribution, `value_source` and origin kind,
`evaluated` status, semantic risks, verification, branch/merge topology, group
geometry and sampling, external disk reads versus embedded caches, and omitted
metadata. Vision claims need an image and bounding box with visible, ambiguous
or outside-image status; pixels cannot establish a formula's numerical cause.

Challenge the reviewer with altered prose containing wrong values, wrong node
attribution despite a valid quotation, cache-to-engine attribution, representative
samples generalized to all cells, branches described as linear chains, embedded
external caches described as disk reads, and invented image columns or clipping.
Report reviewer false positives and missed defects by category separately from
generation coverage. Reuse seeded node selection across both workbooks; the
complete delivery run must still cover every language requested by the fixtures.
A provider seed
or low temperature is replay metadata, not a determinism guarantee.

Deterministic assertions can verify source quotes, supported identifiers and
explicit numeric/provenance fields. They cannot certify arbitrary prose,
business purpose or visual understanding. AI review can propose changes; a
grounded counterexample decides a factual defect. Preserve failed responses and
request metadata locally so retries do not erase evidence.

For this request, the user selected `MiniMax-M3` for documentation and the same
gateway for vision when supported, using `OPENAI_API_KEY` and `OPENAI_BASE_URL`
from an ignored `.env`. The gateway's simple red-square/blue-circle image probe
returned a correct description; that establishes basic image acceptance, not
the factual correctness of workbook image descriptions. Load the file explicitly
with `--env-file .env` and record the resolved non-secret model/provider
configuration in the run. See [local acceptance](manual_validation.md) for the
command. A failed vision stage must remain incomplete; do not replace it with a
text-only assertion. This run-specific choice does not change the permanent local
Ollama default.

## Determinism and lineage contract

Define determinism for immutable input bytes, tool versions and analysis options.
Test stable ordering, IDs, formulas, edges, value types, provenance, coverage and
failure status with repeated isolated runs. Separate canonical analysis data
from durations, run paths and provider telemetry; byte-identical telemetry or
unseeded model prose is not a useful promise. Clock/seed-controlled calculations
need recorded context shared by the result and its decomposition.

Keep raw source tokens, typed saved caches, formatted display text and computed
readings distinct. Document any numerical comparison tolerance and date
normalization. Group-level findings must state sampling scope. Schema changes
should be additive with explicit enum/version handling; neither a formatter nor
a viewer should reinterpret Python's comparison verdicts.

## Markdown inventory and cleanup decisions

The inventory includes tracked Markdown and a search for repository agent/skill
instructions. Only root `AGENTS.md` was found as a repository agent policy.
The 16 October cleanup moved eleven named historical files into
`docs/history/pre-1.15/`. Their original bodies remain byte-identical below a
dated historical banner; current links and the documentation navigation now
point to the archive. Runtime prompts and the active root `AGENTS.md` are kept.

| Files | Decision | Reason and next action |
| --- | --- | --- |
| `AGENTS.md` | Keep active | Current acceptance contract; user-selected model exceptions apply per run. Do not import old mandates or past validation waivers into it. |
| `README.md`, `docs/index.md` | Keep and align | Public entry points; make the distinction between `serve`, standalone reports and development validation easy to find. Recheck duplicated provider/model examples when their behavior changes. |
| `docs/guide/*.md`, `docs/api/*.md`, `tests/fixtures/README.md` | Keep reference documentation | Check commands and links against current APIs. Historical sections within guides need dated scope, not removal of current contracts. |
| `docs/manual_validation.md` | Keep procedure | Centralize how to run acceptance; allow explicit provider/model selection without weakening complete generation and review requirements. |
| `docs/package_research.md`, `docs/powerquery_audit.md` | Keep current research/audit | Dated proposals with limits and source links. Move implemented findings to guide/changelog; carry unresolved actions into this plan. |
| `docs/formualizer_upgrade_plan.md` | Split active plan and historical evidence later | Recent 0.11.0 addendum and integration gates remain useful; old version-specific counts and a previous `--no-ai` waiver apply only to that intervention. Keep those as history after resolving their durable actions. |
| `docs/guide/acceptance.md` | Keep dated baseline | Already states the PR #98 snapshot and known factual limits. Add later evidence as separate fingerprinted records, never silently update old test counts. |
| `docs/guide/calculation-investigation.md`, historical portions of `docs/guide/engine-safety.md` | Merge current conclusions, retain dated experiment | September prototypes and version claims must be reconciled with installed adaptations and October engine work before being treated as present limitations. |
| `docs/reliability_study.md`, `docs/integrated_quality_review.md` | Archive after merging remaining actions | Valuable dated findings and measurements; retain provenance, sampling limits and regression rationale. Earlier test totals are not current acceptance. |
| `docs/viewer_builder_brief.md`, `docs/viewer_review_brief.md`, `docs/perf_brief.md` | Archive obsolete mandates | Branch/path/owner instructions belong to past tasks. The viewer briefs' `uvx` command can validate the published package instead of local changes; current acceptance uses `uv run`. |
| `docs/viewer_polish_notes.md`, `docs/viewer_review_notes.md`, `docs/viewer_final_review.md` | Archive scoped reviews | Commit-specific merge verdicts and old test counts; some visual judgments used DOM/CSS rather than reviewed screenshots. Merge still-useful accessible-graph and dense-layout tasks before archiving. |
| `docs/swarm_v1.8_SPEC.md` | Archive obsolete coordination | Owners, line numbers, test counts and English-only UI instruction are task-specific; the product now documents nine languages. Durable provenance rules belong in current contracts. |
| `docs/swarm_v1.8_research.md`, `docs/swarm_v1.8_axisA_audit.md` | Merge verified conclusions, archive originals | Research raises hypothetical misattribution risks while the subsequent small audit rejects current defects. Preserve that distinction; retain proposals about schema/tolerance only with new counterexamples or a measured need. |
| `src/linexcel/assets/prompts/*/*.md` | Keep runtime assets | These Markdown files are executable product inputs, not obsolete reports. Maintain all supported languages; review changed templates and complete the fixture language scope in the full AI acceptance run. |
| `CHANGELOG.md`, `SECURITY.md`, `THIRD_PARTY_NOTICES.md` | Keep durable records | Release history, reporting policy and license notices have separate purposes; stale review cleanup must not erase them. |

The named cleanup batch is implemented; its delivery still requires a strict
documentation build and incoming-link check. The inventory's old paths above
identify original locations; the [archive index](history/index.md) supplies their
current links. Keep the engine upgrade audit available with a historical-scope
banner because its candidate-version contracts remain relevant.

Surviving actions from those records are bounded as follows:

- Preserve typed input/cache/engine distinctions, error and unknown states,
  literal text comparisons, original formula text and unevaluated local bindings.
- Verify accessible alternatives, keyboard focus, canvas hover cleanup and empty
  states in the actual viewers; old screenshot or DOM verdicts do not certify
  the current interface.
- Keep cyclic models and volatile snapshots qualified; neither stored caches
  nor one converged engine observation establishes a unique Excel result.
- Treat dense graph fitting as structural orientation. Use local exploration,
  search and readable selected-node sources for detailed navigation; measure
  interaction and memory before claiming arbitrary scale.
- Continue the existing engine compatibility matrix before removing adaptations
  or adopting a newer engine. Package research remains a measured trial queue.
- Keep issue #65's large-scale evidence gates and issue #37's counterexample
  requirement open until their specific evidence exists. The current release
  must not claim those issues closed through small-workbook acceptance.

Before any delivery, attach the exact run and working-tree fingerprints, factual
review findings, captures at both required viewports and independent review.
Unit tests, documentation builds and package generation are complementary
evidence. Missing dependencies, empty/truncated AI output or an incomplete
language/vision run must be reported as failed stages.
