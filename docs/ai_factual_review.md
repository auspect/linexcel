# Specialized factual review

## Draft and actual specialized review

Text documentation now makes two model requests by default: a draft, then one
bounded editorial review against the same exact source dossier. The review
prompt requires checking every section, including the Role sentence, as well
as the claim ledger. It must distinguish unsupported claims from missing evidence,
remove invented corrections, reconcile digits and predicates, and preserve
unknown causes. The candidate and its own ledger are never factual evidence.

This second request is made to the same configured provider. It is **not an
independent oracle** and does not certify the final prose. Both requests consume
tokens; the final visible notice states the two-request method and its limits.
`specialized_review=False` explicitly opts out in `document_nodes`,
`document_workbook` and the corresponding `LineageResult` methods. Its report
records `disabled`, never a completed review.

`specialized_review.stages` preserves each phase's exact system/user prompts,
their SHA-256 hashes, RAW response/content hash, usage and status. The top-level
`raw_response` remains the final response. A failed, empty or oversized review
never silently publishes its draft as reviewed. Already consumed usage is still
counted; a budget exhausted after the draft stops before the review. Provider
receipts link both phases independently, including failures. This audit trail
contains workbook data and must remain private when the workbook is private.

The documentation API checks its approximate token ceiling between requests.
A complete final response already started below the ceiling may finish above
it; its text and exact usage remain available, and subsequent queued nodes stop.
Concurrent nodes each start with a remaining-budget snapshot, so both phases
of an already active node may still finish. This is two billed requests per
reviewed document, not one request with an unreported multiplier. The server's
job wrapper additionally enforces its existing strict check after consumption
and retains failure evidence if a job overshoots its own budget.

The workflow requires source identity first, then formula representation,
observed result provenance, arithmetic and branch bounds, reference operators,
group/member scope, directed topology, and finally literal ledger bindings.
An improvement in one paragraph cannot excuse a contradiction in another.

The `linexcel-factual-review-v5` contract makes claims traceable to the exact
dossier sent to the model. It improves the evidence available for generation and
adds deterministic checks of explicit typed bindings. It does **not** certify
the meaning, attribution or completeness of prose, or the accuracy of vision.

Version 3 adds measured formula spelling, typed literal-string comparisons and
separate full-source/graph/excerpt bounds for VBA. The dossiers provide these
facts directly; the prompts require abstention from unsupported numeric
equivalence and exact prose excerpts in the claim ledger. Earlier responses
remain preserved for replay; a new version is not evidence that the model
follows its rules.

Version 4 explicitly requires the claim ledger in new dossiers. Missing, empty
or multiple ledgers visibly qualify the response. Formula-field roles in the
report distinguish presented originals or shared translations from decoded
stored XML text. Legacy dossiers without the explicit requirement retain their
previous behavior; these checks still do not certify prose.

Version 5 adds the executed second editorial pass, explicit reference-operator
and target-identity facts, node-qualified overview proof selection, and optional
bound scalar calculation checks. These mechanisms address observed defect
families without changing old responses, independent source oracles or corpus
expectations to make a run pass.

Focal-node proofs can use `{{SOURCE_FORMULA}}`, `{{STORED_XML_FORMULA}}` and
`{{CANONICAL_FORMULA}}` on separate lines. The program inserts only the exact
corresponding field from the sent dossier, records its pointer, role, UTF-8 hash
and lengths, and reports absence when it was omitted. A model no longer needs
to transcribe a long original formula. Raw responses keep the tokens, and
`deterministic_formula_insertions` records the inserted source separately from
the AI's claims. Delimiter-containing text uses an explicitly labelled JSON
string that decodes to the source. Other explanatory prose still needs review.
An overview selects a retained example explicitly by its exact node ID, for
example `{{SOURCE_FORMULA:c:Summary!C7}}`, with matching qualified stored-XML and
canonical tokens. A unique retained `formula_patterns` entry resolves to its
actual indexed JSON pointer; an omitted, missing or ambiguous ID resolves to
explicit absence. Unqualified focal tokens still never borrow a neighboring
example or reconstruct an unsupplied group member.
Token-only code wrappers are removed before insertion so backticks emitted by
the model cannot nest the source proof's code fences; the raw response stays intact.

## Evidence before prose

Read the target and scope before making a claim. Keep these evidence layers
separate:

- `formula_provenance.xmlText` / `formulaProvenance.xmlText`: decoded character
  data stored inside the XML formula element, without a presentation prefix.
  XML entity escapes are decoded; this is not the original XML byte sequence.
- `source_formula` / `sourceFormula`: a presentation of that original decoded
  text with a leading `=` added. The added character is not part of `xmlText`.
  `formula_text_evidence` records their distinct character counts, the prefix
  relationship and the literal canonical/source relationship. A literal
  difference does not establish semantic equivalence or inequivalence.
  The source hash/part/cell locate the original.
  A shared slave with no stored text has no original formula string; its
  translation remains derived evidence.
- `canonical_formula` / `canonicalFormula` / legacy `formula`: the expression
  used by the parser/engine. Equivalent arithmetic does not prove a verbatim
  source quotation. XML attributes, normalization and workbook hashes remain
  attached to source provenance.
- Decomposition: expression structure and any explicitly evaluated values.
  Missing evaluation is unknown; this trace is not an independent Excel oracle.
  A guarded formula can retain a saved file value while its decomposition
  records a separate scratch evaluation, including `#REF!` or `#NAME?`.
  That scratch reading does not explain how the saved cache was calculated.
  Compare each layer's origin and evaluation status before describing steps;
  a different root reading does not turn the file cache into an engine result.
- Displayed value and origin: formula result, input read, name resolution,
  external file read, saved cache, preserved volatile reading or unknown.
  Typed saved values are distinct from engine results. Disagreement establishes
  a difference, without establishing age or cause. Volatile readings are kept
  by policy, without proving the engine cannot calculate the function.
  `value_comparison.literal_text_equal` compares two strings without parsing
  them. `numeric_equivalence: not_established_from_strings` prevents treating
  separator variants as a verified magnitude agreement. A `format` comparison
  label is a heuristic; it neither parses the strings nor establishes their
  locale, numeric equivalence or mismatch cause. Requested lazy calculations
  have a separate comparison with the saved cache, rather than inheriting the
  structural snapshot's comparison.
- Geometry and topology: group membership is distinct from representative
  values and bounded samples. Directed graph edges establish observed paths;
  multiple inputs merge, multiple outputs branch. A list of patterns establishes
  no execution order or route through a named range.
- Static VBA: supplied source, source hashes, procedure boundaries, extraction
  and truncation. No execution, dynamic-call completeness or business effect is
  established by static inspection alone.

`vba.code_bounds` separates the full decoded procedure, the graph preview and
the actual dossier excerpt. For example, counts of 10,351 / 6,000 / 2,500 mean
4,351 characters omitted before the graph, 3,500 omitted from graph to dossier,
and 7,851 omitted overall. The legacy `code_omitted_characters` counts only the
graph-to-dossier step. Unknown full-source sizes remain null; inconsistent size
metadata does not produce guessed or negative omission counts. Each stage has
its own decoded-text hash. A full-source hash and line span do not make omitted
code readable. Further dossier-size omission sets the supplied-code count to
zero and removes its excerpt hash. Workbook overviews supply procedure metadata
without the procedure bodies.

The literal AST lookup facts explicitly distinguish a one-based VLOOKUP/HLOOKUP
index from the offset from the first column/row. An index of five has an offset
of four. A syntax inventory of functions is not an engine support inventory;
the AGGREGATE selector list is not a function inventory. Defined-name counts
retain their local/global scope and any filtering or omission.

`formula_facts.reference_operators` records parsed intersection (space) and union
(comma) separately. `literal_self_reference` concerns literal single-cell
references; its scope excludes ranges containing the target and dynamic
references. It is distinct from the graph's explicit `self_edges`. Multiple
incoming edges establish merging, not a self-reference or cycle. Neighbor
records carry their exact sheet, address, extent and directed edge; a group
representative's formula cannot exclude a relationship through another member.
`value_coverage.counts` categories are not `value_origin_kind` values.

`formula_facts.three_dimensional_references` identifies literal parsed references
that span sheets. It retains their exact reference and per-sheet address, but
does not resolve the workbook's intervening sheets: `sheet_count` and
`total_cells` stay unknown. Colons inside external file names or quoted string
arguments do not establish a sheet span. For such a reference, a decomposition
counter supplied without its scope is not a verified total cell count.

The workbook overview dossier has a 16,000-character bound. When source
provenance makes its examples too large, it reduces their count while preserving
the complete original/canonical evidence of each retained example. Formula,
preview, comment, VBA and name omissions are explicit. An overview may omit all
formula examples; that absence cannot support a formula claim. The separate node
dossier retains its focal source and declares any necessary size omissions.

## Atomic claim ledger

Localized node and workbook prompts request a trailing `json_claims` array for
up to twelve material claims. The block is removed from visible Markdown and
retained in the raw response and validation report:

```json
[
  {
    "statement": "The displayed reading is a saved volatile cache.",
    "classification": "fact",
    "evidence": [
      {"pointer": "/value_source", "value": "volatile"},
      {"pointer": "/value_origin_kind", "value": "volatile_cache"}
    ]
  }
]
```

Each statement must be an unchanged contiguous excerpt of the actual Markdown
prose, including any emphasis and punctuation; a paraphrase is not a match.
Evidence uses RFC6901 pointers into
the exact sent JSON, with unchanged typed values. A missing/omitted field cannot
support a fact. A contradicted binding or ungrounded fact visibly qualifies the
response; malformed and oversized blocks are recorded, without rewriting the
original response. Inferences and unknowns remain indeterminate. New dossiers
set `review_protocol.claims_required: true`; missing, empty or multiple ledgers
are review issues. An older dossier without that requirement can record a
missing ledger as `not_supplied`. A generic `json` fence is not `json_claims`.

`supported` means **only** that the claimed typed value exists at the pointer.
For example, a false sentence about an engine recalculation could still bind
correctly to `value_source: file`. The deterministic check cannot validate that
sentence's meaning. Formula quotation checks remain lexical and separately
identify their supporting fields. Independent review must check both the prose
and the ledger, rather than trusting the generator's classification.

A claim may additionally supply `calculation={operator,operands,result}`, with
two numeric operands bound by exact pointer/value pairs to the sent dossier.
The checker verifies `+`, `-`, `*`, terminating decimal `/` and numeric comparison
results using a fixed Decimal context. Booleans, parsed strings, missing operands,
division by zero and nonterminating division are not coerced into proof. This is
a bounded scalar consistency check, not Excel rounding/coercion, independent
evaluation or verification of surrounding prose. A wrong copied operand is
distinguished from a wrong predicate result. The RAW claim is always retained.

## Reproducible comparison and independent review

Preserve each full run separately under ignored `validation_screenshots/`.
Record requested/resolved model, parameters, prompt/dossier/workbook/image hashes,
usage, finish reason and raw responses. A hosted-model seed or low temperature
records replay conditions; it does not guarantee identical output.

Text validation reports retain `sent_dossier.json_text`, its UTF-8 SHA-256,
character count and byte count. This is the exact fitted JSON portion supplied
in the user prompt, not a reconstructed graph/context, HTTP request or system
prompt. Its hash excludes the prompt prefix, provider parameters and images.
`factual_review.dossier_sha256` is a separate, sorted parsed-JSON fingerprint
used for binding checks; it must not be called the exact serialized request hash.

`OpenAICompatProvider(streaming=True, ...)` explicitly assembles only content
deltas, requests terminal usage, and requires `finish_reason: stop` for text and
images. The default remains a single non-streaming response. Empty or interrupted
streams, length limits, mixed choices and inconsistent response identities fail;
partial text never becomes accepted documentation. The stream closes before a
completed response receipt is recorded. SDK retries do not restart a stream that
has already begun; failed requests and their available or estimated usage remain
separate from completed generation.

The baseline review from 9 October 2026 found 33 discrepancies across 35 sampled
cards, three overviews and fifteen image descriptions (twelve unique images).
Generation coverage passed separately. Preserve that report unchanged as the
before result; review new responses independently, without carrying old verdicts
over to changed text.

1. Run the required full acceptance from [manual validation](manual_validation.md).
   Compare Sales in French and English (all four formula/group cards in each
   language), Stress in English, all three overviews and all screenshot
   descriptions. Reuse Stress sampling seed `20260912` and the 27 target IDs from
   the baseline ignored `factual-review.json`. Changing fixtures requires a source
   difference report; identical target IDs do not establish identical formulas.
2. Independently read immutable workbook XML, typed cached values, graph edges,
   decomposition and original images. Reconcile source/canonical/shared text,
   value origin, literal text differences, name scope, lookup offsets, group
   geometry and topology. Verify external disk reads from their provenance, never
   from the word “external” or a known value alone.
3. Inspect actual pixels. Locate headers and align values with their columns;
   sample legible top/middle/bottom regions. Record visible, ambiguous and
   off-image details with approximate boxes. Dense images are not evidence of
   blank data, clipping, merged cells or hidden-sheet state. Keep numerical and
   business causes outside image-only evidence.
4. Record atomic findings with response excerpt, target, evidence pointer/image
   region, verdict (`supported`, `contradicted`, `indeterminate`), severity and
   correction. Keep a separate category for source-provenance limits, such as
   missing original XML in an older payload. Report unreviewed cards explicitly.
5. Compare before/after defect categories and exact repeated claims, not only
   the total finding count. Abstention can reduce invented facts while reducing
   explanatory coverage; record both. Report ledger compliance, binding
   contradictions, malformed blocks and independent prose/vision findings
   separately. Successful generation and supported bindings are not factual
   acceptance.

The adversarial regression tests replay a local name presented as a disk read,
volatile/cache origin confusion, a comma/period string mismatch, one-based lookup
offset confusion, wrong topology attribution, and canonical text falsely bound
as original. Other tests cover omissions, type coercion, malformed/oversized
ledgers and false prose with a correctly supported binding. These tests verify
the controls' scope and limits; they are not model-quality measurements.
