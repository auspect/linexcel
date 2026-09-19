# Acceptance evidence and remaining limits

Status on 19 September 2026: **review complete with known limits**. Complete
generation, source checks and desktop/mobile investigation passed their stated
checks. Independent factual review found remaining errors in generated prose
and image descriptions; it does not award an unqualified factual pass.
The reviewed runtime and generated artifacts are bound by fingerprints.

## Execution and API contract

The default API and CLI run analysis in a supervised worker, with a
120-second analysis budget and a 2,048 MiB memory limit. Windows attaches a
Job Object before the suspended worker starts initialization and limits
aggregate committed memory. POSIX applies a per-process virtual-address
limit before importing linexcel. These are different memory measures,
not interchangeable RSS limits. Process-tree cleanup is bounded and can
extend the return time beyond the analysis deadline.

Execution metadata distinguishes completion, timeout, cancellation,
memory-limit failure and native crash. A crash is not automatically
attributed to memory pressure. Interrupted work retains only completed
source evidence; partial native calculations are not presented as completed
recalculations. The lexical formula guard prevents known dangerous cell
expressions from reaching native parsing/import, but does not prove that
all native crashes are impossible.

An isolated result has `result.engine is None`. This means no live native
handle crosses the process boundary, not that the graph contains no engine
results. `valueSource`, cache comparisons, execution metadata and coverage
describe what was obtained. Explicit `ExecutionPolicy(isolated=False)`
restores the live-engine API without the hard process time/memory protection;
there is no automatic fallback to that mode.

Source-file reading, AI requests, screenshot rendering and HTML export are
outside the analysis budget. Dynamic references can require work beyond a
static target trace. Coverage counts graph nodes and records known omissions;
it is not an exhaustive count of independently verified workbook cells.
See [execution and compatibility](execution.md) for the detailed contract.

## Established automated evidence

The frozen candidate passed **1,098 tests**, with three skipped and two slow
tests deselected. The separate slow-test run passed those **two tests**.
These are two recorded runs, not a single uninterrupted 1,100-test run.
The three skips remain exclusions from executed coverage.

Independent source review found no blocker in its reviewed scope, including
name-risk annotations, source evidence, provenance and the changed viewer
behavior. It also exercised worker cleanup and targeted regressions. Source
clearance does not establish that the final generated explanations or
screenshots are factually correct.

## Numerical evidence

The runtime diagnostic contains **40 synthetic observations**, covering
both the 1900 and 1904 date systems. LibreOffice 26.2.4.2 recalculation of
the exact diagnostic input copies confirmed **40/40 expected results**.
Formualizer 0.9.3 matched 21 and differed on **19**:

- Twelve mixed number/text comparison or comparison-guard observations.
- One date-versus-number result from date/serial subtraction.
- Six observations involving numeric or text defined-name constants,
  including sheet-local shadowing.

The observed values and raw types remain visible. Potentially affected
formulas and represented dependents receive semantic-risk annotations;
this does not repair their calculations or prove that every marked result
is wrong. Constant-name diagnostics require matching source definitions
and respect local scope. A generic `#NAME?` does not by itself establish
that a source name is missing or unsupported. Cache agreement is not an
independent correctness oracle, and LibreOffice is not a certificate of
Microsoft Excel equivalence.

A **separate versioned extension contains 42 observations**: 21 cases in
both epochs covering absent cells, empty strings, booleans, errors,
dynamic references, names and local bindings. It overlaps the runtime
diagnostic and must not be added to it as 82 distinct cases. LibreOffice
confirmed 34 provisional expectations; six of those expose the same
constant-name limitation. Eight expectations remain unconfirmed: six
boolean-comparison observations and two guarded references to a missing
sheet. The three-way expected/native/LibreOffice distinction is retained.
These unresolved cases do not introduce additional runtime mismatch claims.

Cycles are covered by explicit iteration and nonconvergence regression
contracts, not by choosing an arbitrary scalar fixed point as an oracle.
Full operator/type combinations, dynamic dependency discovery, all name
expressions and complete Excel function compatibility remain outside the
demonstrated coverage. See [engine safety and capability evidence](engine-safety.md)
for case scope, versions and interpretation.

## Measured performance scope

A local before/after experiment used the same synthetic workbook with
100,000 cells and 20,000 formulas, with three repetitions per version.
Median cache-loading time changed from 82.00 to 67.23 ms, and error-overlay
extraction from 21.31 to 5.76 ms. The detailed note records ranges, sample
standard deviations, runtime versions and host identification.

**Before/after memory consumption was not measured.** Native import was
not optimized by this change. These short local stage measurements do not
establish an end-to-end speedup, lower peak memory or a global performance
guarantee. Isolation adds startup overhead, especially for small inputs.

## AI and factual limits already observed

The final run's bounded review of the synthetic Sales outputs found
improved separation of source definitions, cached values and engine results.
The inspected core formula results and decomposition showed no new
discrepancy. This is a sample finding, not whole-workbook numerical
certification.

Generated prose still contains material errors. One explanation transcribes
a neighboring formula's division by 1,000 as division by 100; another
describes a branching dependency graph as a linear chain. Other remaining
issues include confusing a named constant with a recalculated formula,
unsupported claims about cache age or gaps in a complete group, inconsistent
decimal-separator descriptions and inconsistent illustrative concatenation.
Successful generation has therefore **not established factual correctness**.

The final Stress review compared 17 selected formulas directly with workbook
XML; all matched the recorded source formulas. It confirmed useful explanations
of constant-name limitations, unevaluated LET steps, AGGREGATE selectors,
cycles and dynamic-reference omissions. Remaining prose incorrectly infers
that a source table definition is absent, describes an opaque external
reference as a successful read, and invents gaps in a complete eight-cell
group. Image descriptions also invent columns or borders and misassign error
tokens between columns. These are recorded defects, not verified findings
about the workbook.

Together, the independent factual reviews inspected 25 cards, all three
overviews, all 15 sheet descriptions and all 12 distinct images. The node
selection was predetermined with seed 20260919 and covers formula-bearing
sheets and difficult cases. It does not certify every generated card or
every row of the tall rendered sheets.

AI documentation and image descriptions remain assistive explanations.
Source formulas, graph evidence, provenance and actual screenshots must be
used to check their claims. Built-in AI HTTP timeouts and token budgets are
separate from the analysis budget; token limits are checked between requests
and an already submitted response can exceed the remaining allowance.
Custom providers retain responsibility for their timeout behavior.

## Final acceptance record

The final local generation run completed with `validation.json` status
`passed`, using both generated workbooks without a node limit. It produced
75 node cards (four Sales cards in French, four in English and 67 Stress
cards in English), three workbook overviews and 15 image descriptions.
There are 12 unique sheet PNGs; the empty-sheet exemption was checked
against source XML. Generation took approximately 19 minutes 23 seconds
and consumed 233,117 tokens across 93 requests. This is a record of this
run, not a generation-time or accuracy guarantee.

The local alias `qwen3.8:linexcel-32k` uses the Qwen3.8 weights with a
32,768-token context. The run used an 8,192-token output limit, one worker,
seed 20260919 and a 1,000,000-token budget. The 61 source/configuration
fingerprints recorded in the manifest still match the candidate. These
generation facts do not replace review of the exact outputs.

Six browser sessions visited 18 tabs, 32 sheet views and 48 node selections
at 1440×900 and 390×844, producing 242 screenshots. The UX agent inspected
58 captures, including every tab and sheet at both sizes. A separate reviewer
inspected 16 captures and all ten additional Qwen dashboard descriptions;
those ten calls consumed 21,064 tokens, separately from the workbook run.
No JavaScript errors, external network requests or page/panel width overflows
were found in these sessions. Independent interaction probes confirmed that
mobile Options remains inside the viewport, scrolls internally, keeps keyboard
focus inside while open and restores focus with Escape without losing selection
or camera state. Graph controls remain scoped to the graph tab.

The dense 96-node, 77-edge graph is a map at whole-graph scale, not readable
cell documentation. Local exploration makes a selected neighborhood readable
and discloses hidden-node counts; it does not eliminate every crossing edge.
Tall sheet images require zoom. These checks are not a full accessibility audit.
The model sometimes calls normal internal scrolling or below-the-fold content
clipping; those claims were checked against actual geometry rather than accepted
as UI defects.

| Required evidence | Current acceptance record |
| --- | --- |
| Frozen source identity matches the analyzed and rendered candidate | 61 recorded source/configuration fingerprints unchanged. |
| Both workbooks and complete language runs | Generation passed: Sales FR/EN and Stress EN. |
| Nonempty, complete node cards and workbook overviews | Manifest passed: 75 cards and three overviews; factual review remains separate. |
| Sheet captures and their AI descriptions | 12 unique PNGs and 15 descriptions reviewed; one source-verified empty-sheet exemption. Factual defects retained above. |
| Every dashboard tab and graph at 1440×900 and 390×844 | Six sessions completed; UX and independent visual review found no blocking interface defect in the checked scope. |
| Bounding boxes, clipping and notice scope across tabs | Checked across tabs and viewports, including independent mobile focus/scroll probes. |
| Reproducible nodes across sheets and edge cases, checked against source and provenance | 25 predetermined cards reviewed; source and decomposition kept separate from incorrect AI prose. |
| Actual image content compared with generated descriptions | All sheet descriptions and ten dashboard descriptions reviewed against images; known errors remain. |
| Requests, tokens, model identity and run completion status | 93 requests, 233,117 tokens, local Qwen3.8 alias above; generation status `passed`. |
| Independent final acceptance decision | Review complete with known limits; no unqualified factual AI or Excel-compatibility pass. |

No private workbook, corpus identity, prompt transcript, raw model response
or private audit report is included here. Local evidence is retained
separately. The immutable generation manifest retains its historical
`manual_review: pending` marker; a separate final aggregate binds the completed
reviews, screenshots, model outputs, source and committed-file fingerprints.
The aggregate, not generation success alone, records the final review decision.
