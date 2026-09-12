# Reliability study — 2026-09-12

This follow-up to the integrated branch review uses an immutable baseline and
a second source snapshot. Workbook files, prompts containing their contents,
screenshots and detailed reports remain in a private directory outside Git.
Only generic fixes and synthetic regressions are included in the pull request.

## Findings and action plan

| Priority | Reproduced weakness | Action |
| --- | --- | --- |
| P1 | Date cells imported using the wrong workbook calendar | Set the engine date system before import; verify both epochs and targeted runs |
| P1 | Circular recovery can expose intermediate values as recalculated results | Honor declared iteration limits and keep unverified cycles/dependents cache-only |
| P1 | Empty XML cells cause quadratic scanning and can steal following formulas | Exclude self-closing openings and skip sheets without formulas |
| P1 | Optional sheet context materializes millions of styled cells | Stream bounded previews for large worksheet XML; display omissions |
| P2 | AI calls receive cached values labeled as computed values | Preserve sources, disagreements, partial adjacency and unevaluated steps |
| P2 | Generated prose infers missing inputs or reasons for discrepancies | Tighten node and overview instructions, replay the same sample, retain failures |
| P2 | Dense neighborhoods and mobile panels obstruct inspection | Explicit neighbor framing, usable mobile controls and browser regressions |
| P2 | Markdown and error styling misrepresent content | Render supported Markdown safely and distinguish errors from successful steps |

Work was divided between backend and viewer agents, with a dedicated reviewer
who did not implement the fixes. The coordinator maintained the private study,
addressed XML/context/documentation problems, and integrated feedback. Review
also corrected a methodology mistake: context absent from exported JSON can
still be embedded in HTML, so dashboard coverage must inspect the actual HTML.

## Method

- 31 selected files: 28 XLSX and three legacy XLS rejection cases. The selection
  combines a small established corpus with a size-stratified sample from a
  larger local archive, including its two largest XLSX files. Seed: 20260912.
  This is not an exhaustive execution of the archive.
- Each file runs in a separate process, with a 90-second process limit,
  10-second decomposition budget and 1,800 MiB process memory limit. Windows
  peak working set is measured on the actual interpreter, not its launcher.
- Inputs are hashed and checked unchanged. Baseline commit and source snapshot
  fingerprints accompany results. Invalid harness trials remain separate.
- LibreOffice recalculates copies of 17 XLSX files under separate profiles.
  Original caches, recalculated values and LibreOffice caches are compared
  separately. Constants and cache-only results cannot inflate engine agreement.
- The independent reviewer checks 20 sampled nodes across 11 reports against
  worksheet XML: 65 precedent occurrences and 75 decomposition steps.
- Local Ollama generates 17 node cards, four workbook overviews and three image
  descriptions. The same node IDs are replayed after changes. Actual prompts,
  model identity, seed, token counts, image hashes and responses stay private.
- Browser inspection covers graph, overview, sheet context, visual previews,
  random nodes and high-degree neighborhoods on desktop and mobile. Synthetic
  browser tests cover regression behavior and safe Markdown rendering.

## Observed results

All 28 selected XLSX files export after the changes, versus 25 before. The three
XLS rejection cases remain explicit. No selected input was modified.

| Anonymous case | Baseline | After | Interpretation |
| --- | ---: | ---: | --- |
| A | stopped at 90 s | 3.3 s | Empty-cell XML backtracking removed |
| B | stopped at 90 s | 3.3 s | Empty-cell XML backtracking removed |
| C | stopped at 90 s | 13.1 s | 14.8 million XML cells but only ten formula elements |
| D | 17.7 s / 1,246 MiB peak | 6.4 s / 538 MiB peak | Rich context no longer materializes a large data worksheet |

These cohort timings include other local work and are not portable performance
guarantees. Independent phase probes corroborate the causes: a native import
finished in 0.1 seconds before Python scanning exceeded 30 seconds; another
workbook's optional context alone took 10.9 seconds and added roughly 791 MiB
to its current working set. Synthetic regressions test correctness of the
problematic structures rather than fragile machine-specific timing thresholds.

Triangulation after changes finds 1,487 sampled observations where all three
readings agree, 111 where original and LibreOffice agree but the engine differs,
and two other disagreements. Another 15 observations lack a LibreOffice cache;
seven are not recalculated. This is sampled coverage, not an overall accuracy
score: grouping reduces observations, caches may be missing, and implementation
differences require case-level investigation.

Independent XML review found no factual dependency or decomposition defect in
its limited 20-node sample. It did find incorrect generated prose and visual
descriptions. Successful HTML rendering is never counted as factual validation
of those descriptions.

## Remaining work

1. Expand the compatibility suite for legacy function names. In formualizer
   0.9.3, `NORMSDIST(0)` and `NORMDIST(0,0,1,TRUE)` return `#NAME?`, while their
   modern counterparts return 0.5. Do not classify every `#NAME?` as unsupported:
   genuinely invalid names use the same error. Alias support needs semantic
   tests, including guards and dependent formulas.
2. Investigate typed date arithmetic separately: subtracting a serial from
   `DATE(...)` can return a typed date representing zero instead of numeric zero.
3. Preserve inter-engine disagreement on iterative models. Different starting
   values or evaluation order can produce different converged fixed points;
   neither a cache nor engine telemetry proves a unique Excel result.
4. Add claim-level checks for generated documentation and calibrated visual
   evaluation. The tested vision model can invent columns and row counts. Its
   text remains an aid to inspection, with the original image alongside it.
5. Extend stratification by formula families, macros, dynamic references and
   external links. Static traces and sampled group values are not exhaustive
   proofs of workbook behavior. A dense whole-workbook graph still needs search,
   sheet filtering or neighborhood focus for detailed reading.
