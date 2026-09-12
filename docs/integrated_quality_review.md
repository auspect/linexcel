# Integrated quality review — 2026-09-12

## Scope and branch ancestry

The requested review covers `improve/robustness-formualizer-093`, then
`improve/targeted-evaluation`, then `improve/viewer-design`. The first two are
stacked; viewer-design branches independently from main at `e9c2ac2`.
`codex/integrated-quality-review` integrates all three without changing main.

## Initial findings

| Priority | Area | Reproduction / weakness |
| --- | --- | --- |
| P1 | Evaluation | A quarantined formula without a cache becomes blank; `=blocked_cell+5` can report 5 as a recomputed result. With a cache, the blocked cell can incorrectly claim engine provenance. |
| P1 | Targeted scope | Building all defined names can recalculate an unrelated formula through its name. Sparse targets are scanned through dense rectangles, spending the sheet budget on unrelated cells. |
| P2 | Diagnostics | `_longest_dep_chain` never visits its root and returns zero; tests covered only chains collapsed into a single group. |
| P2 | Targets | Naive comma splitting rejects legal quoted sheet names; empty targets create inconsistent mode selection. |
| P2 | Comparison | Opposite booleans and mixed value types can report agreement, hiding real differences from the new viewer filter. |
| P2 | Trace fidelity | INDIRECT/OFFSET can evaluate cells absent from the engine's static trace. Truncated traces do not establish that omitted cells were never evaluated. |
| P2 | Viewer | Sheet filter style overrides conflict with kind/difference classes; search can select invisible nodes. |
| P2 | Viewer performance | Searching for 400 matching nodes triggers 801 complete node scans in Edge because every selection event updates all labels. |
| P2 | Validation | Viewer tests inspect source strings without exercising browser interactions; README screenshots need recapture after changes. |

## Action plan and ownership

1. Integrate branches and use the locked formualizer 0.9.3 environment.
2. Backend agent: bound targeted scans to reachable cells, isolate graph context,
   and preserve honest values/provenance through quarantine and downstream reads.
3. Coordinator: fix target parsing, linear chain diagnostics, comparison semantics,
   and accurate statements about static/partial targeted lineage.
4. Viewer agent: compose filters, restrict search to visible nodes, coalesce label
   updates, and add browser tests for interactions and responsive layout.
5. Dedicated read-only reviewer: independently reproduce defects and review the
   final changes, with feedback returned to the implementers.
6. Run regression tests, lint, formatting, types, browser checks and packaging;
   recapture and inspect README screenshots; finish changelog and publish a PR.

## Validation notes

The initial local environment contained formualizer 0.8.4, incompatible with the
integrated lockfile. Its seven failing targeted tests are not used as evidence of
a 0.9.3 regression. Validation after `uv sync --locked --extra screenshots` uses
formualizer 0.9.3. Browser validation uses installed Edge on Windows and synthetic
workbooks; it does not require Excel, a network API or AI-generated documentation.

## Results

- Full integrated suite on Windows / Python 3.13.12 / formualizer 0.9.3:
  **767 passed, 3 skipped**, including the eight Chromium browser cases
  (139.91 seconds). The skipped items are existing documentation examples.
  Ruff lint, formatting across 70 Python files, `ty check` for `src/`, and
  `git diff --check` all passed.
- Dedicated reviewer: final feedback favorable for backend and viewer; no
  remaining blocking finding identified. The reviewer independently ran 47
  backend/contract regressions and additional full/targeted counterexamples.
- Viewer agent: 183 viewer tests passed, including eight actual browser cases;
  screenshots inspected for desktop/mobile, light/dark and the dense grid.
- Controlled sparse-reader comparison: five targets on a 5,000 × 16,384 sheet
  requested 4,997,120 cells before and five after, in five calls in both cases.
  This measures the formula reader, not end-to-end analysis time: package loading,
  cached-value loading and structure inspection still have workbook-wide costs.
- Dense search previously made 801 synchronous `cy.nodes` calls for 400 matches;
  the regression test now requires fewer than 20, with one deferred selection
  update. It does not impose a machine-dependent wall-clock threshold.
- Wheel and sdist built successfully; both imported successfully in isolated
  environments and contained the embedded offline viewer assets.
- README's four screenshots were recaptured from a fresh deterministic analysis
  of `validation_demo.xlsx`, preserving its existing fixture AI documentation
  and workbook context. No new model call was used. Source manifest verified.

## Feedback incorporated during independent review

- Replacing an unavailable value with an engine error was insufficient:
  IFERROR/ISERROR could absorb it and manufacture a plausible value. The backend
  now distinguishes engine limitations from real spreadsheet errors and suppresses
  unsupported dependent readings and decompositions, retaining file caches.
- A deferred dependency graph initially forced the conservative fallback for all
  formulas. Building the evaluation plan before the dependent trace preserves
  independent calculations without recalculating the blocked formulas.
- Error objects and serialized dates need normalization before mixed-type
  comparisons; explicit regression cases protect both representations.
- Group comparison uses `groupCachedAgreement`; `cachedAgreement` continues to
  describe the representative cell. The browser test checks the mixed sample case.
- Visual inspection caught a near-empty canvas for 400 disconnected nodes in
  flow layout. Those nodes now use a compact grid, with an aspect-ratio check.

## Reproduce browser validation

```sh
uv sync --locked --extra screenshots
uv run playwright install chromium
uv run pytest tests/test_viewer_browser.py
```

On Windows, `LINEXCEL_BROWSER_CHANNEL=msedge` selects an existing Edge installation.
`LINEXCEL_VIEWER_ARTIFACTS` can capture the states exercised by the tests.
The CI `viewer-browser` job installs Chromium explicitly, so these tests cannot
silently disappear from the required checks because Playwright is optional.

## Known limits

- Targeted lineage uses the engine's static trace. Dynamic references and budget
  truncation can omit dependencies that the engine nevertheless evaluates.
- If the dependent closure of an unavailable deep formula cannot be established,
  formula readings conservatively use stored values and the report warns. Dynamic
  address formulas require the same caution when an unavailable input exists.
- Group comparisons cover a bounded sample, not an exhaustive Excel parity audit.
- Local browser screenshots cover Chromium/Edge, desktop/mobile and two themes;
  they do not establish Firefox/Safari compatibility or performance on every
  workbook topology.
