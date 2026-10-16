# linexcel

Data lineage analysis for Excel workbooks.

Extracts every formula, groups stretched patterns (R1C1 canonicalization), builds a dependency graph (cells, ranges, defined names, VBA), decomposes composite functions with step-by-step evaluation, and optionally documents calculations via AI.

## Install

```bash
uv add linexcel
# AI documentation (optional)
uv add linexcel[ai]
```

## Quick start

```python
from linexcel import analyze

result = analyze("workbook.xlsx")
result.save_html("out.html")
print(result.stats)
```

See the [Quick start guide](guide/quickstart.md) for more.

## Or without writing any Python

```bash
uvx linexcel analyze workbook.xlsx        # -> workbook_lineage.html
```

See the [command line reference](guide/cli.md).

## What would you like to investigate?

| Goal | Start here |
| --- | --- |
| Explore one sheet in a large report | Use the [Graph sheet selector](guide/html.md#focus-on-a-worksheet); directly connected cross-sheet nodes remain visible. |
| Explain a particular total | [Target its output cell](guide/quickstart.md#trace-specific-output-cells) and trace its upstream calculations. |
| Find what an input feeds | [Look up nodes and their dependents](guide/quickstart.md#follow-inputs-and-assess-downstream-impact) in a whole-workbook analysis. |
| Resolve linked workbook inputs | Pass `--refs-dir ./linked`; see [lineage coverage](guide/coverage.md#other-workbooks). |
| Share findings without installing a viewer | Export a [standalone offline HTML report](guide/html.md). |
| Add explanations from your chosen model | Enable [optional AI documentation](guide/ai.md). |
| Browse imported source cells | Open the embedded [Datasette data browser](guide/datasette.md) in `serve`. |
| Read VBA in a workbook or add-in | Use [static VBA reading](guide/macros.md); inspect source and provenance without execution. |

## Guide

- [Quick start](guide/quickstart.md) — analyse a workbook, explore it, export it
- [Command line](guide/cli.md) — the same, from a shell, with no install
- [Lineage coverage](guide/coverage.md) — what is in the graph, and what is not
- [HTML export](guide/html.md) — the standalone offline report
- [Workbook context & screenshots](guide/context.md) — what a reader sees, not
  only what the file computes
- [Source data browser](guide/datasette.md) — read-only worksheet tables and provenance
- [Static VBA and add-ins](guide/macros.md) — modules and procedures without execution
- [Choosing an AI provider](guide/providers.md) — local runtime, hosted
  gateway, or your own callable; none is chosen for you
- [AI documentation](guide/ai.md) — evidence-based cards, quotation checks, token accounting and budgets
- [Languages](guide/languages.md) — nine, for both the prompt and the interface
- [Data handling](guide/data-handling.md) — what leaves the machine, and when

## Maintenance and evidence

- [Continuous improvement plan](continuous_improvement_plan.md) — priorities,
  issue evidence gates, specialized AI review and Markdown cleanup decisions
- [Open-source package research](package_research.md) — optional trials,
  including Excelize display formatting
- [Local acceptance](manual_validation.md) — complete generation and independent
  factual and visual review on the current working tree
- [Specialized factual AI review](ai_factual_review.md) — typed claim evidence,
  independent source/image review and reproducible before/after comparison
