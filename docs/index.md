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

## Guide

- [Quick start](guide/quickstart.md) — analyse a workbook, explore it, export it
- [HTML export](guide/html.md) — the standalone offline report
- [Workbook context & screenshots](guide/context.md) — what a reader sees, not
  only what the file computes
- [Choosing an AI provider](guide/providers.md) — local runtime, hosted
  gateway, or your own callable; none is chosen for you
- [AI documentation](guide/ai.md) — provable cards, token accounting, budgets
- [Languages](guide/languages.md) — nine, for both the prompt and the interface
- [Data handling](guide/data-handling.md) — what leaves the machine, and when

