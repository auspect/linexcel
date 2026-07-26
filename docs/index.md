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
