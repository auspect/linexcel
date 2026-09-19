# Quick start

Prefer a shell? The same analysis runs from the
[command line](cli.md), with no install.

## Analyze a workbook

Analysis is isolated and bounded by default. See [execution and compatibility](execution.md)
for budgets, incomplete results and explicit access to the live engine.

```python
from linexcel import analyze

result = analyze("workbook.xlsx")

# Interactive graph in marimo / Jupyter
result

# Standalone HTML viewer
result.save_html("lineage.html")

# Stats
print(result.stats)  # {totalFormulas, totalNodes, ...}
print(result.warnings)  # list[str]

# Find nodes
result.find("Summary")
result.precedents("c:Summary!B4")
result.dependents("c:Summary!B4")
```

## Tune analysis limits

The Python API accepts four optional non-negative integer ceilings. `None`
keeps the default; `0` is a real ceiling. Invalid values fail before the source
file is read.

```python
result = analyze(
    "workbook.xlsx",
    max_cells_per_sheet=2_000_000,
    max_nodes_per_sheet=600,
    max_chain_depth=32,
    max_dense_cells=5_000_000,
)
print(result.graph["meta"]["analysisLimits"])
```

| Parameter | Default | What it bounds |
| --- | ---: | --- |
| `max_cells_per_sheet` | 64,000,000 | Cached cells retained and formula cells visited, in row-major order per sheet. A partial last row is retained. |
| `max_nodes_per_sheet` | 400 | Detailed formula groups. Excess groups share a `misc` node; inputs, names and other node kinds are additional. |
| `max_chain_depth` | 24 | Precedent depth during fallback value recovery. An incomplete recovery leaves affected values unavailable or uses their file cache, without a calculated decomposition. |
| `max_dense_cells` | 20,000,000 | Declared sheet size allowed through the dense cache reader. Larger primary workbooks use the lazy reader; larger external workbooks are refused. |

In targeted analysis, the extraction ceiling counts cells in the traced closure,
including distant cells without charging for gaps. If that closure exceeds the
ceiling, analysis raises an error instead of returning an incomplete lineage.
Cache retention still follows row-major order, so the cache of a distant target
may be unavailable. External files also have an independent 200,000-cell ceiling
per sheet; a file that exceeds it is refused rather than read partially.

These ceilings do not bound native engine recalculation, total memory, AST depth
or overall runtime. Truncation and fallback warnings must still be inspected.
`linexcel.analyzer.inspect_workbook(data, **limits)` accepts the same options;
its `ceilings` and `recoveryDepth` fields describe the planned limits.

## JSON export

```python
result.save_json("lineage.json")
# or
json_str = result.to_json(indent=2)
```

## Screenshots (optional, requires LibreOffice and Poppler)

Available on Linux, macOS and Windows — see
[Workbook context & screenshots](context.md#screenshots) for the install
command on your platform.

```python
screenshots = result.save_screenshots("screenshots/")
result.save_html("out.html", screenshots=screenshots)
```
