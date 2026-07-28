# Quick start

## Analyze a workbook

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
