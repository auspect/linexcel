# HTML export

## Standalone HTML (offline)

```python
result.save_html("lineage.html")
```

Fully offline — Cytoscape.js embedded, no internet needed.

## Focus on a worksheet

For a workbook with many sheets, use the sheet selector in the **Graph** tab
to focus on one worksheet at a time:

1. Generate a report with `linexcel analyze workbook.xlsx` and open the HTML file.
2. Select `Summary` in the sheet selector (substitute a sheet from your workbook).
3. Inspect the formulas on that sheet. Directly connected nodes from other
   sheets remain visible and dimmed; unrelated nodes are hidden.
4. Choose all sheets to restore the broader graph, subject to any other filters.

For example, if `Summary!B4` reads `Sales!D10`, the `Sales` node remains as
context when you select `Summary`. This keeps a large graph easier to navigate
without hiding immediate cross-sheet connections. It does not display every
transitive dependency automatically.

The selector changes the report's display after analysis. It does not reduce
workbook loading, change the Excel file, or exclude sheets from analysis.
To restrict analysis to particular outputs and their upstream lineage, use
[`--target` or `targets=`](quickstart.md#trace-specific-output-cells).

## Investigating a result

Search returns a list of matching references, formulas and labels. Choose
whether to search the visible graph or all nodes in the analyzed workbook;
the latter includes nodes hidden by the current filters. It does not search
cells omitted from analysis. Use the arrow keys to choose a result, Enter to
open it and Escape to close the results.

The selected-cell panel keeps the formula, value provenance, decomposition,
connections and AI explanation together. Copy actions preserve the exact
reference or formula. Previous/next selection history and selection links
help return to a finding. “See sheet” opens its presentation context, and the
return action restores the graph selection, filters and camera.

On dense graphs, start with the immediate neighborhood and expand when
needed. Hidden-connection counts and the connection list make the reduced
view explicit. Coverage categories open their corresponding nodes; their
counts describe graph nodes, including grouped formulas, not individual
workbook cells. “Export visible view” saves a PNG of the current graph view.

Graph layout, zoom and filtering controls belong to the Graph tab. On narrow
screens, additional options are in the Options dialog. Formula values and
source evidence remain separate from AI prose, whose factual claims still
require verification.

## With AI documentation

```python
docs = result.document(base_url="http://localhost:11434/v1", model="qwen3.8")
result.save_html("lineage.html", docs=docs)
```

## With screenshots

```python
screenshots = result.save_screenshots("screenshots/")
result.save_html("lineage.html", screenshots=screenshots)
```

One image per sheet, keyed by sheet name, so the **Sheets** tab shows each one
under the sheet it renders. See [screenshots](context.md#screenshots) for the
flat print-page alternative.

## In a notebook

```python
# In marimo or Jupyter, just output the result object
result
```

Renders an isolated iframe with the interactive graph.
