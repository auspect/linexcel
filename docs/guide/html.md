# HTML export

## Standalone HTML (offline)

```python
result.save_html("lineage.html")
```

Fully offline — Cytoscape.js embedded, no internet needed.

## With AI documentation

```python
docs = result.document(base_url="http://localhost:11434/v1", model="laguna-xs-2.1")
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
