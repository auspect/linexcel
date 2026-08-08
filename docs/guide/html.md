# HTML export

## Standalone HTML (offline)

```python
result.save_html("lineage.html")
```

Fully offline — Cytoscape.js embedded, no internet needed.

## With AI documentation

```python
docs = result.document(base_url="http://localhost:11434/v1", model="llama3.1")
result.save_html("lineage.html", docs=docs)
```

## With screenshots

```python
screenshots = result.save_screenshots("screenshots/")
result.save_html("lineage.html", screenshots=screenshots)
```

## In a notebook

```python
# In marimo or Jupyter, just output the result object
result
```

Renders an isolated iframe with the interactive graph.
