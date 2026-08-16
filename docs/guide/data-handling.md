# Data handling

Analysis is entirely local. `analyze()`, `save_html()`, `save_json()` and
`save_screenshots()` never open a network connection, and the HTML report is
self-contained — it works from a `file://` URL with no internet at all.

Only the optional AI documentation sends anything, and only to the provider you
configured yourself. Nothing is configured by default: a call without
`base_url=` or `provider=` raises `AiDocError` instead of picking an endpoint.

## What each call sends

| Call | Sent |
| --- | --- |
| `document()` | Per node: formula, computed value, precedent/dependent labels, formula decomposition, stretched-group extent, and extracted VBA code |
| `document_workbook()` | Sheet statistics, the largest formula patterns, defined names, VBA procedures, unresolved references, analysis warnings |
| `document_workbook()` with `include_context=True` *(default)* | **Also** the first rows and columns of every sheet, cell comments and their authors, merged ranges, frozen panes, hidden columns |

That last row means cell *contents* leave the machine, not only formulas and
structure. It is on by default because an overview written without them is not
worth reading — the titles, labels and comments a reader actually sees are what
make it a description of the file rather than of a dependency graph.

Two ways to narrow it:

```python
# Keep cell contents local, still send the lineage
overview = result.document_workbook(base_url=..., model=..., include_context=False)

# Or keep the whole run local — a local runtime receives everything, and it
# never leaves your machine
overview = result.document_workbook(
    base_url="http://localhost:11434/v1", model="laguna-xs-2.1"
)
```

Sheet screenshots are never uploaded. They are rendered for a human to look at
in the report; the model is given the same facts in text, read from the file by
`openpyxl`.

## Where it goes

| Configuration | Destination |
| --- | --- |
| Nothing configured | Nowhere — `AiDocError` is raised and the message lists the options |
| `base_url=` on a local runtime (Ollama, vLLM, LM Studio) | Your own machine |
| `base_url=` on a hosted endpoint or gateway | That operator, under their terms — read them |
| `provider=` | Wherever your own callable sends it |

linexcel takes no position on which is acceptable, because only you know what is
in the workbook. Do not enable AI documentation for a file whose contents must
stay local unless the provider you configured satisfies that requirement.

## Credentials

API keys are read from the argument you pass or from the environment
(`LINEXCEL_AI_API_KEY`, `OPENAI_API_KEY`); they are never written to the report,
the JSON export, or any log. The generated HTML contains the graph, your chosen
interface language and the AI text — never a key, a URL or a model name you
configured.

## Reporting a problem

Please report vulnerabilities privately per
[SECURITY.md](https://github.com/auspect/linexcel/blob/main/SECURITY.md). Do not
attach sensitive workbooks or credentials to public issues.
