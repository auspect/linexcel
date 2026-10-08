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
| `describe_screenshots()` | The sheet screenshots themselves — one PNG per call, inline in the request |

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
    base_url="http://localhost:11434/v1", model="qwen3.8"
)
```

Screenshots are the widest thing linexcel can send, so nothing sends them
unless you ask. `document()` and `document_workbook()` never do: they are given
facts in text, read from the file by `openpyxl`. Only
[`describe_screenshots()`](ai.md#describing-the-screenshots) — and its
`--vision-docs` flag, which additionally requires `--screenshots` — puts a
picture of a sheet in a request, and a picture shows everything on that sheet,
including rows no dossier would have quoted. The image travels inline in the
request body, so nothing is uploaded or left behind under a file id.

## Where it goes

| Configuration | Destination |
| --- | --- |
| Nothing configured | Nowhere — `AiDocError` is raised and the message lists the options |
| `base_url=` on a local runtime (Ollama, vLLM, LM Studio) | Your own machine |
| `base_url=` on a hosted endpoint or gateway | That operator, under their terms — read them |
| `provider=` | Wherever your own callable sends it |

An explicit proxy or an HTTPX environment proxy may carry these requests too.
Configure `http_client_kwargs={"trust_env": False}` on `OpenAICompatProvider`
to ignore environment proxy and certificate settings, or set `trust_env=False`
on an injected HTTPX client. Custom transports, hooks and authentication
handlers can also inspect the payload; their behavior belongs to your code.

linexcel takes no position on which is acceptable, because only you know what is
in the workbook. Do not enable AI documentation for a file whose contents must
stay local unless the provider you configured satisfies that requirement.

## Credentials

API keys are read from the argument you pass or from the environment
(`LINEXCEL_AI_API_KEY`, `OPENAI_API_KEY`). linexcel does not serialize provider
configuration into the report or JSON graph export. The generated HTML contains
the graph, your chosen interface language and the AI text. Source workbook
content or generated prose can still contain sensitive values, including
credentials embedded in M source. Token accounting exposed through `result.token_usage`
includes the model ID. Provider API errors report the exception type without
quoting the server payload; the chained SDK exception and SDK debug logging can
still contain request details. Caller-supplied hooks and providers control their
own logging.

## Reporting a problem

Please report vulnerabilities privately per
[SECURITY.md](https://github.com/auspect/linexcel/blob/main/SECURITY.md). Do not
attach sensitive workbooks or credentials to public issues.
