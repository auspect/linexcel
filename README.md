# linexcel

[![PyPI](https://img.shields.io/pypi/v/linexcel)](https://pypi.org/project/linexcel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/auspect/linexcel/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/auspect/linexcel/actions/workflows/ci.yml)
[![Docs](https://img.shields.io/badge/docs-auspect.github.io-blue)](https://auspect.github.io/linexcel/)

Data lineage analysis for Excel workbooks.

Extracts every formula, groups stretched patterns (R1C1 canonicalization), builds a dependency graph (cells, ranges, defined names, VBA), decomposes composite functions with step-by-step evaluation, and optionally documents calculations via the AI provider of your choice.

![Dependency graph of a workbook](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/viewer_graph.png)

## Install

```Shell
uv add linexcel               # pip install linexcel
uv add linexcel[ai]           # + AI documentation (optional)
```

> **Note:** `linexcel` depends on [formualizer](https://pypi.org/project/formualizer/), a Rust-based spreadsheet engine. Prebuilt wheels are available for Linux, macOS, and Windows. If no wheel matches your platform, a Rust toolchain is required to build from source.

## Usage

### Command line

No install needed — `uvx` fetches and runs it in one step:

```Shell
uvx linexcel analyze workbook.xlsx           # -> workbook_lineage.html
uvx linexcel analyze workbook.xlsx --json graph.json --no-html
```

The default is deterministic: lineage only, no network, no key. `--ai-docs`
opts in, and needs the `ai` extra plus an OpenAI-compatible endpoint:

```Shell
uvx --from "linexcel[ai]" linexcel analyze workbook.xlsx --ai-docs \
    --base-url http://localhost:11434/v1 --model laguna-xs-2.1 --language fr
```

`--base-url`, `--model` and `--api-key` also read `LINEXCEL_AI_BASE_URL`,
`LINEXCEL_AI_MODEL` and `LINEXCEL_AI_API_KEY`. Run `linexcel analyze --help`
for the full list, including `--token-budget` to cap what a run may cost.

### Python

```python
from linexcel import analyze

result = analyze("workbook.xlsx")
result                        # interactive graph in marimo / Jupyter
result.save_html("out.html")  # standalone offline HTML viewer
result.stats                  # {totalFormulas, totalNodes, ...}
result.warnings               # list[str]
```

Everything above is local and needs no key. AI documentation is optional, and
you choose the provider — nothing is sent anywhere until you name one:

```python
# A local runtime keeps the workbook on your machine and costs nothing
docs = result.document(base_url="http://localhost:11434/v1", model="laguna-xs-2.1")
overview = result.document_workbook(base_url="http://localhost:11434/v1", model="laguna-xs-2.1")
result.save_html("out.html", docs=docs, workbook_doc=overview, language="en")
```

Any OpenAI-compatible endpoint works the same way — a local Ollama or vLLM
runtime, a gateway such as OpenRouter, a vendor's own API — and `provider=`
takes any callable for anything else. See
[Choosing an AI provider](https://auspect.github.io/linexcel/guide/providers/).

## Features

- **Formula extraction** via [formualizer](https://pypi.org/project/formualizer/) (Rust engine)
- **Stretched pattern grouping** — 1000 identical formulas → 1 node
- **Dependency graph** — cells, ranges, defined names, VBA procedures
- **Step-by-step evaluation** — each operator/function evaluated individually
- **Standalone HTML viewer** — Cytoscape.js embedded, fully offline, keyboard-navigable, light by default with a dark toggle
- **Values you can check** — each figure states whether it was read from the workbook or recalculated by linexcel, and shows both side by side when the file disagrees
- **Workbook context** — sheet previews, comments, merged ranges, frozen panes and hidden columns, plus optional LibreOffice-rendered screenshots
- **AI documentation** — vendor-neutral, grounded in deterministic lineage, with token accounting and a spend ceiling
- **Nine interface languages** — for both the report and the AI prompts

## Documentation

| Guide                                                                                  |                                                                          |
| -------------------------------------------------------------------------------------- | ------------------------------------------------------------------------ |
| [Quick start](https://auspect.github.io/linexcel/guide/quickstart/)                     | Analyse a workbook, explore it, export it                                |
| [Lineage coverage](https://auspect.github.io/linexcel/guide/coverage/)                  | What is in the graph, and what is not                                    |
| [HTML export](https://auspect.github.io/linexcel/guide/html/)                           | The standalone offline report                                            |
| [Workbook context &amp; screenshots](https://auspect.github.io/linexcel/guide/context/) | What a reader sees, not only what the file computes                      |
| [Choosing an AI provider](https://auspect.github.io/linexcel/guide/providers/)          | Ollama, OpenRouter, any OpenAI-compatible endpoint, or your own callable |
| [AI documentation](https://auspect.github.io/linexcel/guide/ai/)                        | Provable cards, token usage, `token_budget=`                            |
| [Languages](https://auspect.github.io/linexcel/guide/languages/)                        | The nine supported locales                                               |
| [Data handling](https://auspect.github.io/linexcel/guide/data-handling/)                | What leaves the machine, and when                                        |
| [API reference](https://auspect.github.io/linexcel/api/result/)                         | `LineageResult`, `analyzer`, `aidoc`, …                           |

## Sample output

Every image below is captured from a real report by
`scripts/capture_viewer.py`, so they cannot drift from the viewer without the
`readme-shots` commit hook noticing.

### A node, documented

Formula, step-by-step evaluation, precedents and dependents, and the AI card
written from that same deterministic dossier.

![A node selected in the viewer](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/viewer_node_documented.png)

### Workbook overview

![AI-written workbook overview](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/viewer_workbook_overview.png)

### Sheet context

Each sheet rendered whole, over a grid of its first cells, alongside its
comments, frozen panes, merged ranges and hidden columns.

![Sheet context tab](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/viewer_sheets_context.png)

## Security

Analysis is entirely local. AI documentation sends dossiers only to the provider
you configure — see
[Data handling](https://auspect.github.io/linexcel/guide/data-handling/).

Please report vulnerabilities privately according to
[SECURITY.md](SECURITY.md). Do not include sensitive workbooks or credentials in
public issues.

## Changelog

See [CHANGELOG.md](CHANGELOG.md).

## License

MIT — see [LICENSE](LICENSE).
