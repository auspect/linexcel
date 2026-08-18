# linexcel

[![PyPI](https://img.shields.io/pypi/v/linexcel)](https://pypi.org/project/linexcel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/auspect/linexcel/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/auspect/linexcel/actions/workflows/ci.yml)
[![Docs](https://img.shields.io/badge/docs-auspect.github.io-blue)](https://auspect.github.io/linexcel/)
[![uv](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json)](https://github.com/astral-sh/uv)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
[![ty](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ty/main/assets/badge/v0.json)](https://github.com/astral-sh/ty)

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
uvx linexcel analyze workbook.xlsx --refs-dir ./linked   # workbooks it reads
```

A workbook that reads `'[Budget FY26.xlsx]Annual'!B4` depends on a file
linexcel does not have. It always names that file and the path the workbook
declares; `--refs-dir` points at a folder holding them, and the reference is
then read for real — the same folder is searched for the `.xlam`/`.xla`
add-ins whose VBA the workbook calls.

The default is deterministic: lineage only, no network, no key. `--ai-docs`
opts in, and needs the `ai` extra plus an OpenAI-compatible endpoint:

```Shell
uvx --from "linexcel[ai]" linexcel analyze workbook.xlsx --ai-docs \
    --base-url http://localhost:11434/v1 --model qwen3.8 --language fr
```

`--base-url`, `--model` and `--api-key` also read `LINEXCEL_AI_BASE_URL`,
`LINEXCEL_AI_MODEL` and `LINEXCEL_AI_API_KEY`. Run `linexcel analyze --help`
for the full list, including `--token-budget` to cap what a run may cost.

Sheets can also be rendered and, separately, read by a multimodal model:

```Shell
uvx linexcel analyze workbook.xlsx --screenshots shots/    # LibreOffice, local
uvx --from "linexcel[ai]" linexcel analyze workbook.xlsx \
    --screenshots shots/ --vision-docs --base-url ... --vision-model ...
```

`--vision-docs` is the only option that puts a picture of a sheet in a request,
so it is opt-in and independent of `--ai-docs`.

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
docs = result.document(base_url="http://localhost:11434/v1", model="qwen3.8")
overview = result.document_workbook(base_url="http://localhost:11434/v1", model="qwen3.8")
result.save_html("out.html", docs=docs, workbook_doc=overview, language="en")
```

Any OpenAI-compatible endpoint works the same way — a local Ollama or vLLM
runtime, a gateway such as OpenRouter, a vendor's own API — and `provider=`
takes any callable for anything else. See
[Choosing an AI provider](https://auspect.github.io/linexcel/guide/providers/).

## Features

- **Formula extraction** via [formualizer](https://pypi.org/project/formualizer/) (Rust engine)
- **Stretched pattern grouping** — 1000 identical formulas → 1 node
- **Dependency graph** — cells, ranges, defined names, VBA procedures, Power Query queries
- **Power Query lineage** — each query with its M source, what it reads and the range it fills, so data from Get &amp; Transform is not a dead end
- **Step-by-step evaluation** — each operator/function evaluated individually
- **Standalone HTML viewer** — Cytoscape.js embedded, fully offline, keyboard-navigable, light by default with a dark toggle
- **Values you can check** — what the file stores and what linexcel recomputed, always side by side, each named and each stated when it is missing; a stretched formula is compared cell by cell over a sample spanning the whole group
- **Honest about what it cannot compute** — volatile formulas (`TODAY`, `NOW`, `RAND`) are shown as *not recalculated* rather than compared against the clock, and a cell reading another workbook names that file, its path, and whether it was read
- **Dependencies you can supply** — `--refs-dir` resolves linked workbooks and reads the VBA of the add-ins a file calls into
- **Workbook context** — sheet previews, comments, merged ranges, frozen panes and hidden columns, plus optional LibreOffice-rendered screenshots
- **AI documentation** — vendor-neutral, grounded in deterministic lineage, with token accounting and a spend ceiling
- **Screenshots a model can read** — optional, opt-in: a multimodal model describes each rendered sheet, for the colour conventions and layout no extraction reaches
- **Nine interface languages** — for both the report and the AI prompts
- **Command line** — `uvx linexcel analyze workbook.xlsx`, no install required

## Roadmap

Shipped:

- [x] Deterministic lineage — formula extraction, stretched-pattern grouping, dependency graph, VBA
- [x] Step-by-step evaluation, with every value checked against the one stored in the file
- [x] Standalone offline HTML viewer, in nine languages
- [x] Workbook context and LibreOffice-rendered sheet screenshots
- [x] AI documentation — any OpenAI-compatible endpoint, token accounting, spend ceiling
- [x] Command-line interface, installable-free through `uvx`
- [x] Power Query lineage ([#34](https://github.com/auspect/linexcel/issues/34)) — queries as nodes, with their M source, their sources and the range they fill
- [x] Vision ([#46](https://github.com/auspect/linexcel/issues/46)) — an optional multimodal description of each sheet screenshot, for what a text dossier cannot carry

Planned:

- [ ] **`formulas` as a fallback** ([#37](https://github.com/auspect/linexcel/issues/37)) — a second parser for the workbooks formualizer cannot read, so an unsupported construct degrades the graph instead of failing the analysis.

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
| [API reference](https://auspect.github.io/linexcel/api/result/)                         | `LineageResult`, `analyzer`, `aidoc`, `powerquery`, `external`, … |

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
