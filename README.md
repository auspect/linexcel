# linexcel

[![PyPI](https://img.shields.io/pypi/v/linexcel)](https://pypi.org/project/linexcel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
![CI](https://github.com/auspect/linexcel/actions/workflows/publish.yml/badge.svg)

Data lineage analysis for Excel workbooks.

Extracts every formula, groups stretched patterns (R1C1 canonicalization), builds a dependency graph (cells, ranges, defined names, VBA), decomposes composite functions with step-by-step evaluation, and optionally documents calculations via AI.

![Global overview](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/overview_example_01.png)

## Install

### uv

```Shell
uv add linexcel
# AI documentation (optional)
uv add linexcel[ai]
```

### pip

```Shell
pip install linexcel
# AI documentation (optional):
pip install "linexcel[ai]"
```

> **Note:** `linexcel` depends on [formualizer](https://pypi.org/project/formualizer/), a Rust-based spreadsheet engine. Prebuilt wheels are available for Linux, macOS, and Windows. If no wheel matches your platform, a Rust toolchain is required to build from source.

## Usage

```python
from linexcel import analyze

result = analyze("workbook.xlsx")
result                    # interactive graph in marimo / Jupyter
result.save_html("out.html")     # standalone offline HTML viewer
result.stats              # {totalFormulas, totalNodes, ...}
result.warnings           # list[str]

# AI documentation (optional, requires google-genai):
# Supports "en" (default) or "fr" language for both documentation and UI
docs = result.document(api_key="...", language="en")
result.save_html("out.html", docs=docs, language="en")

# Workbook-level overview, shown in the separate overview tab:
workbook_doc = result.document_workbook(api_key="...", language="en")
result.save_html("out.html", docs=docs, workbook_doc=workbook_doc, language="en")
```

## Workbook context and screenshots

`result.workbook_context` extracts bounded first rows and columns for every
sheet, without assuming a header row. It also exposes comments, merged cells,
frozen panes, hidden columns, and sheet visibility using `openpyxl`; Excel is
not launched.

These structural details are automatically rendered in a structured summary list
within the **Workbook overview** tab of the HTML report.

You can also generate and embed high-resolution sheet screenshots using LibreOffice Calc:

```python
# 1. Render one PNG per printed workbook page
screenshots = result.save_screenshots("screenshots/")

# 2. Map pages to sheet names to display them inline under each sheet card
sheets_screenshots = {
    "Ventes": screenshots[0:3],
    "Synthese": [screenshots[3]],
    "Params": [screenshots[4]]
}

# 3. Embed them directly inside the offline HTML report
result.save_html("out.html", screenshots=sheets_screenshots)
```

Screenshots require LibreOffice and Poppler's `pdftoppm` installed on the system (e.g. on Debian/Ubuntu: `sudo apt install libreoffice-calc poppler-utils`). Rendering runs via LibreOffice headless, without opening a desktop Excel application.

## AI documentation (optional, multi-provider)

AI documentation is opt-in and supports any LLM provider.

### Google Gemini (default)

```python
docs = result.document(api_key="...", language="en")
```

Requires `google-genai` (`pip install linexcel[ai]`).

### OpenAI-compatible (Ollama, vLLM, LM Studio, OpenAI, …)

```python
# Ollama (local)
docs = result.document(
    base_url="http://localhost:11434/v1",
    model="llama3.1",
    language="en",
)

# Or via env vars
# LINEXCEL_AI_BASE_URL=http://localhost:11434/v1
# LINEXCEL_AI_MODEL=llama3.1
```

Requires `openai` (`pip install linexcel[openai]`).

### Custom provider (any callable)

```python
def my_llm(system_prompt: str, user_prompt: str, *, temperature: float = 0.2) -> str:
    # call your model here
    return response_text

docs = result.document(provider=my_llm)
```

### Workbook-level overview

```python
workbook_doc = result.document_workbook(language="en")
result.save_html("out.html", docs=docs, workbook_doc=workbook_doc, language="en")
```

## AI data handling

AI documentation is opt-in. Calling `result.document()` sends a deterministic
dossier for each requested node, while `result.document_workbook()` sends a
workbook-level dossier, to the configured Gemini model. The dossiers can include
formulas, computed values, precedent/dependent labels, formula decomposition,
sheet structure, defined names, and extracted VBA code. Do not enable this
feature for a workbook whose contents must remain local, unless its data-sharing
requirements permit processing by Google. See the
[Google Generative AI Terms of Service](https://ai.google.dev/terms).

## Features

- **Formula extraction** via [formualizer](https://pypi.org/project/formualizer/) (Rust engine)
- **Stretched pattern grouping** — 1000 identical formulas → 1 node
- **Dependency graph** — cells, ranges, defined names, VBA procedures
- **Step-by-step evaluation** — each operator/function evaluated individually
- **Standalone HTML viewer** — Cytoscape.js embedded, fully offline
- **AI documentation** — Gemini generates provable docs from deterministic lineage

## Sample output

### Global overview

![Global overview](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/overview_example_01.png)

![Global overview (node selected)](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/overview_example_02.png)

## Security

Please report vulnerabilities privately according to
[SECURITY.md](SECURITY.md). Do not include sensitive workbooks or credentials in
public issues.

## Changelog

See [CHANGELOG.md](CHANGELOG.md).

## License

MIT — see [LICENSE](LICENSE).
