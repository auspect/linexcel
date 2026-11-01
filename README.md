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

![Interactive dependency graph](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/app_graph.png)

## Install

```Shell
uv add linexcel               # pip install linexcel
uv add linexcel[ai]           # + AI documentation (optional)
```

> **Note:** `linexcel` depends on [formualizer](https://pypi.org/project/formualizer/), a Rust-based spreadsheet engine. Prebuilt wheels are available for Linux, macOS, and Windows. If no wheel matches your platform, a Rust toolchain is required to build from source.

## Usage

### Interactive web application

Start the application, then upload the main workbook and optional reference
workbooks in the browser:

```Shell
uvx linexcel serve --open --port 8765
# http://127.0.0.1:8765/
# From a development checkout: uv run linexcel serve --open --port 8765
```

`--open` opens the browser; the terminal always prints the local URL. Use
`--port 0` to select a free port when the default is occupied.

This separate path imports a structural graph without evaluating formulas.
It shows source formulas, defined names, saved inputs and Excel caches;
whole-column ranges remain compact. Selecting a cell starts an isolated,
budgeted calculation using its complete supported dependency closure. Missing
caches remain unknown. An unresolved dependency stops calculation rather than
substituting a partial range or another formula's saved value. The graph and
previous results remain available after a task failure.

The backend is Python (WSGI and isolated calculation workers). The browser
interface is HTML, CSS and JavaScript, with Cytoscape.js for the graph; it does
not run Python in the browser. Choose English or French in the interface, or
open `/?lang=en`. This choice is independent of the AI documentation language.

Large workbooks are indexed once on disk. The interactive path streams sparse
OOXML records through native ZIP/XML parsers into SQLite (C), without loading an
openpyxl workbook or a complete Python graph. Formula calculation uses the Rust
engine. Shared strings and copied-formula memberships stay on disk; formatting
only cells do not create nodes. The browser requests pages of at most 200 nodes
and bounded graph neighborhoods. Search covers the disk index, including cells
outside the current page. Counts describe the full index, and omitted graph
neighbors remain explicit.

Uploads use resumable 8 MiB chunks, with a 256 MiB combined workbook/reference
limit by default. No base64 workbook copy is made in the browser. On the local
launcher, **Open a local path** copies a file directly from the server's disk
into an immutable project, bypassing browser transfer. `--no-local-import`
disables this option; it is disabled by default in embedded applications.
For example, allow more disk space for several large indexed workbooks with:

```Shell
uv run linexcel serve --max-upload-mb 256 --max-storage-mb 8192
```

The compressed workbook size does not predict its expanded data or index size.
Time, worker memory and disk quotas still apply. Errors in a worksheet or formula
are reported locally; readable sheets remain browsable, and calculations that
depend on incomplete data are refused. Refreshing an import publishes its new
index atomically. A failed calculation leaves the index and other results usable.
Operation history stores results on disk and loads the selected result only.
Large decompositions have an explicit preview limit and a separate JSON download.
Large ranges stream their literal inputs from SQLite to the calculation engine;
formula dependencies and explicit cell references occupy the Python closure.
Calculation coverage counts
all selected inputs, while dependency cards keep a declared sample of at most 200
nodes, including in the export. A separate 20,000 explicitly traversed-node limit
(excluding literal range members) protects dependency traversal; exceeding it
refuses the calculation explicitly. Very large calculations can require a higher
worker memory budget than the default 512 MiB even though browsing stays small.
Captures are served as individual images with thumbnails, without embedding every
image in the project response. Rendering remains an explicit, budgeted operation.

**Explore the workbook graph** starts a cancellable server task and saves a
hierarchy on disk. It does not start Formualizer. Each view contains at most 32
groups or cells, plus the rest of the workbook. Click a group to expand it, use
the breadcrumb to go back, or search for an exact cell reference to jump to its
group. Positions and dependency counts are prepared by the server; the browser
never downloads or lays out the entire workbook graph. The accessible list
exposes each group's internal relationship count and weighted links between
groups, including links outside the current group. Ranges remain grouped;
unresolved or dynamic references retain the index's limitations. The saved graph
is reused across sessions and invalidated after a successful structural reimport.

Import and graph construction have a separate default time budget of 600 seconds;
targeted recalculation keeps its 120-second default. Both are adjustable under
**Limits per operation**. Exact node references use the index directly; other
searches still match substrings and can take longer on very large workbooks.

Selecting a formula immediately starts a fresh targeted calculation;
simultaneous visits share the active task. **Recalculate** retries or refreshes
that cell and its dependencies. Source inputs and static VBA do not trigger
calculation; navigation never evaluates the whole workbook.

The graph opens **Linexcel calculation** first: compare the value saved in the
file with the targeted result, then inspect intermediate dependency-cell values.
These are engine snapshots, not a trace of executed branches or subexpressions.
The separate **AI documentation** tab shares the graph's wide, adjustable reader.
Long documents scroll inside the panel; on narrow screens it sits below the graph.

![Graph reader with optional AI documentation](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/app_graph_documented.png)

![Selected node with source and calculation evidence](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/app_node_documented.png)

The interface provides sheets, a selected-node graph neighbourhood, dependency
decomposition, cache comparison, task cancellation, deterministic screenshots
and JSON export. Screenshots use the existing LibreOffice/Poppler renderer and
can show LibreOffice-recalculated values; they never change the uploaded source.
The deterministic features require no AI package, provider, Ollama process or
account. Optional AI documentation is requested explicitly for a selected node,
the workbook overview or one selected screenshot. Importing, selecting nodes,
recalculating and rendering screenshots never request AI documentation by
themselves.

Enable those buttons on a local server with the optional AI client and a running
Ollama model:

```powershell
uv run --extra ai --extra screenshots linexcel serve --ai --ai-model qwen3.8
```

The local example uses `qwen3.8`; choose a model actually installed in your
runtime. AI validation is exploratory: model responses vary between runs and
require review against the source formulas, values and images.

The host configures the endpoint and text/vision models; browser requests cannot
override the endpoint or supply API keys. The interface shows the configured
model and keeps generated explanations separate from source formulas, caches
and calculation snapshots. AI output can be wrong, including about image
columns or value provenance; it is not proof that a formula was recalculated.
Use `--ai-base-url`, `--ai-model`, `--ai-vision-model` and `--ai-token-budget`
to configure the service; `LINEXCEL_AI_API_KEY` supplies an optional server-side
key. The default endpoint is local Ollama. AI tasks have a separate duration
and token budget in the interface. Input tokens are estimated before sending,
the output limit is sent to the provider, and reported usage is checked on
return. Image token accounting can remain estimated. Worker memory limits do
not limit the separate model server, and cancelling a task cannot guarantee
that a remote provider stops an already accepted request.
Each generated document shows input, output and total token usage separately;
missing provider counts remain unavailable rather than being reported as zero.
Text documentation makes two provider requests: a draft and a specialized
review against the same bounded source dossier. Both consume the token budget.
The review improves guidance and traceability but cannot certify the final prose.

The **Data** tab offers a static Excel presentation with direct cell styles,
row and column dimensions, merged cells and common number/date formats. Choose
**Datasette filters and exports** for filtering, sorting and CSV/JSON exports.
Both views retain source inputs and saved formula caches; missing caches remain
explicit. Unsupported formats use raw display, and conditional formatting, charts
and print layout are outside the preview. See the [data browser guide](docs/guide/datasette.md)
and [proposed agent workflows](docs/guide/agent-workflows.md).
For `.xlam` and `.xla` add-ins, use the [static VBA reader](docs/guide/macros.md)
to inspect decoded modules and procedures without opening Excel or executing
macros. Add-in reports explicitly have no worksheet calculation or data view.

Drop the main workbook and its optional references into their separate upload
areas using accessible file buttons or drag and drop, inspect or remove selected
files, then follow transfer and import progress. Captures use a thumbnail gallery
and a selected-image view rather than a full-size vertical image list. Operation
details use bounded, highlighted JSON previews with image payloads omitted;
explicit JSON downloads retain the complete result.

![Sheet capture gallery](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/app_captures.png)

The interactive graph shows directed links between sources and their uses, with
pan, zoom, fit-to-view, sheet navigation, search and back navigation. A selected
node preview keeps its formula, value evidence and requested AI documentation
beside the graph. Generate or refresh the same explanation from either the graph
or the detail card. An accessible
list provides the same navigation without requiring canvas interactions. Choosing
a sheet in the graph moves the center onto that sheet. Explicit navigation to
another sheet updates the selector to match the selected cell.
Compact ranges show represented member cells and their full dimensions
separately, without materializing blank positions. Display limits do not reduce
the dependency closure used for calculation.
The detail card includes the original formula, saved and recalculated values,
source provenance, nearby text with cell coordinates, hidden-sheet/row/column
context and dependency formulas. After a targeted calculation, dependency
formula results are read from that same fresh engine operation. Saved caches,
literal inputs and calculated values remain distinct; an unavailable engine
result is never filled from the cache. Reading these results does not launch
separate recalculations of volatile dependencies. Coverage describes the static
formula closure, not which branches executed at runtime. Nearby text is a
bounded contextual snippet, not an inferred business meaning or a change to
calculation inputs.

Copied formula patterns are detected structurally without evaluating the
workbook. Relative and absolute references remain distinct, and each sheet has
its own groups. Group counts and members supplement the individual cells; a
shared formula pattern never implies shared input values or an identical result.
Unsupported reference syntax is kept separate conservatively. Older saved
projects can refresh their structure without uploading their sources again.

The indexed application resolves source-declared tables, scoped static names,
supported 3D aggregates and supplied external formula/range dependencies. It
also handles bounded visibility aggregates, declared array regions, checked
dynamic references and convergent source-enabled iterative cycles. Each feature
has explicit limits in the versioned capability registry and the
[coverage guide](docs/guide/coverage.md#indexed-application-and-batch-analysis).
Missing source metadata, unresolved prerequisites and nonconvergence remain
explicitly unsupported. Decomposition includes dependency formula values
available from the calculation, not evaluated subexpression steps. The native
engine can still differ from Excel;
comparison with a saved value is evidence, not proof of Excel compatibility.

Compatibility corrections run within the calculation graph: `SUMPRODUCT`
ignores nonnumeric entries in separate array arguments, typed errors propagate
through concatenation, and comparisons distinguish numbers, text and Booleans.
An expression inside a `SUMPRODUCT` argument keeps its own arithmetic coercion.
Dependents consume these corrected values; source formulas and saved caches
remain unchanged. Reading a literal input requires no formula engine. Formula
evaluation continues to use formualizer with targeted Python corrections;
Polars is not required. These corrections are covered by independent arithmetic
and type-sensitive regression cases, not just agreement with saved values.
Text ordering involving punctuation or locale-specific collation can still
differ from Excel (for example, `="+">"^"`). These comparisons retain the
engine's approximate ordering; a completed calculation is not a claim of full
Excel compatibility. Saved-value comparisons and source formulas remain visible.
Date calculations preserve numeric serials in both workbook calendars, including
Excel's fictitious leap day; source formatting supplies readable date context.
ISO date text is supported by the corrected date-part functions. Other textual
date formats still use the native parser and can depend on unsupported locale
rules; they are not inferred from the computer's regional settings.

Sources are immutable per imported project. Importing changed workbook or
reference files creates a new content revision and fresh results. Completed
tasks are cached against the sources, engine version and implementation;
the **Recalculate** button forces a fresh calculation, including volatile formulas.
Each displayed result retains its timestamp. Browser reconnection and server
restart recover the task history; work interrupted by a server restart is
explicitly marked failed and can be retried.
Targeted volatile calculations also retain a UTC clock and random seed; the
Python evaluation API can replay that context. This does not establish Excel
random-sequence parity.

### Embed the application

```python
from linexcel.web import create_app

application = create_app(
    "/srv/private/linexcel",
    mount_path="/tools/linexcel",
    resolve_user=lambda environ: environ.get("my_host.authenticated_user_id"),
    max_workers=2,
    total_memory_mb=2048,
    max_upload_mb=256,
    max_storage_mb=2048,
    retention_days=7,
)
```

Mount this WSGI application in the host server. The host must populate the
authenticated identity itself; do not trust a caller-supplied identity header.
A missing identity returns 401. Storage paths and task access are scoped to
that identity. The standalone server binds only to loopback and uses a durable
local profile. Its one-use session link connects a browser to that profile;
`linexcel session-link --data-dir ~/.linexcel` creates another link while the
server is running. Projects survive browser changes and server restarts.
`--mount-path`, `--data-dir`, `--workers`,
`--memory-mb` and `--retention-days` configure the launcher. The default data
directory is `~/.linexcel`; uploaded files and diagnostic formula text are private.

Use **one application coordinator per data directory** (an exclusive lock
enforces this). Its shared queue limits worker concurrency and admission by
the sum of requested memory budgets. Windows Job Objects enforce each worker
tree's memory cap. On POSIX, the current limit is per-process address space;
aggregate renderer memory requires host container/cgroup limits. Upload reads,
graph pages and task-result previews are bounded. LibreOffice temporary disk use
and overall service resources also require host resource limits.
This first implementation is not a distributed worker pool.

Tasks have their own duration and memory limits, indeterminate running progress,
error and cancellation state. Output/storage limits are checked during work
and before publishing, with temporary output removed afterwards; host filesystem
quotas remain necessary for a hard disk ceiling. Standalone projects are retained
until explicit deletion by default. When `--retention-days` is selected, expired
projects are purged at startup based on their import date. Embedded applications
retain their explicit host policy. **Delete** removes a project, sources and
results; running tasks must finish cancellation first.

Download a workspace archive before migrating or deleting a project. Restoration
verifies its inventory and hashes, creates a new project and rebuilds the index;
retained results are historical evidence rather than reusable calculations.
The archive excludes browser credentials. `linexcel doctor` checks local tools,
storage and model discovery without generating AI. Distributed scheduling,
periodic retention enforcement and full subexpression decomposition remain
outside this local workflow.

Selected findings can be shared as a static offline HTML investigation, with
source formulas, saved and calculated values, dated documentation, bounded
relationships and chosen captures. It works with JavaScript disabled. See
[the report limits](docs/guide/html.md#selected-investigation-from-the-local-application).

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

## Focus on what matters in a large workbook

After a confirmed memory failure, isolated analysis can make up to two recovery
attempts within the original time and memory budgets. Recovery reads a bounded
sample of stored cells, divided between worksheets, and reduces that sample
fourfold if it also runs out of memory. It returns formulas and Excel's saved
values only: no recalculation, dependency edges or decomposition. Dates may
remain raw Excel serial numbers. Missing cached values remain unknown.
The result keeps `execution.status = "memory_limit"`, its failure evidence and
attempt history; a recovered inventory is still an incomplete analysis.
Use `--memory-retries 0` (Python: `ExecutionPolicy(memory_retries=0)`) to disable
recovery. Targeted analysis is never replaced with an unrelated cell sample.

The existing `max_cells_per_sheet` limit bounds reading/extraction, not the
native engine's whole-workbook evaluation. Lowering it alone does not prevent
an evaluation memory failure. Recovery counts stored cell entries, so a distant
cell does not require materializing millions of empty cells before it.

Open the HTML report's **Graph** tab and use the sheet selector to focus on
one worksheet at a time. For example, select `Summary` to inspect its
calculations: directly connected nodes from other sheets stay visible and
dimmed, so you can still see where inputs come from and where results go.
Switch back to all sheets to restore the broader view. This filters the
display after analysis; it does not skip loading or analyzing worksheets.

If you only need to explain a particular output, target that cell from the
start instead:

```Shell
uvx linexcel analyze workbook.xlsx --target "Summary!B4" -o summary.html
```

```python
result = analyze("workbook.xlsx", targets=["Summary!B4"])
result.save_html("summary.html")
```

Targeted analysis traces the output's static upstream dependencies, including
those on other sheets, without requesting global recalculation. It is useful
when only a few results matter in a large workbook. It does not guarantee that
only those cells are loaded; dynamic references such as `INDIRECT` and `OFFSET`
can also reach dependencies missing from the graph. Check the report's warnings.

See [filtering and investigation](https://auspect.github.io/linexcel/guide/html/#focus-on-a-worksheet)
and [targeted analysis examples](https://auspect.github.io/linexcel/guide/cli/#trace-only-the-outputs-you-need).

## Features

- **Sheet filtering** — focus the interactive graph on one worksheet while retaining its directly connected cross-sheet neighbors
- **Targeted analysis** — trace selected output cells and their upstream dependencies with `--target` or `targets=`
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
| [AI documentation](https://auspect.github.io/linexcel/guide/ai/)                        | Evidence-based cards, quotation checks, token usage, `token_budget=`                            |
| [Languages](https://auspect.github.io/linexcel/guide/languages/)                        | The nine supported locales                                               |
| [Data handling](https://auspect.github.io/linexcel/guide/data-handling/)                | What leaves the machine, and when                                        |
| [API reference](https://auspect.github.io/linexcel/api/result/)                         | `LineageResult`, `analyzer`, `aidoc`, `powerquery`, `external`, … |

## Sample output

Every image below is captured from a real report by
`scripts/capture_viewer.py`, so they cannot drift from the viewer without the
`readme-shots` commit hook noticing.
The application images above come from `scripts/capture_app.py`, using the
generated B2B commercial review fixture and real sheet captures. The current
images use deterministic analysis without AI documentation.
Its manifest records the fixture, image hashes and selected documentation mode;
`--check` verifies those assets without a browser or model. Both screenshot
checks also run in CI. Workbook sources and full validation reports stay local.

### A node, explained

Source formula, step-by-step evaluation, precedents and dependents. AI cards
can be generated separately from that same deterministic dossier.

![A node selected in the viewer](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/viewer_node_documented.png)

### Workbook summary

![Rendered executive summary](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/viewer_workbook_overview.png)

### Sheet context

Each sheet rendered whole, over a grid of its first cells, alongside its
comments, frozen panes, merged ranges and hidden columns.

![Sheet context tab](https://raw.githubusercontent.com/auspect/linexcel/main/imgs/viewer_sheets_context.png)

## Contributing

Run the [local acceptance workflow](docs/manual_validation.md) before delivering
changes: both fixtures, all report languages and sheet captures, followed
by inspection of dashboard tabs and sampled nodes. AI documentation and image
analysis are optional; use `--no-ai` for deterministic validation. The script preserves each
run and reports incomplete coverage as a failure. See [AGENTS.md](AGENTS.md) for
the repository's validation requirements.

For additional local corpus coverage of the lazy application:

```Shell
uv run python scripts/validate_lazy_corpus.py --corpus /path/to/corpus/files --output validation_screenshots/corpus/my-run
```

The corpus directory contains dataset subdirectories. This read-only probe
selects up to 80 OOXML workbooks per dataset across file sizes, then up to 12
formula targets per workbook across sheets and formula signatures. Imports and
calculations use isolated workers with explicit resource budgets. Each run keeps
its selection, hashes, outcomes and saved-value comparisons in a new ignored
directory. Legacy `.xls` files and archives are inventoried but not converted.
Cache agreement is observational evidence, not an Excel correctness oracle;
this supplementary sample does not replace the full acceptance workflow.

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
