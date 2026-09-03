# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- **`analyzer.py` is now a thin orchestrator.** The 1882-line monolith was
  split mechanically into focused modules — `structure`, `engine`, `resolver`,
  `sweep`, `graph` (a `GraphBuilder` that owns all node/edge state) — and
  `analyze_workbook` now sequences the phases and assembles the result.
  `analyzer.py` is down to ~200 lines. Behaviour and output are unchanged:
  the same 586 public tests pass, the public API (`analyze_workbook`,
  `inspect_workbook`, `LineageResult`) is untouched, and the graph snapshot is
  byte-identical before/after. Internals were retargeted in tests and CLI to
  their new modules; no business module imports `analyzer`.
- Internal helper `a1(row, col)` moved into `refs.py` alongside `num_to_col`.

## [1.4.0] — 2026-08-25

### Fixed

- **A comma where the engine writes a dot is no longer a divergence.** A
  workbook saved by a French Excel stores `6,7 €`, linexcel recalculates
  `6.7 €`, and the report used to show a red disagreement on a cell that is
  right — the separator comes from the regional settings of the machine that
  saved the file, and is written nowhere in it. Underneath were two
  contradictory comparisons: the Python one returned "no difference" for
  *every* pair of strings, so a recalculated `oui` over a stored `non` raised
  no warning at all, while the viewer compared rendered text and called any
  difference a disagreement. There is now one rule, with three answers —
  `same`, `format` (the same value written with another locale's separators)
  and `differ` — decided once and carried on the node as `cachedAgreement`.
  German `1.234,5` and Swiss `1'234.5` grouping are read too; grouping is held
  to groups of three digits, so `Salle 1 2` against `Salle 12` stays a
  disagreement. Both readings are still shown side by side; only the verdict
  changes, and the card names the case.
- **A text result that genuinely differs is now reported.** It never was: the
  old comparison ignored every pair of strings. Expect warnings on workbooks
  that were silent before — they were silent wrongly.

### Added

- **A run that will be long says so before it starts.** Past a few seconds,
  the estimate is printed and the two ways out are named:

      16 MB of formulas: this should take about 10 seconds. --dry-run says
      what is in the file without analysing it; -v shows progress.

  It comes from the uncompressed weight of the sheet parts, which the zip
  index already carries, so asking costs about a twentieth of a millisecond
  and every run can afford to. An order of magnitude, not a stopwatch reading:
  it knows nothing about the machine it runs on, and a workbook of mostly
  values finishes faster than it says. Smaller files say nothing. `--dry-run`
  states the estimate too, and `inspect_workbook()` returns it as
  `estimatedSeconds` alongside the new `sheetBytes`.
- `linexcel.analyzer.sheet_bytes()`: the uncompressed size of a workbook's
  sheet parts, read from the zip index without unpacking one.

## [1.3.2] — 2026-08-23

Housekeeping. Nothing an analysis does changes: the reports this version
produces are identical, byte for byte, to the ones 1.3.1 produced, and the
public API is unchanged. Upgrade for the packaging, or skip it.

### Changed

- **The report template and the AI prompts ship as files, not as Python
  strings.** The viewer's 1,700 lines of HTML, CSS and JavaScript now live in
  `assets/viewer.html`, and the 27 system prompts in
  `assets/prompts/<language>/{node,workbook,vision}.md`. Both were unreachable
  by every tool that could check them — no syntax highlighting, no linting, no
  formatter — and a translator had to edit Python to correct a phrasing.
  Adding a language is now three Markdown files and the interface strings; see
  [Languages](https://auspect.github.io/linexcel/guide/languages/).
- **`linexcel.analyzer` is 2,335 lines rather than 2,714**, with two of its
  jobs in modules of their own: `linexcel.loader` reads the values a workbook
  already carries and answers how big it claims to be, and `linexcel.values`
  decides what a value read from anywhere means. Every public name is still
  importable from `linexcel.analyzer`, and is the same object.
- A build installed without its assets now fails at import, naming the cause,
  instead of raising `KeyError` on a language code in the middle of a request.

### For anyone reaching into internals

Private names moved with the code that owns them, and are not re-exported:
`_load_cached_values_calamine`, `_load_cached_values_openpyxl` and
`_detect_epoch_1904` are in `linexcel.loader`; `_jsonable`, `_is_uncomputed`,
`_values_differ`, `_date_text_of` and `_error_kind` are in `linexcel.values`.

## [1.3.1] — 2026-08-22

### Added

- `-v` / `--verbose`: progress while a workbook is analysed, including per-sheet
  progress through the long phases and the same timings even without the
  optional `rich` dependency.
- `--dry-run`: the workbook's declared sheets, sizes, linked workbooks, and the
  ceilings that would apply, then stops before the full analysis.

### Fixed

- A fallback from the fast reader to `openpyxl` no longer happens silently: a
  verbose run now says that it switched reader and names the exception that
  caused it.

## [1.3.0] — 2026-08-18

### Added

- **Power Query is part of the lineage** ([#34](https://github.com/auspect/linexcel/issues/34)).
  A workbook fed by Get & Transform used to show the range the data landed in
  and nothing above it: no query, no M source, no table it read. Each query is
  now a node carrying its M source verbatim, the range it loads into, and the
  sources it names — so `Source!A1:B4` → `BusyProducts` → `Loaded!A1:B3` reads
  end to end. A source that lives in the same workbook (`Excel.CurrentWorkbook`)
  is linked to the table or defined name it reads; a query chaining off another
  is an edge between the two; a source outside the file (`File.Contents`,
  `Web.Contents`, `Sql.Database`, any `*.Database`/`*.Feed` connector) is named
  as a node and said, in the panel and in the warnings, to have been named
  rather than read. Connection-only queries are shown as loading nowhere
  instead of being dropped. Read from the `DataMashup` custom XML part, with
  `xl/connections.xml` and `xl/queryTables/*.xml` for the destination; two new
  stats, `queries` and `queriesLoaded`.
- **An optional multimodal description of each sheet screenshot**
  ([#46](https://github.com/auspect/linexcel/issues/46)).
  `result.describe_screenshots(shots, ...)` — `--screenshots DIR --vision-docs`
  on the command line — sends each rendered sheet to a model that can look at
  it, and the description lands under the image in the Sheets tab, badged *read
  from the screenshot* so it is never confused with a claim about the lineage.
  This is the one thing linexcel documents from a picture rather than from the
  graph, because colour conventions, conditional formatting, charts and the
  shape of a layout never reach a text dossier. It is opt-in, and separate from
  `--ai-docs`, because a picture of a sheet shows every row on it: nothing
  sends an image unless you ask. A text-only endpoint raises rather than having
  the image quietly dropped, and `--vision-model` names the model that looks
  when it differs from the one that writes. Providers may implement
  `generate_with_image` to serve the same call. Note that this is the only card
  in the report with nothing deterministic behind it: a weak vision model
  invents confidently, and the description is placed under its own image so it
  can be read against it.
- **`refs_dir=` / `--refs-dir`: the workbooks this one depends on.** A cell
  reading `'[Budget FY26.xlsx]Annual'!B4` used to be a grey "external
  reference" node with no value, and every formula above it lost its own.
  Three answers are now given, in order: the file **names** the workbook it
  depends on and the path it declares (always, from `xl/externalLinks`); the
  values Excel **cached** across the link are used when it saved them; and,
  given a folder holding the referenced files, the workbook is **read** and the
  reference evaluates to the value it stands for — the reference is stripped of
  its path and resolved against that folder. The node panel lists each linked
  workbook with its path and says which of the three it got, and the same
  folder is searched for `.xlam`/`.xla`/`.xlsm` add-ins so the VBA a workbook
  calls into becomes part of the lineage, each module tagged with its file.

### Changed

- **The value panel always separates what the file carries from what linexcel
  recomputed.** The card used to collapse to a single figure whenever the two
  could not be compared, so "linexcel agrees with Excel" and "Excel never wrote
  a value here" looked identical on screen. Both columns are now always laid
  out, and a column with nothing to show says so — *Not stored*, *Not
  recalculated* — with a sentence naming the reason (a workbook saved without a
  calculation cache carries no value for its formula cells).
- **A stretched formula shows one row per sampled cell** instead of a single
  headline figure followed by a plain list. Each sampled cell keeps its own
  pair of readings, a row whose two readings disagree is marked on the row, and
  the sample is now spread evenly over the group — first cell, last cell and
  three in between, where a pattern that broke actually shows — rather than
  being the first three, near-identical, neighbours.
- **Volatile formulas are reported as not recalculated, on purpose.** `TODAY`,
  `NOW`, `RAND`, `RANDBETWEEN` and `RANDARRAY` answer differently every time
  they are computed, so a recomputed `=TODAY()` could never agree with a file
  saved last week — and linexcel called that a divergence, blaming the workbook
  for the calendar. Such a cell now keeps the value the file stores, says *Not
  recalculated (volatile)* in the recalculation column with a sentence
  explaining why, is not decomposed step by step (every step would carry a
  figure from today's clock), and raises no divergence warning. `OFFSET`,
  `INDIRECT`, `CELL` and `INFO` — volatile to Excel but stable for a given
  workbook — are still computed.
- **Analysis is roughly 2.5× faster on a dense workbook**: 246 s → 94 s on a
  47 MB file of 2.6M formulas, with an identical graph. Two costs went away.
  The step decomposition no longer re-evaluates the *root* step — it is the
  whole formula, so the value the engine already computed for the cell answers
  it; re-running it made a root such as `SUM(Calculs!H2:H200001)` walk 200,000
  formula cells again, 29 s for one node and 114 s over the run. And table
  detection no longer opens the workbook a second time with openpyxl: declared
  tables are read from the `xl/tables/*.xml` parts, static ones from the 30×50
  window the engine already holds, which turns 45 s into nothing measurable.
- **The per-sheet scan ceiling went from 4M cells to 64M**, and the last chunk
  is clipped to what is left rather than dropped whole — a 4,000,000-cell
  budget used to stop at 3,600,000. Sweeping costs about 0.7 µs per cell, so
  the old ceiling bought 1.6 s on the reference workbook and cost that sheet
  its tail; the new one is past any real used range and still bounds the file
  that *declares* more than it holds (one stray cell at XFD1048576 makes the
  used range 17 billion cells). The warning now names the first row left out
  and says its formulas are missing from the lineage. Reads are chunked by
  cells rather than by rows as well: 20,000 rows of a 16,384-column sheet
  meant 327 million strings in one call.

### Fixed

- **A step lost the parentheses of the formula it came from, and reported the
  wrong number.** Each step renders its own expression from the AST and is then
  evaluated by re-parsing that text — but the parser keeps grouping in the
  *shape* of the tree, so `=D2*(1-Rate)` came back as `D2 * 1 - Rate`. The step
  card showed 2470.06 under a cell plainly holding 1976.208. Parentheses are
  now restored from operator precedence, both in what is displayed and in what
  is evaluated.
- **A step could report a value belonging to another cell.** The batch
  evaluation reuses one scratch row across nodes, and `set_formula` does not
  raise on an expression the engine will not take — a structured
  `Table[Column]` it never loaded, a 3D `'A:B'!A1` — it leaves the cell as it
  was. The step then read back whatever the previous node had computed in that
  column. Each scratch cell is now primed with the sentinel first, so a refused
  expression reads as *not evaluated*, which is what it is.
- A root step over an array literal (`=SUM({1,2,3;4,5,6})`) read as `#NAME?`:
  the AST renders an array as the placeholder `{...}`, which evaluates to
  nothing. It now carries the value of its cell.

## [1.2.3] — 2026-08-16

### Added

- Command-line interface: `linexcel analyze workbook.xlsx`, exposed as a
  `[project.scripts]` entry point so `uvx linexcel analyze ...` works without
  installing anything. Deterministic by default; `--ai-docs` opts into AI
  documentation and needs the `ai` extra
  (`uvx --from "linexcel[ai]" linexcel ...`). Also runnable as
  `python -m linexcel`.

### Changed

- The git tag is now the only source of the version, and the `VERSION` file is
  gone. It held `0.0.0` in the repository and was overwritten from the tag
  during a release, so it never showed anything true: every development build,
  and `linexcel --version` with it, reported `0.0.0` while PyPI was at 1.2.2.
  A clean `v*` tag still builds exactly that number; anything else builds
  `1.2.2+dev.8.g7f5caf5`, with a `.dirty` suffix for uncommitted edits. PyPI
  rejects local versions, so a build that is not exactly a tag cannot be
  published by accident.

### Fixed

- `publish` replayed via `workflow_dispatch` from a branch wrote the branch
  name as the version and failed later inside `uv build`. It now stops
  immediately, naming the ref and what to do instead.

## [1.2.2] — 2026-08-13

### Added

- `verbose=` / `--verbose`: a timing breakdown of each phase of the analysis,
  for finding where a large workbook spends its minutes.

## [1.2.1] — 2026-08-13

### Fixed

- A static table whose header cell held a formula, or nothing at all, took that
  cell's rendering as the column name. It falls back to the column letter.

## [1.2.0] — 2026-08-12

### Added

- Cached values are read with `python-calamine`
  ([#41](https://github.com/auspect/linexcel/issues/41)), so the value stored in
  the file is available for every cell without a second openpyxl pass.

## [1.1.1] — 2026-08-12

### Fixed

- Step evaluation called the engine once per cell. The calls are batched, which
  is roughly 1200× faster on a large workbook.

## [1.1.0] — 2026-08-10

### Added

- Excel tables are detected, and the lineage carries the header and the index of
  a cell that sits in one.

### Fixed

- `formualizer>=0.8.0` is now required: earlier versions expanded shared formulas
  incorrectly.

## [1.0.0] — 2026-08-08

First stable release. The AI surface changed in two ways that a caller can
notice, hence the major version: no vendor is built in any more, and the
workbook overview now sees the workbook, not only its formula graph.

### Changed (breaking)

- **AI documentation is vendor-neutral.** No provider is named in the code and
  none is chosen for you. `document()` and `document_workbook()` used to fall
  back to Google Gemini when nothing was passed; they now raise `AiDocError`
  listing the options. Two ways in remain, and they are the whole surface:

  | | |
  | --- | --- |
  | `base_url=` + `model=` | any OpenAI-compatible endpoint — a local Ollama, vLLM or LM Studio runtime, a gateway such as OpenRouter, a vendor's own API |
  | `provider=` | your own callable or `LLMProvider` object |

  Removed with it: the `_GeminiProvider` client, the `google-genai` dependency,
  the `DEFAULT_MODEL` constant, and the `GEMINI_MODEL` / `GOOGLE_API_KEY` /
  `GEMINI_API_KEY` environment variables. A model that a vendor happens to
  privilege is a decision for the caller, not a constant in a lineage library —
  and the previous "opt-in" still left one vendor reachable by a shorter path
  than the others.

  The `linexcel[ai]` extra now installs `openai`, which is the client for
  every OpenAI-compatible endpoint rather than a choice of vendor; `[openai]`
  stays as an alias. Migration: replace `model="gemini-..."` with the
  `base_url=` of an endpoint plus its model id — Gemini's own
  OpenAI-compatible URL works — or wrap the native SDK in a `provider=`
  callable.

- **A model must always be named.** `base_url=` without `model=` used to fall
  back to `gpt-4o-mini`, a name that means nothing to a local runtime. It now
  raises, naming the endpoint it could not pick a model for. `LINEXCEL_AI_MODEL`
  (or `OPENAI_MODEL`) is the environment equivalent, and `LINEXCEL_AI_API_KEY`
  joins `OPENAI_API_KEY` for credentials.

- `document_workbook()` now sends the workbook context — the first rows and
  columns of every sheet, cell comments and their authors, merged ranges, frozen
  panes, hidden columns — alongside the lineage dossier. Cell *contents* therefore
  leave the machine, where previously only formulas and structure did. This is
  the point of the change: an overview written from the graph alone describes a
  dependency structure nobody recognises, while the titles, labels and comments a
  reader actually sees are what make it a description of the file. Every fact is
  still read deterministically from the workbook by `openpyxl`, so the "cite only
  the dossier" guarantee is unchanged, and no image is ever uploaded. Migration:
  pass `include_context=False` to restore the previous payload.

- **`save_screenshots()` returns a mapping keyed by sheet name**, not a flat
  `list[Path]` of print pages. The flat list could not be shown under the sheet
  it belonged to, which is the whole reason to render a workbook at all; the
  report was reduced to a separate tab of unlabelled pages. Migration: pass
  `per_sheet=False` for the previous return value and the previous tab. Note
  that a mapping is still downgraded to that flat list when the renderer does
  not produce exactly one page per sheet — an older LibreOffice ignores the
  single-page option — so a caller that does not pass `per_sheet=False` should
  handle both shapes. `to_html()` and `save_html()` already accept either.

### Added

- `token_budget=` on `document()` and `document_workbook()`: a ceiling on the
  **total** tokens a run may spend, input and output across every request.
  `token_usage` already reported what a run cost afterwards; this decides it
  beforehand. It is deliberately global rather than per node — one card's cost
  is never the question, and a workbook with 900 formula patterns is 900
  requests whose sum is the only figure that reaches an invoice.

  Enforced between requests, the only point at which a cost is known: nodes
  still queued when the tally reaches the ceiling are never sent, the cards
  already written are returned, and a `UserWarning` names how many were skipped.
  Requests already in flight are allowed to finish, so the final tally can
  exceed the ceiling by up to `max_workers` responses: a budget is an order of
  magnitude, not a to-the-token limit. Counted against
  `result.token_usage`, which spans the result's lifetime, so one budget covers
  the overview and the node cards together. A budget already spent raises
  `AiDocError` rather than sending a request that would breach it

- `include_context=` on `document_workbook()` (default `True`), and a
  keyword-only `context=` on `aidoc.build_workbook_dossier()`. Sheets that hold
  no formula — a data-only tab contributes no lineage node — now appear in the
  dossier through the context rather than being invisible to the overview

- `per_sheet=` on `save_screenshots()` (default `True`). LibreOffice is asked to
  put each sheet on a single page, which is what makes a page map onto a sheet
  at all: under the workbook's own print layout a long sheet spans several pages
  and short ones share one, so no page number maps back and the images could
  only ever be shown as an unlabelled flat list. Rendered files are named after
  the sheet they show (`demo-Synthese.png`). Pass `per_sheet=False` for the
  print pages instead. A sheet with nothing on it renders one pixel tall and is
  left out rather than shown as a broken image

- `tests/fixtures/`, holding the two workbooks `openpyxl` cannot author:
  `macros.xlsm` (a real `vbaProject.bin`) and `power_query.xlsx` (a real M
  mashup). The VBA extraction code had been tested only against a stubbed
  `olevba`, which is to say the boundary with it had never run at all; the
  fixture put `vba.py` at 99% coverage and immediately turned up the sheet
  class modules above. `.gitignore` exempts that one directory from the blanket
  Excel-file ignore, so a workbook of your own dropped in the tree stays ignored

- `--check-max-tokens` in `validate_manual.py`, which documents one node under
  several `max_tokens` ceilings and reports which of them actually cut the
  response. An OpenAI-compatible endpoint is free to ignore the parameter, and
  no test suite can answer that for the endpoint you happen to run; the check
  reports rather than asserts, and says so when a ceiling was simply never
  reached — the same prompt sampled twice comes back at very different lengths,
  so a shorter answer under a ceiling is no evidence the ceiling caused it.
  `--max-tokens` also now applies to a whole validation run

- `scripts/capture_viewer.py`, which captures the README images from a real
  report by driving the interface (clicking tabs, searching for a node) instead
  of photographing the landing state repeatedly. It records the hashes of
  `viewer.py` and `i18n.py` in `imgs/manifest.json`, and the new `readme-shots`
  prek hook fails a commit that changes either without recapturing — the README
  can no longer advertise an interface the package does not have. The hook only
  *checks*: capturing needs Playwright (`linexcel[screenshots]`), an ordinary
  commit does not

### Changed — HTML report

- **The Sheets tab shows the sheets.** It listed badges and comments and left
  the rest of the pane blank: the first cells of every sheet were already
  embedded in the document and never drawn, and the rendered image only appeared
  when the caller happened to hand in a per-sheet mapping, which nothing
  produced. Each sheet now shows its own rendered image — in a scrollable frame,
  since a sheet renders onto one page and a long one is a very tall picture —
  above a grid of its first cells, with column letters, row numbers and the
  hidden columns marked. The grid comes from `openpyxl` alone, so the tab has
  content on a machine with no renderer installed at all

- The **value card now states what it is comparing**. Reading values back is a
  headline feature and the report undersold it: one figure and a small
  provenance pill, with the workbook's own stored value mentioned only when it
  happened to differ — so in the ordinary case a reader could not tell whether
  they were looking at Excel's number or linexcel's. When both exist the card
  shows both, labelled `Excel file` and `linexcel recalc`, each with its
  provenance. Agreement is stated rather than implied by silence; a
  disagreement gets an icon, a sentence and a border, never colour alone

- Light is the default theme and dark is opt-in, via a toggle in the top bar.
  Spreadsheet work is done in light, and a report that opened dark because the
  reader's OS was dark was the wrong default for this audience. The toggle sits
  outside the graph-only tool group so it stays reachable on every tab, and
  `localStorage` is wrapped in try/catch — it *throws* inside the sandboxed
  `data:` iframe used for notebooks, and a thrown exception must never break the
  report. Cytoscape paints to a canvas and cannot resolve CSS variables, so the
  graph re-reads its colours and restyles on each flip

- The search box is a proper control: an inline SVG magnifier, a clear button,
  Escape to reset the graph, and the match count on screen in an `aria-live`
  region instead of hidden in a `title` attribute, with a distinct "no matches"
  state

- The detail panel is 440px rather than 340px, and below 560px the two value
  readings stack so the comparison survives the narrow overlay

- The top bar is three groups — identity, tabs, graph tools — instead of one
  undifferentiated row. The graph-only controls (search, sheet filter, layout,
  zoom, fit) now hide on the Overview, Sheets and Visual preview tabs, where
  they did nothing

- Accessibility: real tab semantics (`role="tablist"`, `aria-selected`,
  `aria-controls`, roving tabindex, Arrow/Home/End), a visible focus ring,
  `aria-pressed` on the layout and page buttons, and accessible names on the
  icon-only controls. Clicking a precedent or dependent now moves focus into
  the detail panel — the button that was clicked is destroyed by the re-render,
  so focus used to fall back to `<body>` and the keyboard path ended there

- Contrast fixes, with the CVD-safe palette unchanged: badge ink is computed
  from its fill (white on the amber `#eda100` was about 2:1), links use a token
  of their own, and the muted grey went from 3.3:1 to 4.7:1. The four sheet
  badges dropped their hard-coded hex pairs — one of which was a pink invisible
  in dark mode and a CVD risk against the amber — for the shared palette

- The detail panel overlays the graph below 900px instead of squeezing it, and
  the empty-state panel stands down entirely rather than covering the graph

- Ten interface strings that were hard-coded English are now translated in all
  nine languages: the layout buttons, and the whole sheet detail pane (`rows ×
  cols`, freeze, hidden columns, merged ranges, comments). A French report was
  previously part English

- Report layout moved out of per-element JavaScript `style` assignments into the
  stylesheet; only genuinely data-driven values (a palette colour, the step
  indent) are still set inline

### Fixed

- **Excel error values are shown as Excel spells them.** The engine reports an
  error as a `{"type": "Error", "kind": ...}` dictionary, and nothing turned it
  into text, so the panel showed `{'type': 'Error', 'kind': 'Na'}` — Python
  source, where the reader was looking for `#N/A`. Worse, that string could
  never equal the `#N/A` stored in the file, so the value card announced *every*
  error cell as "the recalculated value differs from the file" while both
  readings said the same thing. `#DIV/0!`, `#N/A`, `#NAME?`, `#NUM!`, `#NULL!`,
  `#REF!`, `#VALUE!`, `#SPILL!` and `#CALC!` now read as themselves, compare
  against the stored value as themselves, and a divergence warning names them in
  the same terms

- **A limit of the engine is no longer reported as a value of the cell.** A
  formula the engine does not implement — the range intersection in
  `=SUM(D2:D10 D5:D20)` — came back as an internal `NImpl` error and was
  presented as what linexcel recalculated, then set against the file's own
  figure as if the workbook were wrong. Nothing was recalculated. Such a cell
  now keeps the value stored in the file and says so, its step decomposition
  reads *not evaluated*, and a warning names the cells it happened to. Reference
  cycles (`Circ`), a cancelled evaluation, and any error kind whose spreadsheet
  meaning we cannot vouch for are treated the same way

- **A VBA project that cannot be read no longer fails silently.** Every failure
  in `extract_vba_modules()` — an unreadable project, a module stream olevba
  chokes on, oletools missing — returned the same empty mapping as a workbook
  holding no macro at all, so a `.xlsm` full of code was reported as having
  none and nothing said otherwise. The reason is now a workbook warning, as is
  the case where the file declares macros but no module could be read. A failure
  part-way through keeps the modules already read rather than discarding them:
  their procedures and call graph are real, and the warning says the rest is
  missing

- **A workbook with one macro module no longer reports five.** Excel writes a
  class module per worksheet and one for `ThisWorkbook` whether or not anybody
  puts code in them, and they are not empty — they hold `Attribute VB_*`
  declarations, so they counted. `vbaModules` now counts modules somebody wrote.
  A sheet module that does hold code — a `Worksheet_Change` handler — has lines
  beyond the attributes and is kept. Found by reading a real `.xlsm` rather than
  a stub

- A date cell previewed as `2026-01-03T00:00:00`, a timestamp nobody typed:
  `openpyxl` reads a date back as a midnight datetime and the whole ISO form was
  kept. It now reads as the day it holds, in the Sheets tab and in the dossier
  the AI overview is written from

- One unresolvable reference no longer costs the whole workbook its computed
  values. `evaluate_all` is all-or-nothing and gives up on the *first* reference
  it cannot resolve, so a single formula pointing at a closed workbook dropped
  every cell in the file back to the slow per-cell recovery. Those few cells are
  now set aside before the pass and the global evaluation completes; only the
  isolated cells lose a recomputed value. On the `stress` fixture that is 6
  cells set aside instead of 561 cells recovered one at a time.

  Guarded formulas are never isolated. `IFERROR(NOSHEET!A1, 456)` has a correct
  value despite an unresolvable reference, and blanking it would not merely cost
  that cell — every range spanning it would quietly return a smaller number.
  A workbook whose only blockers are guarded keeps the previous behaviour

- Defined names declared on a single worksheet are now collected. Only
  workbook-scoped names were read, so a formula using a per-sheet `Total` or
  `Limit` — ordinary in real files — resolved to nothing and showed up as an
  external reference. Workbook scope still wins over sheet scope for the same
  name, so a shadowed name cannot displace the one most formulas mean

- `LET` bindings and `LAMBDA` parameters are no longer graph nodes. The parser
  reports them as references because that is what they look like, so every
  intermediate a modeller named produced an "external reference" node. The
  binding is local to its formula and points at no cell; it stays visible in
  that cell's step decomposition, which is where it belongs

- Value samples never showed their provenance. The report read `sample.source`
  while the graph emits `sample.valueSource`, so the label saying a sampled cell
  came from the file rather than from a recalculation silently never rendered

- A defined name whose target sits on a sheet with an apostrophe never resolved.
  `openpyxl` strips the surrounding quotes from `DefinedName.destinations` but
  leaves the doubled apostrophes of the escaped form, so a sheet called
  `O'Brien` arrived as `O''Brien`; storing that verbatim made the label quote it
  a second time (`'O''''Brien'`) and the name pointed at a sheet nobody has. It
  became an unresolved external reference instead of an edge to the real cell.
  Found by the new `stress` validation workbook

### Changed

- The workbook system prompt now states that titles, labels and comments quoted
  in a sheet preview are citable evidence, while a sheet name alone still is not.
  Without this the model kept answering "not determined by lineage" for a
  workbook whose purpose was written in the cell above its table

- `aidoc.MAX_WORKBOOK_DOSSIER_CHARS` raised from 12,000 to 16,000 to fit the
  presentation context. An oversized workbook now sheds detail in order — long
  previews shrink, then the tail of the pattern and VBA lists, and only as a last
  resort are previews and comments dropped — so a workbook that fitted before
  loses nothing

- `document_nodes()` submits work to its thread pool a few nodes at a time
  rather than all at once, which is what lets a budget stop a run. Concurrency,
  results and partial-failure behaviour are unchanged

- `validate_manual.py` rewritten around a **local** model: it defaults to an
  OpenAI-compatible endpoint at `http://localhost:11434/v1`, checks the model is
  actually served before starting (and lists what is, if not), takes
  `--model` / `--base-url` / `--token-budget` / `--max-nodes` / `--no-ai`, and
  prints the token tally at the end. A full validation run now costs nothing and
  sends nothing off the machine

- A second validation fixture, `--workbook stress`: an English workbook built to
  break the analyser rather than to look plausible. Every Excel error value and
  its `IFERROR` guard, external and 3-D and structured references, `INDIRECT`
  and `OFFSET`, a circular pair, LET bindings, a formula past the dossier
  truncation limit, 30-deep nested `IF`, a sheet named `O'Brien's Café`, hidden
  and very-hidden and empty sheets, numbers stored as text, the 1900 leap-year
  window, and cell text aimed squarely at the report's HTML escaping. Both
  fixtures moved to `validation_workbooks.py`; `--workbook both` runs the pair,
  and `--max-nodes` keeps a local model's turnaround reasonable over the ~60
  calculation nodes the stress workbook produces

- Generated fixtures are round-tripped through LibreOffice so they carry the
  results a real Excel file stores. `openpyxl` writes formulas but never their
  values, so until now no fixture had a stored value at all and the report's
  file-versus-recalculated comparison had nothing to compare. The round-trip
  also yields honest disagreements — LibreOffice does not implement every
  function formualizer does, so the file stores `#NAME?` for `XLOOKUP` where
  linexcel computes the number — which is precisely what that comparison exists
  to surface. `--no-recalc` skips it

- `validate_manual.py --file PATH` runs the whole pipeline against a workbook of
  your own. This is the only way to exercise VBA end to end: `openpyxl` can
  preserve a `vbaProject.bin` but cannot author one, so no generated fixture can
  contain macros. The workbook is read, never rewritten. `linexcel.vba`'s
  parsing and graph integration stay covered by the unit suite, which injects
  modules directly

### Documentation

- A [Lineage coverage](https://auspect.github.io/linexcel/guide/coverage/) page:
  what is in the graph, what is represented but not resolved, and what is not
  there at all. A lineage tool is only useful if you know where it stops, and
  the honest entry is **Power Query**: a workbook whose data arrives through a
  query shows the range it loaded into and nothing about where that data came
  from — not the query, not its M source, not the table it reads. Everything
  needed is in the file (`customXml/item1.xml` carries the M source of every
  query, `xl/connections.xml` names the range each lands in), so this is a gap
  to close rather than a limit of the format; tracked in
  [#34](https://github.com/auspect/linexcel/issues/34) for a release after 1.0.
  `tests/fixtures/power_query.xlsx` pins what the graph produces today, so the
  gap is visible in the suite and not only in a document

- The README is a landing page again — install, usage, features, screenshots,
  and a table of links. It had grown to carry the AI provider matrix, the
  screenshot install commands, the language table and the data-handling rules in
  full, all of which also existed in the guide; two copies of a provider list is
  one copy too many the first time they disagree

- Three guide pages added, two of them extracted from the README so the detail
  lives in exactly one place:
  [Choosing an AI provider](https://auspect.github.io/linexcel/guide/providers/)
  (worked examples for a local runtime, a gateway and a custom callable, with
  the environment-variable table), plus
  [Languages](https://auspect.github.io/linexcel/guide/languages/) and
  [Data handling](https://auspect.github.io/linexcel/guide/data-handling/).
  The AI guide keeps what is specific to the feature — grounding, the workbook
  overview, concurrency, token usage and budgets

## [0.7.0] — 2026-08-02

### Added
- Token accounting: `result.token_usage` tallies every `document()` and
  `document_workbook()` call, exposing `input_tokens`, `output_tokens`,
  `total`, `requests` and a readable `str()`. Recovered from the abandoned
  `feat/agnostic-ai` branch, whose provider abstraction was superseded
- Counts are read from the provider when it reports them — Gemini's
  `usage_metadata` and the OpenAI-compatible `usage` block — so the figure
  matches what is billed. They are approximated only as a fallback, and
  `TokenUsage.estimated` then flags the tally
- `aidoc.estimate_tokens()`, the fallback estimator, counts CJK characters
  individually instead of letting `\w+` swallow a spaceless Japanese or Chinese
  sentence as a single word
- Optional `usage=` accumulator on `aidoc.document_nodes()` and
  `aidoc.document_workbook()`, and a `UsageReportingProvider` protocol that
  custom providers may implement to report real counts. Return types are
  unchanged, so existing callers are unaffected

- Seven more languages for `language=`, covering both the AI prompts and the
  viewer interface: `es`, `de`, `it`, `pt`, `nl`, `ja`, `zh` (with `en` and
  `fr`, nine in total). The set stays a closed allowlist — `language` selects a
  stored system prompt and reaches the generated JavaScript, so free-form input
  would be a prompt-injection and interpolation vector
- `linexcel.i18n`, single source of truth for the language list and the viewer
  interface strings, which used to be hand-maintained inside the JS template.
  The suite asserts `i18n.UI_STRINGS` and the two `aidoc` prompt registries
  cover the same languages and the same keys
- Translation provenance is stated in the README, the AI guide and the
  `linexcel.i18n` docstring: `en` and `fr` are hand-written, the other seven
  were produced with AI assistance and are not natively reviewed

- Multi-provider AI support: Google Gemini (default), OpenAI-compatible endpoints (Ollama, vLLM, LM Studio), and custom callables
- `base_url` and `provider` parameters on `document()` and `document_workbook()`
- Environment variables: `LINEXCEL_AI_BASE_URL`, `LINEXCEL_AI_MODEL`
- Optional dependency: `pip install linexcel[openai]`
- README: multi-provider documentation section
- `linexcel.__version__`, read from the installed distribution metadata
- `max_workers` on `document()` to tune AI request concurrency

### Fixed
- Workbook screenshots now work on Windows and macOS, not only Linux.
  `save_screenshots()` looked for LibreOffice and `pdftoppm` on `PATH` alone,
  which neither the Windows nor the macOS installer extends, so the renderer
  reported itself as missing on machines where it was installed. Both binaries
  are now also looked up in their standard install directory, including
  winget's package store for Poppler
- On Windows the renderer invokes `soffice.com` rather than `soffice.exe`:
  `soffice.exe` detaches and returns immediately, so the conversion was reported
  as finished before the PDF existed and the run failed with "LibreOffice did
  not produce a PDF"
- Screenshot rendering uses a throwaway LibreOffice user profile
  (`-env:UserInstallation`). A LibreOffice already open on the desktop owns the
  default profile, and the headless process would exit successfully having
  converted nothing
- `save_screenshots()` reports which of the two binaries is missing and gives
  the install command for the running platform, instead of naming both and
  assuming Debian
- The screenshot pane hard-coded its heading and its `Page N` buttons in French,
  so every report in the eight other languages was partly French. Both now go
  through `linexcel.i18n`, which gains a `page` string
- The workbook-context example in the documentation iterated
  `result.workbook_context` as if it were keyed by sheet name; per-sheet context
  lives under its `sheets` key
- `linexcel.PackageNotFoundError` was never meant to be public; the
  `importlib.metadata` import is now private
- `provider=` now accepts a plain callable, as the documentation always claimed;
  previously only objects exposing a `generate` method worked
- VBA call edges were never emitted: the procedure lookup was keyed on
  `Module.Name` while the call scanner reports unqualified names. Calls now
  resolve in the calling module first, then across modules, and an ambiguous
  name stays unresolved instead of pointing at an arbitrary module
- VBA call detection is case-insensitive throughout, matching the language:
  procedure lookups are keyed on the lowercased name, so two modules declaring
  `Taux` and `taux` no longer capture each other's calls nor defeat the
  ambiguity guard, and a function's own return assignment (`Taux = 1`) is not
  mistaken for a call
- A dotted name is a VBA call only when the qualifier is a module
  (`Module1.Taux`, now resolved exactly); `.Value` or `.Count` is member access
  and no longer produces an edge to a similarly named procedure
- `document()` no longer discards every successful card when a single node
  fails; failures are reported through a `UserWarning` and `AiDocError` is
  raised only when all nodes fail
- `to_html()` no longer fails on a `LineageResult` built without the workbook
  bytes; the sheet-preview tab is simply omitted
- Viewer template placeholders are substituted in a single pass. Chained
  replacements let `__TITLE__` and `__LANG__` be rewritten *inside* the
  already-injected graph, so a cell, sheet name, or comment containing either
  literal was silently replaced by the title — and a title ending in a
  backslash truncated the embedded JSON, blanking the viewer
- The `cells` interface string was used twice by the viewer but defined in no
  locale, so reports fell through to the raw key ("1 000 cells" in French).
  The unused `sheets_summary_title` key was dropped

### Changed
- **Breaking:** `render_html()`, `to_html()` and `save_html()` reject a
  `language` outside `("en", "fr")` with `ValueError`. It was previously
  interpolated raw into the generated JavaScript, where it could terminate the
  string literal; unknown values used to fall back to English at runtime
- `pytest` now collects the `src/` doctests, which `testpaths` used to exclude
- Generated validation artifacts are no longer tracked in the repository
- AI data-handling notes describe every provider, not just Gemini
- `save_screenshots(timeout=)` defaults to 180 seconds instead of 60. A first
  headless run builds the LibreOffice profile from scratch, which alone can
  exceed the old budget on Windows
- `linexcel.insights` exposes `find_libreoffice()` and `find_pdftoppm()`, so a
  caller can check the renderer is available before analyzing a workbook
- `linexcel.i18n` gains a `page` string, used by the screenshot page switcher

## [0.3.0] - 2026-07-14

### Added
- Localization: `en` (default) and `fr` language support for AI documentation and UI
- Sidebar worksheet overview in HTML report
- Sheet-specific screenshot embedding in HTML report
- Workbook-level AI documentation (`document_workbook()`)
- `SECURITY.md`, `THIRD_PARTY_NOTICES.md`

### Changed
- Refined AI data handling documentation and security notes

## [0.2.2] - 2026-07-10

### Changed
- Improved formula decomposition and step evaluation
- Refined dependency graph edge resolution for named ranges

## [0.1.0] - 2026-07-10

### Added
- Initial release
- Formula extraction via formualizer (Rust engine)
- Stretched pattern grouping via R1C1 canonicalization
- Dependency graph: cells, ranges, defined names, VBA procedures
- Step-by-step composite function evaluation
- Standalone HTML viewer (Cytoscape.js, fully offline)
- Optional AI documentation via Google Gemini
- LibreOffice screenshot rendering
- PyPI trusted publishing
