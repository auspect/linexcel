# Reading, resilience and the data node (batch 2)

Date: 2026-11-14. Starting point: `e011ef3` (head of draft PR #149, `v1.19.2`
plus the batch 1 and 1 bis commits). No product cap, behaviour or dependency
was changed in this batch. Deliverables: this note, the measurement scripts and
an isolated prototype in `prototypes/data_node/` (see its `README.md`).
The findings register lives in [`critical_review.md`](critical_review.md); new
findings from this batch are added there as F17–F26.

## 1. Environment and method

- Windows 11, AMD Ryzen 9 5900HS (8 cores), 31 GB RAM, NVMe disk.
  Python 3.14.4, `uv` 0.13.5. formualizer 0.14.0 (installed = locked = latest
  on PyPI), python-calamine 0.9.1, openpyxl 3.1.5.
- **Corpus** (`linexcel-lab/workbooks`, outside the repository, public
  licences): `dgt-elections-2017-2020-tour1.xlsx` (103 MB, one sheet
  `A1:BF346691`, 796 MB of worksheet XML, no formula),
  `ons-population-2011-2025.xlsx` (50 MB, `MYEB2` `A1:FN57878` in an Excel
  table, 336 MB of XML, no formula), `nrel-atb-2024-v4.xlsx` (5 MB, 39 sheets,
  233,410 formulas), `dft-tag-data-book-v2-03.xlsm` (6 MB, 122 sheets,
  79,529 formulas), and the generated `fpna-20mb-seed42.xlsx` (20 MB, a
  249,675-row ledger table read by 12,483 `SUMIFS`). Damaged copies of
  `ons-mye23tablesuk.xlsx`, `nrel-atb-2024-v4.xlsx` and `qra-template.xlsm`
  were made by `make_damaged.py`; the originals were not modified.
- Every measurement runs in a fresh process. Peak memory is the Windows peak
  working set of that process. Import, calculation and the Data tab are
  measured through the functions `linexcel serve` calls (`build_index`,
  `evaluate_node`, `build_source_database`, `sheet_presentation`,
  `build_graph`), and end to end through the HTTP coordinator with
  `serve_path.py` (local import as the UI does it, UI default budgets 600 s and
  1024 MiB, server defaults 256 MiB upload and 2048 MiB storage).
- **Contention.** The first import series (ATB, TAG, fpna, ONS, DGT) ran while
  the test suite or other probes were running on the same machine. The figures
  marked *quiet* were taken afterwards with nothing else running. Where both
  exist, the quiet figure is the reference.
- Outputs, indexes and logs are in `validation_screenshots/lot2/` and
  `linexcel-lab/runs/lot2-*` (ignored, never committed).

Baseline test suite on this machine (`uv run pytest`, run while the first
imports were running): **5385 passed, 0 failed**, 62 skipped, 537 deselected,
4 xfailed, 1023 s. Skips: Playwright not importable in a plain sync (39),
`openai` (4), `opentelemetry` (1), Darwin/POSIX-only process tests (12),
real `libproc` (2), Windows symlink privilege (1), doctest `+SKIP` (3).

## 2. Inventory of reading and analysis caps

Family **H** = protection against a hostile or broken file, or against
resource exhaustion of the machine: keep. Family **C** = convenience of the
current implementation: candidate for replacement. Response bounds (what is
*shown*) are kept: the plan wants caps on what is displayed, not on what is
read.

| Cap | Value | Where | Protects | What the user sees when hit | Family |
| --- | --- | --- | --- | --- | --- |
| Upload size | 256 MiB (`--max-upload-mb`) | `WebConfig`, `UploadStore.begin` | disk, memory | 413 "Import trop volumineux" before transfer | H |
| Storage quota | 2048 MiB (`--max-storage-mb`), includes indexes and the Data tab copy | `TaskStore`, worker supervision | disk | at upload: 413 "Stockage plein"; during a task: task failed, kind `storage_limit`, generic advice "Consultez la formule et ses sources…" (F22) | H, but it bites honest workbooks because of index amplification (F06) |
| Task time budget | UI default 600 s, 1–3600 s | `TaskStore.budget`, supervision loop | CPU | task failed `timeout`, advice "Ajustez les limites de l'opération" | H (the budget); its *effect* is global (F21) |
| Task memory budget | UI default 1024 MiB (API 512), ≤ `--memory-mb` 2048 total | Windows Job / `RLIMIT_AS` / macOS RSS sampling | memory | "Budget mémoire de N Mio insuffisant…" | H |
| JSON body, chunks, references, pending tasks | 8 MiB `/api/projects`, 64 KiB otherwise; 8 MiB chunk; 32 references; 32 pending tasks; uploads expire after 24 h | `web.py`, `web_upload.py` | request abuse | 413/429 with a message | H |
| XML entity expansion | bundled expat limits; `DOCTYPE`/`ENTITY` rejected in presentation, tables, documentation evidence | `lazy_store`, `excel_presentation`, `structured_refs` | billion laughs | sheet `partial` (import), 400 (presentation) | H |
| In-memory XML parts | 4 MiB metadata (presentation), 4 MiB table part, 32 MiB DataMashup, 6000 VBA characters per procedure shown | several | memory on hostile parts | part diagnostic or 413 | H |
| Formula depth | `MAX_AST_DEPTH` 900 | `engine.is_too_deep` | native stack | calculation refused, "safe complexity limit" | H |
| Excel coordinates | 1 048 576 × 16 384 | readers | invalid input | cell diagnostic | H |
| Presentation window | 100 rows × 40 columns, 2 MiB response, 8192 styles, 2000 merges per window, 1024-character format | `excel_presentation` | response size | 413 with a French message | H (response bounds) |
| **Presentation worksheet size** | **64 MiB of sheet XML** | `excel_presentation.MAX_WORKSHEET_BYTES` | time: every window re-parses the whole sheet | 413 "Feuille trop volumineuse pour cet aperçu statique" | **C** (F08) |
| **Presentation element count** | **1,000,000 XML elements** | `excel_presentation.MAX_ELEMENTS` | same | 413 "Feuille trop complexe pour cet aperçu statique" | **C** (F08) |
| Datasette | 50 rows per page, 200 rows max, 1 s SQL, no custom SQL, 8 cached instances, grid tables up to 512 columns | `web_data` | response size, safety | Datasette's own messages | H |
| **Data tab snapshot** | full copy of every cell, built on first open, counted in the storage quota | `web_data.build_source_database` | (none: it is a derived copy) | 413 "Stockage insuffisant pour la vue des données"; meanwhile the whole server waits (F18) | **C** |
| Graph pages | 200 nodes per page, depth ≤ 20, hierarchy fan-out 32, 16 MiB task result, 200 steps and 200 range members shown | `lazy_store`, `graph_store`, `web` | response size | `truncated`, "omitted" counts | H (response bounds) |
| **Calculation closure** | 20,000 formula/structural nodes (literal range members excluded) | `lazy.evaluate_node` | time and memory, already bounded by the worker budget | refusal "Exact dependency closure exceeds the 20,000-node calculation budget" | **C**, not measured in this batch |
| **Engine input loading** | one `set_value` per input cell | `lazy._evaluate_index` | (none: implementation choice) | slow automatic calculation on selection (F07) | **C** |
| **Index publication** | all-or-nothing at the end of `build_index` | `lazy_store.build_index` | atomicity | a timeout or quota stop leaves no index at all (F21) | **C** (the atomicity is kept, its granularity is the convenience) |
| Library search | 100,000 entries and 100,000 numeric cells per workbook, 50,000 header-scan cells, 256 MiB catalog, 120 s build | `CatalogLimits` | catalog size | coverage `capsHit` in the library | C, not measured in this batch |
| `analyze` loader | 64 M cells per sheet sweep, 20 M dense cells | `loader.py` | memory (calamine dense range) | `analyze` only; not on the serve path | H for `analyze` |

The serve import itself (`_read_source`) has **no cell or size cap**: it is a
streaming reader with bounded memory (63 MB peak on the 20-million-cell DGT
file). The caps a user of `serve` actually meets on large workbooks are the
time budget and the storage quota, because of what is *stored* (F06), and the
presentation caps, because of what is *re-read* for every window (F08).

## 3. F06–F08 on the real corpus

### F06 — import cost

| Workbook | Stored cells | Formulas | Import | Index | Bytes per cell | Peak memory |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| DGT (103 MB) | 18,131,244 | 0 | 1310 s | 10,969 MB (107×) | 605 | 63 MB |
| ONS (50 MB) | 11,235,995 | 0 | 733 s | 6,717 MB (134×) | 598 | not recorded (probe bug, fixed afterwards) |
| fpna-20mb (20 MB) | 2,709,887 | 174,770 | 293 s; quiet 265 s | 2,022 MB (100×) | 746 | 64 MB |
| ATB (5 MB) | 353,515 | 233,410 | 144 s; quiet 134 s | 454 MB (90×) | 1,284 | 71 MB |
| TAG (6 MB) | 124,106 | 79,529 | 80 s; quiet 72 s | 169 MB (28×) | 1,358 | 99 MB |

The generated stand-in of batch 1 (100,000 × 10, 80 s, 522 MB) understated
the problem: real data tables cost 65–72 µs and about 600 bytes per stored
cell, and formula models 380–650 µs per cell (F23). The `nodes` table holds
85–90 % of every index (one zlib-compressed JSON record of ≈440 bytes per literal).

With the defaults of `linexcel serve` (2 GiB storage quota, 600 s, 1024 MiB),
DGT and ONS cannot be opened at all: the import stops on the storage quota
after 219 s and 230 s and publishes nothing (section 4, F21). fpna-20mb
imports, but its default Graph view and its Data tab then stop on the same
quota.

### F07 — targeted calculation over a real table

Selecting a cell in `serve` calculates it (section 8), so this is the cost of a
click. `fpna-20mb`, quiet, `bench.py evaluate` (serve's `_evaluate_index`
both times; the second run only replaces how region inputs reach the engine):

| Cell | Inputs | Today | Region sub-ranges, chunked `set_values_batch` | of which loading | Value, comparison |
| --- | ---: | ---: | ---: | ---: | --- |
| `Forecast!D2` = `SUMIFS` over 4 ledger columns | 998,703 | **156.9 s** | **23.5 s** | 5.9 s | 563,775.07 both, `equal` to the cache |
| `Summary!B2` = `SUM(tblLedger_01[Amount])` | 249,675 | 33.1 s | 7.2 s | 1.06 s | 19,951,757,538.32 both, `equal` |

The remaining 17 s for `D2` are spent outside value loading (closure walk,
step presentation, index reads; see F24 for the Windows part). Even a small
calculation pays a fixed cost on a large index: `Forecast!E2` (56 nodes)
takes 3.5 s, almost all of it in two calls of `selected_inputs`, whose join
scans the whole position index (F26; 0.0 s when the selection drives the
join).

### F08 — Data tab

Excel presentation window (row 1, 60 × 20), `sheet_presentation` in a fresh
process, quiet:

| Sheet | Worksheet XML | Result | Time |
| --- | ---: | --- | ---: |
| DGT `Feuil1` | 796 MB | refused: "Feuille trop volumineuse pour cet aperçu statique" | 0.8 s |
| ONS `MYEB2` | 336 MB | refused, same message | 0.8 s |
| fpna `Ledger_01` | 94 MB | refused, same message | 0.8 s |
| ONS `MYEB1` (row 1 and row 60,000) | 42.5 MB | refused: "Feuille trop complexe pour cet aperçu statique" (> 1,000,000 elements) | 2.7 s |
| **TAG, every sheet** (`Cover`, `A1.3.1`, `MECs`) | 0.1–15 MB | refused: "Métadonnées de présentation trop volumineuses" — `styles.xml` is 7.0 MB, over the 4 MiB metadata cap (F25) | 0.8 s |
| fpna `Forecast` | 13.4 MB | 1,020 cells | 1.9 s |
| ATB `Summary` | 1.9 MB | 1,200 cells | 1.05 s |

The same windows from the prototype store take 1.5–2.0 ms at the top, the
middle and the bottom of DGT, `MYEB2` and the ledger (section 7).

Datasette view (`build_source_database`, built on the first opening of the
tab, inside the HTTP request and under the coordinator lock):

| Workbook | Build | Copy on disk | Effect |
| --- | ---: | ---: | --- |
| ATB (5 MB) | 50–64 s | ≈470 MB (project storage 925 MB minus the 454 MB index) | an unrelated `GET /api/projects` waited 64 s (F18) |
| fpna-20mb, no quota | 161 s | 1,931 MB | — |
| fpna-20mb, serve defaults | 100.8 s, then 413 "Stockage insuffisant pour la vue des données" | — | the server is blocked for those 100 s, and the tab stays unavailable |
| ONS, no quota | 745 s | 6,320 MB | in serve, 12 minutes during which every request waits |

The default **Graph view** (`build_graph`, hierarchy of every node) takes
74.7 s and 215 MB on fpna; with serve defaults it stops on the storage quota
after 14 s and the view answers "Construisez le graphe du classeur". On ONS,
a workbook without a single formula (one edge, from its one defined name), it
takes 274 s and 783 MB to arrange 11.2 million input cells into a hierarchy.

## 4. Failure modes

Through the HTTP coordinator, UI defaults (`serve_path.py`, results in
`validation_screenshots/lot2/serve-runs/`).

| Case (real workbook) | Import | Extent of the failure | What the application shows |
| --- | --- | --- | --- |
| Worksheet XML truncated (ONS `MYE2 - Persons`, ATB `Land-Based Wind`) | succeeded, `partial` | that sheet | sheet `partial` with "Worksheet import incomplete: ParseError…"; other sheets, graph, Data tab and calculation work; presentation of that sheet 400 "Présentation de la feuille indisponible" |
| Worksheet bytes corrupted inside the ZIP (same sheets) | succeeded, `partial` | that sheet | same as above (expat stops before the CRC check) |
| Worksheet part missing, relationship kept | succeeded, `partial` | that sheet | sheet `partial` ("There is no item named…"); presentation 404 "Feuille introuvable", although the sheet is listed (F22) |
| `xl/_rels/workbook.xml.rels` missing | succeeded, `partial` | every sheet | all 10 sheets `partial` ("Worksheet relationship is missing"); an empty but openable workbook |
| `xl/workbook.xml` truncated | **failed** | whole workbook | task failed `ParseError` |
| ZIP truncated (no central directory) | **failed** | whole workbook | task failed `BadZipFile: File is not a zip file` (inherent) |
| `styles.xml` truncated (fixed in 1.19.2+F09) | succeeded, `partial` | styled cells | workbook diagnostic on the part; values of styled cells unknown; presentation 400 for every sheet (known, left as is) |
| `sharedStrings.xml` truncated | succeeded, `partial` | strings after the cut | workbook diagnostic; those cells unreadable |
| **`sharedStrings.xml` missing, relationship kept** | succeeded, **`complete`** | every text cell | **no workbook diagnostic**: 560 of 2,733 inputs unreadable, visible only cell by cell (F17) |
| Table part truncated | succeeded, `partial` | that table | diagnostic on the part; the sheet stays readable |
| External link to an absent `https` workbook (QRA) | succeeded, `complete` | none at import | a formula reading it is refused on calculation |
| External workbook missing, references not supplied (supply chain) | succeeded | the formula | calculation `unsupported` (refused); with the two references supplied: computed |
| ONS (50 MB) with serve defaults | **failed** after 230 s, `storage_limit` | **whole workbook**, including its 4 small cover and notes sheets | task error `storage_limit` with the generic advice "Consultez la formule et ses sources…"; nodes list 409 "Actualisez la structure pour activer l'index sur disque" (F21, F22) |
| DGT (103 MB) with serve defaults | **failed** after 219 s, `storage_limit` | whole workbook | same |
| fpna-20mb with serve defaults | succeeded (277 s) | the default Graph view and the Data tab | `build_graph` stops on the quota after 14 s ("Construisez le graphe du classeur"); the Data tab blocks for 100.8 s then answers 413; the ledger has no presentation; clicking `Forecast!D2` calculates for 172 s; project storage 2.06 GB |
| First Data tab open on ATB (5 MB) | — | **the whole server** | 50–64 s synchronous build; an unrelated `GET /api/projects` waited 64 s (F18) |

Target from the plan: a failure per sheet or per part, never global. Reached
for damaged parts except `workbook.xml` and the ZIP directory. **Not reached
for resource limits** (F21) and for the Data tab (F18).

## 5. formualizer beliefs checked

| Belief | Source | Evidence | Verdict |
| --- | --- | --- | --- |
| 0.14.0 has no values-only region or lazy loading API | batch 1, plan | stubs: no region/provider/columnar name; `load_workbook(strategy=…)` "accepted for backward compatibility; loading is eager"; `dir(fz)` has no such name | confirmed (0.14.0) |
| `set_values_batch` is much faster than `set_value` | batch 1 (100,000 cells: 1.5 s vs 5.6 s) | real ledger columns, 2 columns: 3.4 µs/cell at 10,000 rows, 7.6 at 50,000, 15 at 100,000, **53 µs/cell at 250,000 rows** in one call; `set_value` 22–108 µs/cell. 500,000 cells: one call 24.7 s, chunks of 10,000 rows 1.9 s, of 1,000 rows 1.5 s, per cell 41.6 s. Same `SUMIFS` result every time | **changed**: true only per chunk; one call is superlinear (F19) |
| `add_table` works over batch-loaded values | plan | `SUMIFS(tblLedger_01[Amount],tblLedger_01[Entity],…)` = 990,826,382.78, equal to the cell-range form | confirmed |
| `range_page` pages a range without evaluating | plan | `Ledger_01!A1:J249676`, `total` 2,496,760, `next_offset` 3 | confirmed |
| Batch values keep their types | plan | `ISTEXT`, `ISLOGICAL`, `ISBLANK`, `ISERROR` all TRUE; `ERROR.TYPE` of a `#DIV/0!` input 2 | confirmed |
| formualizer's own reader could serve as a fast reader | review hypothesis | `load_workbook(fpna-20mb)` → `OSError: Unsupported feature: ZIP64 member in xl/worksheets/sheet3.xml`; on a workbook with one damaged sheet, or damaged shared strings or styles, the whole load fails (eager); DGT and ONS → `Unsupported feature: XML part byte limit` with default load limits | refuted for this use (F20) |

## 6. Readers compared

Quiet, one fresh process per measurement (`reader_bench.py`). "linexcel
stream" is serve's reading primitive (ZIP + ElementTree `iterparse`,
detaching rows) with no decoding and no storage: the floor of today's import.
fastexcel 0.21.0 and Polars 2.1.0 ran in a throw-away environment outside the
project.

| Reader | DGT `Feuil1` (20.1 M cells) | ONS `MYEB2` (9.8 M) | fpna `Ledger_01` (2.5 M) |
| --- | --- | --- | --- |
| linexcel stream, parse only | 96 s, 41 MB | 47 s, 41 MB | 13 s, 41 MB |
| python-calamine 0.9.1 (`to_python`) | 15.3 s, **1,559 MB** (2,059 MB committed) | 5.6 s, 746 MB | 2.0 s, 355 MB |
| fastexcel 0.21.0 (`load_sheet` → Polars) | 9.9 s, **1,446 MB** (2,088 MB committed) | 3.9 s, 757 MB | 1.9 s, 280 MB |
| Polars `read_excel(engine="calamine")` | 10.0 s, 1,446 MB | — | 2.0 s, 295 MB; then filter + group-by + sort over the whole table **0.04 s** |
| formualizer 0.14.0 `load_workbook` | refused: "XML part byte limit" | refused: same | refused: "ZIP64 member" |

- **Row limit.** `nrows=1000` (calamine) and `n_rows=1000` (fastexcel) do not
  reduce time or memory: DGT 13.6 s / 1,417 MB and 8.0 s / 1,447 MB. Both
  decode the whole sheet into a dense range first, then cut.
- **Column selection.** fastexcel `use_columns` with 2 of 58 columns on DGT:
  7.5 s, 1,447 MB. No memory gain, for the same reason.
- **Excel tables.** fastexcel `load_table` reads `tblLedger_01` in 2.0 s and
  ONS `Table2` in 3.8 s; calamine needs `load_tables=True`.
- **Schema.** fastexcel guesses each column from 1,000 rows. On DGT, ONS and
  the ledger the guesses match the per-column types the prototype counts over
  every row (for example DGT `effectif`: 272,157 numbers and 74,534 genuinely
  empty cells; fastexcel's 74,534 nulls are those blanks). On the small ONS
  tables with a two-row header (names, then codes) every numeric column is
  guessed `string`, and table column names keep their ST_Xstring escapes
  (`England _x000a_and Wales`). The default `dtype_coercion="coerce"` turns a
  value that does not fit the guessed type into null without error;
  `to_arrow_with_errors` reports them. No reader gives number formats, styles
  or the stored cell type (`s`, `str`, `inlineStr`), which provenance needs.
- **Damaged files** (`readers-damaged/summary.txt`): with a truncated or
  missing shared-string part, or truncated styles, calamine, fastexcel and
  formualizer refuse the **whole workbook**; linexcel's reader continues and
  marks what it cannot read. With one damaged sheet, calamine and fastexcel
  still read the healthy sheets (they load per sheet); formualizer refuses
  the whole workbook (eager load).
- **Memory against serve's budgets.** A task gets 1,024 MiB by default
  (Windows Job commit limit) out of 2,048 MiB for the server. A dense reader
  on DGT needs about 2 GB committed: it would be stopped with the default
  budget. The streaming reader stays near 41–64 MB whatever the size.

Dependencies, if Philippe wanted them: **fastexcel** 0.21.0, MIT, 11.8 MB
installed, no required dependency, but it only hands data over through the
Arrow PyCapsule interface, so in practice it needs Polars (or pyarrow) to be
useful. **Polars** 2.1.0, MIT, 4.8 MB plus `polars-runtime-32` 164 MB.
Neither opens the network. Neither is needed by the proposal below.

## 7. Data node prototype

Code: `prototypes/data_node/` (see its `README.md`). It is not
imported by the product and not packaged.

### What it does

- **Reading is unchanged.** `build_store` calls serve's own
  `lazy_store._read_source` and only replaces the storage it writes to. Every
  decoding rule (ST_Xstring, shared strings, shared-formula translation, date
  systems, unreadable values, per-sheet failure locality) is therefore the
  product's.
- **Storage.** Literal values go to one SQLite grid per sheet: a `row` key
  and one column per Excel column, holding the saved value as an SQLite
  integer, real or text (a date keeps its serial). Booleans, errors, ISO
  dates, JSON values and unreadable cells are exceptions in a side table; an
  unreadable cell is never stored as a blank. Number formats are kept per
  column as runs of rows (`[first, last, format]`), exact and small for real
  tables. Formula cells, names and dependency nodes stay per-cell nodes, as
  today. Provenance becomes per sheet (workbook SHA-256, part, status) plus
  the value: the raw XML spelling of a number is not kept.
- **Region detection.** (1) An Excel table: bounds and column names from the
  table part, header and totals rows excluded, columns holding any formula
  split off; each remaining run of adjacent columns of at least 10,000 cells
  becomes a region. (2) Otherwise, runs of adjacent columns that hold literals
  and no formula; a header row is the first row within 20 whose non-empty
  cells are all text and cover at least half the columns, followed by a row
  holding a non-text value; rows above it (titles, notes) stay per cell.
  (3) A column mixing literals and formulas, and a block under 10,000 cells,
  stay per cell. A region never contains a formula, so the fail-closed
  calculation rules keep applying to every formula.
- **The node** carries its bounds, header row and names, a schema (one type
  per column: `number`, `text`, `date_serial`, `mixed`, `empty`, with counts
  of what was observed), the number formats, the row count, the exception
  counts, a SHA-256 fingerprint of its values and a bounded preview
  (10 rows × 12 columns).
- **Lineage.** `link` maps each range node of the existing index to the
  region it intersects, with the sub-range read.
- **Calculation.** `bench.py evaluate` runs serve's `_evaluate_index`
  unchanged, with an index whose `select_range_inputs` keeps region cells out
  of the per-cell selection and loads the region sub-ranges into formualizer
  in chunks of 10,000 rows with `set_values_batch`. Closure, refusals, cache
  comparison and coverage counts are serve's. An unreadable cell in a loaded
  sub-range raises instead of loading a blank.

### Results

| Workbook | Today: import, index | Prototype: store, size | Regions | Verified against serve's index |
| --- | --- | --- | --- | --- |
| DGT | 1310 s, 10,969 MB (605 B/cell) | **510 s, 79 MB** (4.4 B/cell), 53 MB peak | 1: `Feuil1` 346,690 × 58, header row 1 detected, 100 % of literals | 3,000 random cells, 0 mismatch |
| ONS | 733 s, 6,717 MB | **279 s, 26 MB** (2.3 B/cell), 55 MB peak | 4 Excel tables (`MYEB1`–`MYEB4`), header row 2 under a title row, 99.98 % of literals | 3,000 cells, 0 mismatch |
| fpna-20mb | 265 s (quiet), 2,022 MB | **95 s, 33 MB** (13 B/cell), 50 MB peak | `tblLedger_01` 249,675 × 10; `tblForecast` key columns (3 of 17; the 14 formula columns stay per cell) | 3,000 cells, 0 mismatch |
| TAG | 72 s (quiet), 169 MB | 9 s, 1.7 MB (literals only) | none: parameter blocks under 10,000 cells, interleaved with formulas | — |
| ATB | 134 s (quiet), 454 MB | 17 s, 2.5 MB (literals only) | none, same reason | — |

A first fpna run took 658 s: the prototype's private scratch database answered
each shared-string lookup outside a transaction, which costs about 270 µs on
Windows (F24). Holding its lock brought the same build to 95 s with identical
region fingerprints. The prototype times include serve's reader (the parse
floor alone is 96 s on DGT and 47 s on ONS, section 6) and the description
of every region
(fingerprint and schema over every value: 9.2 s on DGT). For TAG and ATB the
prototype does not extract dependencies, so its time is not comparable with
the import: it only shows that their literals are small.

- **Windows.** A 60 × 20 presentation window from the store: 1.5–2.0 ms at the
  first, middle and last rows of DGT, `MYEB2` and the ledger, three sheets the
  application refuses to present today.
- **Lineage.** On fpna, the four `Ledger_01` ranges read by every `Forecast`
  `SUMIFS` map to the ledger node with their exact sub-ranges (columns A, B, C
  and J, rows 2–249,676): 4 edges instead of membership of 998,700 cells.
  The `tblForecast` key-column region is read cell by cell: 49,932 single-cell
  references from formulas on the same rows. A data node must therefore
  accept single-cell sub-ranges, and lineage to such a cell must stay as
  precise as today.
- **Calculation**: section 3, F07 — 156.9 s to 23.5 s and 33.1 s to 7.2 s,
  same values, same comparisons, same input counts.

### Mixed cases and what they mean

- *Calculated column in a data table* (`tblForecast`): the formula columns
  stay per-cell nodes grouped by the existing copied-formula patterns; only
  the literal columns form a region. A column with one formula anywhere in
  the sheet is excluded entirely — conservative, and simple to explain.
- *Title and notes above a table* (ONS): handled by the table part, or by the
  header detection; the preamble stays per cell.
- *Two-row headers* (small ONS tables): not recognised as such; the second
  row would be data. Under 10,000 cells they stay per cell anyway.
- *Formula models* (TAG, ATB): nothing to gain from a data node; their cost
  is elsewhere (F23).

## 8. Other leads

**Streaming each sheet with a budget per part, then windows served by
request.** The reader already streams one sheet at a time with bounded memory
(63 MB on DGT) and keeps a broken sheet local. What is not per part is
*publication*: `build_index` writes into a temporary file and publishes it
with one `os.replace` at the end. A time or storage stop therefore discards
every sheet already read (F21). Committing and publishing per sheet (or per
part), with a sheet that exceeds its share marked `not_imported` with the
reason, keeps the atomicity of each published state and turns a global
failure into a local one. Windows served by request are what the data node
store provides (section 7): the cap then applies to what is displayed (100 ×
40 cells, 2 MiB), not to what was read.

**Decomposition computed on demand for the opened cell, never upfront.** This
is already how `serve` works: import evaluates nothing, and selecting a cell
schedules a targeted calculation of that cell only (`scheduleAutoCalculation`
in `app.html`, 180 ms after selection; a previous result for the same
revision is reused). Nothing to change in the model. What makes it costly is
the calculation itself on large inputs (F07) and, on Windows, the per-lookup
locking cost of the index (F24).

**Windows: one read transaction per index connection.** On this machine an
SQLite `SELECT` outside a transaction costs 260–300 µs, because each one takes
and releases the file lock; inside one read transaction, with
`locking_mode=EXCLUSIVE` or on an `immutable=1` URI it costs 8–9 µs
(100,000 primary-key lookups on a 500,000-row table). `WorkbookIndex` opens
read connections in `mode=ro` autocommit, so every `nodes[id]` of the serve
read paths pays the lock. Measured on serve's own functions
(`probe_win_lock.py`, same values in both modes): a depth-2 neighbourhood
(200 nodes) takes 0.57–0.98 s as shipped and 0.10–0.16 s with one read
transaction (ATB, TAG, fpna: 5–9×); a targeted calculation with an
11,532-node closure 2.89 → 2.02 s, small closures unchanged (the first
cell of each run also carries module warm-up and is not counted).

## 9. Proposals

### Convenience caps and mechanisms

| Cap or mechanism | Proposal | Measured gain | Risk | What stays bounded |
| --- | --- | --- | --- | --- |
| One zlib/JSON record per literal (F06), the reason the quota and the time budget stop honest tables | Grid store and data nodes (section 7) | DGT 10,969 → 79 MB and 1310 → 510 s; ONS 6,717 → 26 MB and 733 → 279 s; fpna 2,022 → 33 MB and 265 → 95 s. With serve defaults both large tables would fit the 2 GiB quota; DGT stays close to the 600 s budget (reader-bound, see F23 and the 96 s parse floor) | schema change and re-import; per-cell provenance becomes per sheet plus the value (the raw XML spelling of a number is not kept); input nodes synthesised on read; region detection is a heuristic | upload size, storage quota, time and memory budgets, unchanged |
| All-or-nothing publication of the index (F21) | Commit and publish per sheet within the budget; a sheet not reached is `not_imported` with its reason; a formula reading it is refused | not prototyped: on ONS, the cover, notes and contents sheets and every sheet finished before the stop would stay available instead of nothing | a partially published index must say so everywhere; a closure must refuse, not skip, a missing sheet | the same budgets |
| Presentation re-reads the whole sheet: 64 MiB and 1,000,000-element caps (F08) | Serve windows from the store; index row heights, column widths, merges and style ids once at import | refused today on DGT, `MYEB2`, `MYEB1` and the ledger; 1.5–2.0 ms per window from the store (values and formats measured; styles not prototyped) | style fidelity of blank styled cells; import-time metadata cost | 100 × 40 cells and 2 MiB per window, `MAX_STYLES`, merges per window, format length, `DOCTYPE` rejection |
| 4 MiB byte cap on `styles.xml`, read whole (F25) | Stream the styles part; protect with the count caps that already exist | TAG (3,947 cell formats, 37,924 named styles): every sheet refused today; not measured after the change | low | `MAX_STYLES` 8,192 per collection, format length, entity rejection, a byte ceiling equal to the worksheet one |
| Data tab copy of every cell, built inside the request under the lock (F18) | Point Datasette at the store's grid tables (immutable SQLite, already the layout Datasette shows); expose provenance as a view; never build in a request | fpna: 161 s and 1,931 MB → none; ATB: a 64 s freeze of the whole server → none. A grouping over the ledger runs in 0.13–0.16 s in SQL on the grid (Polars: 0.04 s) | the `cells` table becomes a view; Datasette URLs change | Datasette's page, row and SQL time limits |
| One `set_value` per input (F07) | Load range inputs in chunks of 1,000–10,000 rows with `set_values_batch` (F19); with data nodes, straight from the grid | `Forecast!D2` 156.9 → 23.5 s; `Summary!B2` 33.1 → 7.2 s; same values and comparisons | value mapping must stay identical to the per-cell path (verified on 1.25 M inputs) | closure budget, worker time and memory |
| Index reads in autocommit on Windows (F24) | One read transaction per `WorkbookIndex` read connection | neighbourhood 5–9× faster on Windows (0.57–0.98 → 0.10–0.16 s); large closures about 30 % | low; readers see one snapshot (stricter than today) | — |
| Hierarchy graph holds every input node | One node per data region | today fpna `build_graph` 74.7 s / 215 MB, ONS 274 s / 783 MB for no formula at all; with data nodes ONS has 4 nodes (not prototyped in the graph) | changes aggregate counts (also touched by S01 in batch 3) | fan-out 32, page sizes |
| Calculation closure 20,000 nodes; library 100,000 entries and numeric cells per workbook | Keep for now | not measured in this batch | — | — |

### Readers

Keep serve's streaming reader for import. calamine, fastexcel and Polars read
the large tables 6–12× faster than the reader's own parse floor, but they
decode the whole sheet in memory (DGT: 1.45–1.56 GB, about 2 GB committed,
over a task's default 1,024 MiB), a row limit or a column selection does not
reduce that, they give no number format, style or stored cell type, and a
damaged shared-string or styles part makes them refuse the whole workbook.
The data node's gains come from storage, not from the reader. fastexcel and
Polars would be new dependencies (licences and sizes in section 6) for a gain
that SQL on the grid tables mostly covers; this note does not propose them.
python-calamine stays where it is today, in the `analyze` pipeline
(`loader.py`, `external.read_workbook_values`); serve's import does not use
it.

### Integration of the data node, in independent steps

1. **Now, small and local**, each with its own behavioural test: chunked
   `set_values_batch` in `lazy._evaluate_index` (F07, F19); a join driven by
   the selection in `selected_inputs` (F26); one read
   transaction per `WorkbookIndex` read connection (F24); a workbook
   diagnostic when a declared part is absent (F17); clearer messages for
   quota stops and missing parts (F22); streamed styles in the presentation
   (F25).
2. **Data tab off the request path** (F18): until step 4, build the copy in a
   worker task with a visible status instead of inside the request.
3. **Store and data node at import** (index schema 5): `build_index` keeps
   `_read_source` and sends literals to per-sheet grids; `WorkbookIndex`
   synthesises an input node from the grid and the per-sheet provenance;
   regions are detected per sheet and stored as `data` nodes; range nodes get
   an edge to the data node with the sub-range read; publication per sheet
   (F21). The calculation path loads region sub-ranges in chunks.
4. **Views from the store**: Datasette on the grids (drop
   `build_source_database`), presentation windows by range query with
   indexed styles and dimensions (F08), the hierarchy graph with one node per
   region.
5. **Batch 3**: highlights of the sub-range a formula reads (plan, batch 3
   criterion), S01 on the hierarchy view.

### Step 3 as built (batch 2 ter, step A)

Adopted on 15 November 2026, with these choices:

- **Grid.** `cell_grid.py` keeps a literal input in its sheet's grid only when
  the node rebuilt from it equals the reader's node field by field, except
  `rawXmlValue`. Formulas, array members, cells with a diagnostic, unreadable
  or missing values, ISO dates (`t="d"`) and cells out of row order stay
  records. A supplied external workbook's temporary index keeps every cell a
  record, so a copied external cell keeps its XML spelling.
- **Order.** Reads keep the import's order: names, each sheet's cells row by
  row, then nodes added afterwards. For worksheet XML in row order, as Excel
  writes it, this is the old storage order; a hostile file listing cells out
  of order now reads them row by row.
- **Migration.** The file name carries the schema (`index-v5.sqlite`); an older
  version finds no index and asks for an import instead of reading the grids
  as blanks. A schema 4 index is never converted: it stays on disk, unused,
  until the user reimports the project, which deletes it with its derived
  views. The notice is in the two languages of the `serve` interface.
- **Regions.** The rule of the prototype with one change: only a formula
  between a column's first and last values makes it mixed; a title or total
  row above or below no longer excludes it. A region needs 10,000 saved
  values, counted: measured as rows times columns, a column with values in
  rows 1 and 1,048,576 became a region.
- **Per-sheet publication (F21).** Half of the storage left at the start for
  reading; the sheet in progress is rolled back and marked `not_imported`
  with its reason. As first built, reading had 60 % of the time budget and
  the budget was checked only while sheets were read: on ATB (233,000
  formulas, reading 29 s, dependencies 93 s), budgets of 25, 60 and 120 s
  timed out in the dependency extraction and published nothing (measured on
  17 November).
- **Every phase within the time budget (17 November, after that
  measurement).** The import's own work ends 5 s before the task's end,
  counted from the coordinator's start; the 5 s cover links, VBA, counts and
  publication (measured 0.2 to 1.3 s). Reading stops when the time left no
  longer covers the analysis at its fastest, 0.30 ms per formula with no
  margin, of the sheets read whole (a first reserve of 1.25 x 0.65 ms left
  29 to 40 % of the energy model's budget unused, measured on the corpus on
  17 November) (the sheet being read does not count its own formulas, so a sheet
  that can be read is never rolled back for them). The analysis
  checks the deadline before each formula; the formulas it does not reach
  keep the diagnostic "Formula not analyzed", written by the reader in every
  formula and name node and removed by the analysis, so nothing is written
  at the deadline (rewriting 150,000 nodes then would take 8 s) and every
  read of a node sees it. Their sheets are `partial` with the reason
  `time_budget`; a calculation reaching such a formula is refused. Data
  regions are described per sheet only when their predicted time fits. The
  proposed budget counts reading, analysis and regions, at the rates measured
  on the import when there was enough work, and says when it is uncertain.
  What remains true: a sheet too large to read within the budget is still
  only "not imported"; links, VBA and publication cannot be stopped, and an
  import whose final steps outlast the 5 s still times out without a new
  index (the error then proposes twice the budget, not an estimate);
  loading a supplied external workbook during the analysis is not
  interrupted. Verified on generated workbooks; the real-corpus runs (ATB at
  25, 60 and 120 s, DGT and fpna at 600 s) await approval.

### Decisions for Philippe

- Adopt the store and data node (step 3) as the storage of literals, with the
  re-import it implies?
- Schedule the small step-1 corrections, and in which batch?
- Confirm that fastexcel and Polars are not adopted.
- Report F19 (superlinear `set_values_batch`) and the reader limits (F20)
  upstream, or not?

## 10. Reproduction

All scripts are in `prototypes/data_node/probes/` and take their paths as arguments; the corpus and every output stay outside
the repository (`validation_screenshots/` is ignored). From the repository
root, with `$W` the corpus folder and `$O` an output folder:

```powershell
# Import cost (F06), one fresh process per workbook
uv run python prototypes/data_node/probes/measure_import.py $W\public\dgt-elections-2017-2020-tour1.xlsx $O\dgt.sqlite $O\dgt-import.json
uv run python prototypes/data_node/probes/profile_index.py $W\public\nrel-atb-2024-v4.xlsx $O\atb.sqlite $O\atb-profile.json
# Readers (fastexcel and Polars from a separate throw-away environment)
uv run python prototypes/data_node/probes/reader_bench.py calamine full $W\public\dgt-elections-2017-2020-tour1.xlsx Feuil1 0
<venv>\Scripts\python prototypes/data_node/probes/reader_bench.py fastexcel head1000 $W\public\dgt-elections-2017-2020-tour1.xlsx Feuil1 0
# formualizer loading (F19)
uv run python prototypes/data_node/probes/probe_fz_chunks.py $W\generated\large\fpna-20mb-seed42.xlsx $O\chunks.json
# Current views on a built index (F08)
uv run python prototypes/data_node/probes/probe_views.py presentation <xlsx> <index> <sheet> 1 $O\view.json
uv run python prototypes/data_node/probes/probe_views.py datasette <index> $O\source-data.sqlite $O\datasette.json
# The serve path end to end, UI default budgets (failure modes, F21)
uv run python prototypes/data_node/probes/serve_path.py <case> <xlsx> <empty data dir> $O\case.json cell=Sheet!A1
$env:LINEXCEL_LAB_WORKBOOKS = $W; uv run python prototypes/data_node/probes/make_damaged.py $O\damaged
uv run python prototypes/data_node/probes/probe_data_lock.py $W\public\nrel-atb-2024-v4.xlsx <empty data dir> $O\lock.json
# Prototype: see prototypes/data_node/README.md (build, windows, verify, link, evaluate)
```

`verify` samples with the fixed seed 20261114. Peak memory is the Windows
peak working set (`probes/winmem.py`); on Linux the scripts report
`ru_maxrss`.
