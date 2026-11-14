# Critical review of linexcel 1.19.2 (batch 1)

Date: 2026-11-14. Starting point: `v1.19.2` (`da1748f`), branch
`claude/audit-agents`. Read-only review: no file under `src/`, `frontend/` or
`tests/` was changed.

This review challenges the existing code with reproducible evidence. Each
finding is marked **verified** (reproduced by a command or probe given below)
or **suspected** (code reading or indirect evidence only). Severity describes
the effect on a `linexcel serve` user: *high* means a wrong or unusable result
is presented as reliable, or a core workflow fails; *medium* means a wrong but
visibly qualified result, or a needless refusal; *low* means cost or polish.

## 1. Scope, environment and limits

- Environment: Linux cloud clone, Python 3.13.18, `uv`, Node 22, npm 10.9.
  formualizer **0.14.0** installed, locked and latest on PyPI (no newer
  release to compare). python-calamine 0.9.1, openpyxl 3.1.5.
- Oracles available: formualizer itself, LibreOffice (headless, as a second
  opinion only), openpyxl. **Excel was not available.** Where a finding depends
  on Excel semantics, it says so and is listed for batch 7.
- **The private `linexcel-lab` corpus was not available** (the 103 MB and
  50 MB data tables, the TAG and ATB models). Measurements use the generated
  validation workbooks, `tests/` fixtures and probe workbooks generated locally
  with openpyxl, including a formula-free 100,000 × 10 data table (6.6 MB
  `.xlsx`, 44.8 MB worksheet XML). All probe files live in the ignored
  `validation_screenshots/lot1/` directory; the main probe scripts are
  reproduced in the appendix.
- Depth was given to the parts batches 2 and 3 touch: OOXML import
  (`lazy_store.py`), targeted calculation (`lazy.py`), lineage and graph
  neighbourhoods, the Data tab (`excel_presentation.py`), resilience. AI
  documentation, questions, VBA, Power Query, reports and catalog search were
  not reviewed in depth (see section 6).

## 2. Baseline (batch 0)

| Check | Result |
| --- | --- |
| `git describe --tags --always` | `v1.19.2` |
| formualizer after `uv sync` | 0.14.0 (declared `>=0.10.1`, locked 0.14.0, latest PyPI 0.14.0) |
| `uv run pytest` | 5331 passed, 59 skipped, 537 deselected, 4 xfailed, 0 failed, 318.65 s |
| skipped tests | 59 (breakdown below) |
| the 11 browser test files, with `--extra screenshots` and Chromium installed | 381 passed, 0 skipped, 354 s |
| `npm ci` then `npm run check` (frontend) | 0 vulnerabilities; "Committed frontend bundle matches sources." |

Exact skip breakdown from `pytest -rs`: Playwright not importable in a plain
sync (39 tests in 11 files), `openai` not importable (4), `opentelemetry` not
importable (1), Windows-only process, memory and path contracts (9),
PowerShell marshaling (1), real `libproc` enforcement (2), doctest `+SKIP` (3).
With the `screenshots` extra and `playwright install chromium`, the 11 browser
files pass completely.

## 3. Findings register

### 3.1 Verified findings

| ID | Subsystem | Finding | Evidence (reproducible) | Severity | Recommendation |
| --- | --- | --- | --- | --- | --- |
| F01 | Lineage edges, targeted calculation | Position-only references are stored as value dependencies. `=ROW(A1)` in A1, `=ROWS(D1:D10)` in D1, `=ADDRESS(ROW(E1),COLUMN(E1))` in E1 and `=COLUMNS(A1:G1)` in G1 become self-loops, so targeted calculation refuses them: "Circular dependencies require valid workbook iteration settings." `=ROWS(C2:C11)` is shown as depending on the *values* of C2:C11. formualizer 0.14.0 itself returns `#CIRC!` for the four cells; LibreOffice computes 1, 10, `$E$1`, 7. | `make_position.py` + `probe_lazy.py` (appendix): edges `S!A1→S!A1`, `S!E1→S!E1`, `S!D1:D10→S!D1`, `S!A1:G1→S!G1`; `evaluate_node` status `unsupported`. `probe_analyze.py`: formualizer `get_value` = `{'kind': 'Circ'}`. | medium (fail-closed: no false value, but false lineage and needless refusal) | **correct**, but not locally: the edge model must distinguish value and position references (`lazy_store.py`, `lazy.py`, catalog), and removing the false cycle alone would expose the engine's `#CIRC!` as a false "difference". Needs a pre-evaluation rewrite of position-only arguments to constants, or an upstream fix. |
| F02 | Graph neighbourhood (serve) | An input read only through a range has no neighbour in the serve graph. Range nodes carry no edges from their member cells, and `graph_neighborhood` reads the `edges` table only. `S!C5`, read by `S!C1 =ROWS(C2:C11)` and `S!I1 =SUM(C2:C11)` and directly by `S!J1 =C5*2`, has neighbourhood {C5, J1} at depth 2, with `truncated: false`. | `probe_usages.py lo/position.xlsx 'S!C5'` → `['S!C5', 'S!J1']`. The library lineage (`catalog_lineage.py`) does handle membership through `lineage_ranges` (code reading). | medium today (the UI calls the view a "limited preview"); **high for batch 3** (usages highlighting) | **correct**: one membership-aware "usages" query shared by the graph, the library and the future grid highlights. |
| F03 | Reference extraction (serve import) | One failing reference aborts extraction of every later reference in the same formula, under the misleading diagnostic "Formula tokenization failed". `=B!A1+A:B!A1+C!A2` keeps only `B!A1`; `=A1#+C!A3` (undeclared spill) keeps no dependency at all. Calculation is refused (fail-closed), but `C!A2` and `C!A3` lose their usage edge with no diagnostic on their side. | `make_misc.py` + `probe_lazy.py lo/misc.xlsx 'A!C2' 'A!C3'`. Cause: the per-formula `try … except Exception` in `_index_dependencies` (`lazy_store.py` ≈1457–1579) wraps the whole token loop. | medium | **correct**: handle each reference separately, record an explicit unresolved node, keep the remaining references. |
| F04 | Reference extraction (serve) | `LET` variable names become `unresolved:` nodes, and the calculation is refused although formualizer evaluates `LET`. The static `analyze` path supports `LET`, and the tests (`test_local_scope_steps.py`) cover only that path. | `probe_lazy.py lo/stress.xlsx 'Lookups!B10'`: deps `unresolved:Lookups:total`, `unresolved:Lookups:n`; diagnostics "Unresolved dependency: n". Native: `=LET(x,A1,x*3)` → 6.0. | medium | **correct** in the serve extractor; add a serve-path test. |
| F05 | Calculation semantics | Plain `<f>` formulas (no dynamic-array metadata) are evaluated with dynamic-array semantics, without implicit intersection. Results are then reported as differences from the saved cache. LibreOffice, which applies legacy implicit intersection, agrees with the caches. | `implicit.xlsx` (A1:A5 = 1..5, recalculated by LibreOffice). Serve: `E3 =A1:A5*10` → 10, cache 30, `comparison: different`; `C2 =SUM(A1:A3*2)` → 12, cache 4, `different`; `B3 =A:A` refused ("engine limitation"). `analyze`: `B3` value 1.0 from engine, warning "recalculated 1.0 differs from file value 3.0". No `cm`/implicit-intersection handling found (`grep` in `lazy_store.py`, `lazy.py`, `excel_compat.py`, `engine.py`). | **high** (a false "difference observed" presented as a calculation fact). Excel semantics to be confirmed in batch 7. | **correct**: read the dynamic-array metadata; for legacy formulas either apply `@` semantics or refuse with a named diagnostic. Never report "different" on these cells until then. |
| F06 | Import storage (serve) | One SQLite row with a zlib-compressed JSON blob per literal cell. Each blob repeats the 64-character workbook SHA-256, worksheet part, sheet, address, raw XML value and context. A formula-free 100,000 × 10 table (6.6 MB `.xlsx`) takes **80 s** to import and produces a **522 MB** index (≈80×). Peak RSS stays at 60 MB. For comparison, python-calamine reads it in 1.0 s (140 MB), formualizer loads it in 0.6 s (164 MB), openpyxl read-only in 7.6 s (44 MB). | `prof_index.py {index,calamine,formualizer,openpyxl_ro} big100k.xlsx`. cProfile (154 s with profiler overhead): `_pack` 52 s (JSON 14 s, `zlib.compress` 25 s), `cell_source_fields` 17 s, `ElementTree.find` 20 s. A stored number `5` takes 344 bytes compressed, 765 bytes of JSON. | **high** (usability on real data tables; disk quota) | **rewrite** the storage of literal regions (data node, columnar or per-sheet provenance); keep the streaming reader. Batch 2. |
| F07 | Targeted calculation cost | Range inputs are streamed into the engine one `set_value` call per cell. `=SUM(Data!G:G)` over 100,000 rows takes 13.5 s; `SUMIFS` over two columns 30 s. Natively, `load_workbook` + the same `SUM` takes 0.67 s, then `SUMIFS` 0.01 s. In isolation, 100,000 `set_value` calls take 5.6 s against 1.5 s for one `set_values_batch`. | `prof_eval.py big100k.xlsx big100k.sqlite 'Calc!A1' 'Calc!B1' 'Calc!C1'`; `prof_fz.py big100k.xlsx`. Code: `lazy.py` 859–865. | medium | **correct** (batch loading) now; reconsider with the data node in batch 2. |
| F08 | Data tab (`excel_presentation.py`) | Every window request re-parses the whole worksheet XML to collect styles, dimensions and merges, and refuses any sheet with more than `MAX_ELEMENTS` = 1,000,000 XML elements. The first 60 rows of the 100,000-row table cannot be displayed ("Feuille trop complexe pour cet aperçu statique", 1.7 s, same at row 99,000). Error texts are hard-coded in French in Python. | `sheet_presentation(Path('big100k.xlsx'), Path('big100k.sqlite'), 'Data', row=1, rows=60, columns=10)` → `PresentationLimitError`. | **high** (Data tab unusable on medium tables) | **rewrite**: index styles, dimensions and merges once at import; serve windows from the index. The cap is a convenience cap (batch 2 inventory). |
| F09 | Import resilience | A truncated `styles.xml`, or one defined name with a non-integer `localSheetId`, makes the whole import fail. A truncated or missing worksheet stays local (sheet marked `partial`, other sheets complete), and so does an entity-expansion payload. | `probe_damage.py sales.xlsx` and `probe_damage.py lo/misc.xlsx`: `truncated_styles: WHOLE IMPORT FAILED ParseError`, `bad_localSheetId: WHOLE IMPORT FAILED ValueError`, `truncated_sheet2: OK status=partial`. `probe_entities.py`: sheet B `partial` ("limit on input amplification factor"), A and C complete, 0.23 s. | medium | **correct**: workbook-level parts (styles, names, relationships) degrade to diagnostics with safe defaults. |
| F10 | Architecture | Two independent lineage and calculation pipelines. `analyze` (static HTML report: `analyzer`, `engine`, `graph`, `resolver`, `loader`, `decompose`, `result`, `structure`, `sweep`, `recovery`, `viewer`, `insights`: 7,661 lines) loads the workbook with openpyxl and the full engine. `serve` (`lazy`, `lazy_store`, `graph_store`, `excel_presentation`, `catalog*`, `web*`, `link_facts`, `external_index`: 15,217 lines) has its own reader, extractor and closure. They disagree on the same cells: `=OFFSET(F2,1,0)` has source F3 in serve but F2 only in `analyze`; `LET` works only in `analyze` (F04); `=A:A` in B3 is refused by serve and computed as 1.0 by `analyze` (F05). | `probe_analyze.py lo/position.xlsx …` (edges `i:S!F2→c:S!F1` only) vs `probe_lazy.py` (edges `S!F2`, `S!F3 → S!F1`); F04, F05. | **high** (maintenance; contradictory answers) | **rewrite / supprimer**: keep one core (the serve index and targeted calculation) and rebuild the static report on it, or drop it. |
| F11 | Redundancy with formualizer | References are extracted with openpyxl's tokenizer plus linexcel regexes. A differential against formualizer 0.14 `Workbook.precedents()` agrees on 50/50 formula cells of the sales workbook and 545/566 of the stress workbook. Of the 21 differences, 20 are representation or capability, not linexcel errors: the engine trims whole columns to the used range (10), leaves table (2) and external (2) references unresolved, linexcel alone resolves `INDIRECT`/`OFFSET` (3) and a 3D span (1), and `precedents()` raised `InvalidInspectionAddressError` on 2 cells (a 19,176-character formula of 700 `ROUND` terms, and the DDE text `=cmd\|'/c calc'!A1`). The remaining one is `LET` (F04), where the engine is right. `formualizer.visitor.collect_references` is weaker than linexcel (treats `LET` names as references, splits `F:F`). | `diff_deps.py sales.xlsx`, `diff_deps.py stress.xlsx`; `collect_references` probe in the appendix. | low | **keep** the extractor (after F03/F04); add a contract test that diffs it against `precedents()` on the validation workbooks so drift is visible. |
| F12 | Shared formulas | Shared-formula translation (openpyxl `Translator`) is correct on mixed anchors, whole columns and rows, quoted sheet names, string literals containing `A1`, intersections and function names such as `LOG10`: identical to formualizer and LibreOffice on 30 cells; all members of each master share one pattern id. | `make_shared.py` + `probe_shared.py shared.xlsx lo/shared.xlsx`. | — | **keep**. |
| F13 | Number formats | `format_display` is a deliberately small subset and labels every fallback `unsupported`. Common formats fall back to the raw value: `dd/mm/yyyy;@` (shows `45000.5`), a three-section format with a text zero section, `#,##0,`, fractions, `"$"#,##0` for negatives. | `format_display` probe in the appendix. | low (honest fallback) | **keep** the fail-closed principle; extend against an oracle (batch 6, `numfmt`). |
| F14 | Hostile XML | The import relies on the bundled expat (2.9.0 here) for entity-expansion protection; it works and the failure stays local (F09). `excel_presentation`, `structured_refs`, `doc_evidence` and `insights` additionally reject `DOCTYPE`/`ENTITY` themselves; `lazy_store` does not. | `probe_entities.py`; `grep -n DOCTYPE src/linexcel/*.py`. | low | **keep**, unify into one XML-reading helper with one policy. |
| F15 | Tests | No test covers F01, F02, F03, F05 or the serve path of F04. Many presentation tests assert on generated text rather than behaviour, for example `tests/test_viewer_ui.py::TestInlineStylesLiftedToCss::test_only_data_driven_style_assignments_survive` asserts the exact set of `.style.X =` property names found by regex in the HTML. | `grep -rn "ROW(A1)\|ROWS(" tests/` (only table/external contexts); `grep -rn "LET(" tests/` (analyze path only). | medium | **correct**: add behavioural tests for the verified findings first; treat regex-on-HTML tests as candidates for removal in a new repository. |
| F16 | Import memory and atomicity | The streaming import keeps peak memory low (60 MB on a million cells), publishes the index atomically (`os.replace` of a temporary file), and keeps worksheet failures local. | F06 and F09 measurements; `build_index` in `lazy_store.py`. | — | **keep** (principles and reader), independently of F06. |

### 3.2 Suspected findings (not reproduced here)

| ID | Subsystem | Suspicion | Why not verified | Next step |
| --- | --- | --- | --- | --- |
| S01 | "Explore the workbook graph" (`graph_store.py`) | Aggregated dependency counts come from the `edges` table only, so they share F02's blind spot for range members. | Code reading only (`graph_store.py` has no `lineage_ranges` or bounds lookup). | Probe `build_graph`/`read_view` on `position.xlsx`. |
| S02 | `format_display` | `45000.99999999` with `yyyy-mm-dd hh:mm:ss` renders `2023-03-15 23:59:59` labelled `formatted`; Excel is believed to round to the displayed second (`2023-03-16 00:00:00`). | No Excel; LibreOffice CSV export does not apply the format. | Batch 6/7 oracle. |
| S03 | Position-reference semantics in Excel | Excel is believed not to flag `=ROW(A1)` in A1 or `=COLUMNS(A1:G1)` in G1 as circular (LibreOffice computes them). | No Excel. | Batch 7 sample. |
| S04 | Legacy formulas in Excel 365 files | F05 assumes Excel applies implicit intersection to plain `<f>` formulas (Microsoft's "implicit intersection operator" documentation). | No Excel. | Batch 7 sample: `implicit.xlsx`. |
| S05 | Data tab cost | Even under the element cap, each window costs a full re-parse of the sheet XML (O(sheet size) per scroll step). | Only the refusal path was timed (1.7 s for 44.8 MB). | Time a sheet just under the cap. |
| S06 | Large AI/doc modules | `aidoc.py` (2,992 lines), `i18n.py` (2,380), `web.py` (2,196, regex routing in one handler) mix concerns; `CHANGELOG.md` is 122 KB and `docs/` carries plan and history documents. | Not reviewed in depth (outside the batch 2/3 priority). | Review when the AI features are next touched. |

## 4. Beliefs checked (formualizer skill)

| Belief | Source | Evidence | Verdict |
| --- | --- | --- | --- |
| The installed formualizer is older than the lock (0.13.0 vs 0.14.0). | Plan (Codex note) | `uv sync` then `importlib.metadata.version` → 0.14.0; inventory: installed = locked = latest 0.14.0. | refuted for this environment |
| formualizer can load a values-only range or table as one unit, without one engine cell per value. | Plan (batch 2 hypothesis) | 0.14.0 stubs: no region or provider API. `load_workbook(strategy=…)` is documented "accepted for backward compatibility; loading is eager". Available instead: `set_values_batch` (bulk values, still cells; 1.5 s vs 5.6 s per-cell for 100k), `add_table` (table metadata over existing cells), `range_page` (paged read without evaluation), `EvaluationConfig.span_evaluation` (experimental, for copied *formula* spans). | refuted in 0.14.0: a data node must live in linexcel; the engine only offers faster bulk loading. |
| `=ROW(A1)` in A1 is not a cycle for the engine. | Codex conversation | `load_workbook` + `evaluate_all` → `#CIRC!` for A1, D1, E1, G1 of `position.xlsx`; `precedents()` reports the reference as `Declared`. | refuted (engine treats position references as value precedents) |
| formualizer's parser helpers could replace linexcel's reference extraction. | Review hypothesis | `visitor.collect_references` treats `LET` names as references and splits `F:F`; `Workbook.precedents()` is semantic and correct on `LET` but needs the workbook loaded in the engine. | partly: `precedents()` is a good oracle, not a drop-in import-time extractor. |
| formualizer evaluates `LET`. | Review hypothesis | `=LET(x,A1,x*3)` with A1 = 2 → 6.0. | confirmed (0.14.0) |

## 5. Map of the core worth keeping

For a possible new repository starting from 1.19.2 without history.

**Keep (reuse the ideas and, after review, the code):**

- The streaming OOXML reader principles in `lazy_store._read_source`: native
  ZIP/Expat streaming, element detaching, bounded memory, per-worksheet failure
  locality, `ST_Xstring` decoding, shared-formula translation, date-system
  handling (F12, F16).
- The fail-closed targeted calculation contract in `lazy._evaluate_index`:
  exact static closure, refusal rather than partial ranges or substituted
  caches, explicit coverage counts, deterministic clock and seed, iteration
  only with valid workbook settings.
- Cache comparison with explicit tolerances and three outcomes (`values.readings_agree`,
  `temporal_serials_agree`), and the provenance vocabulary (saved input,
  saved cache, targeted recalculation, unknown).
- R1C1 pattern grouping (`lazy._pattern_signature`), dynamic-reference
  rewriting with evidence (`dynamic_refs.py`), structured-table resolution
  (`structured_refs.py`): verified on the validation workbooks (F11).
- Worker isolation and budgets, atomic index publication, CSP and offline
  operation, upload limits against hostile files.
- `format_display`'s rule "label every fallback" (F13).

**Rewrite:**

- Storage of literal cells: one data node per formula-free region, with
  per-sheet (not per-cell) provenance (F06, F07).
- The edge model: separate value dependencies, position references and range
  membership, with one membership-aware usages query (F01, F02, F03).
- The Data tab window: styles, dimensions and merges indexed once (F08).
- Workbook-level part handling so that styles, names and relationships fail
  locally (F09).
- One XML-reading helper with one hostile-input policy (F14).

**Drop or merge (decision for Philippe):**

- The second pipeline (`analyze` and its modules, 7,661 lines) as an
  independent implementation: rebuild the static report on the serve core or
  retire it (F10). Its standalone Cytoscape viewer (`viewer.html`, 3,334 lines,
  plus Cytoscape assets) duplicates the React Flow graph of `serve`.
- Tests that freeze generated text rather than behaviour (F15), and the
  accumulated plan and history documents, which a new repository need not carry.

## 6. Not done here

- Anything requiring the `linexcel-lab` corpus: the 103 MB and 50 MB tables,
  TAG and ATB models. The 100,000 × 10 generated table is a stand-in, not a
  substitute; real tables may have different string ratios and styles.
- Anything requiring Excel: S02–S04, and confirming F05.
- `validate_manual.py` and the visual checks of `AGENTS.md`: not required for
  this read-only batch and not run.
- No in-depth review of AI documentation, questions (`ask*`), VBA, Power Query,
  reports and the library search; no profiling of the library catalog.

## 7. Effect on the following batches

| Batch | Status | Change |
| --- | --- | --- |
| 2 — Reading, resilience, data node | **stays, strengthened** | F06–F08 give a baseline (80 s / 522 MB import, 13.5 s `SUM`, Data tab refusal). formualizer 0.14.0 has no values-only region API, so the data node lives in linexcel; the engine side needs at least `set_values_batch`. Add F09 to the failure-mode table. |
| 3 — Audit highlights in the grid | **stays, but depends on fixes** | Usages require range membership (F02); sources are incomplete after a failed reference (F03). The "engine" part (position references) is a broad correction in the plan's terms (F01: crosses `lazy_store`, `lazy`, catalog, and the engine returns `#CIRC!`): describe, do not fix inside batch 3. The Data tab cannot host highlights on medium tables until F08 is addressed. |
| 4 — Agent skills | **stays** | The library lineage already handles range membership. Skills must not present a cache "difference" as an error until F05 is fixed, and should surface refusals (F01, F04) as coverage limits. |
| 5 — Version comparison (design) | **stays** | Storage cost (F06) doubles with two revisions; the data node design should come first or be assumed. Add fixtures for legacy vs dynamic formulas (F05). |
| 6 — Univer bench, `numfmt` oracle | **stays** | Add position references (F01), implicit intersection (F05) and the S02 rounding case to the bench. |
| 7 — Excel as oracle | **moves earlier** | It is the only way to settle F05, S02–S04, which decide whether today's "different" verdicts are true. The sample is small and ready (`position.xlsx`, `implicit.xlsx`). |

Recommended order (Philippe decides): a short lineage-correctness batch (F02,
F03, F04, F09, and a refusal for F05 legacy formulas), then the Excel oracle on
the F01/F05/S02 sample, then batch 2, batch 3, batch 4; batches 5 and 6 after.

## 8. Batch 1 bis status (lineage correctness)

Date: 2026-11-14, same environment as above (no Excel, no `linexcel-lab`
corpus). Each finding was first reproduced with its probe, then covered by a
behavioural test on the serve path (`tests/test_serve_lineage_correctness.py`,
through the HTTP coordinator) that failed before the fix.

| ID | Status | Commit | What changed | Remaining limits |
| --- | --- | --- | --- | --- |
| F02 | **fixed** (serve neighbourhood) | `a991d3c` | `graph_neighborhood` treats range membership as a relation, read from `lineage_ranges` (fallback for older indexes): a cell reaches the ranges that contain it, a range reaches its indexed members. Membership is returned as a `range` edge (member → range). Each relation family is bounded by the node budget, so `truncated` stays exact. Probe: `S!C5` now reaches `S!C2:C11`, `S!I1`, `S!C1`. | Cost on the 100,000-row table (warm, `big100k.sqlite`): depth 1 from `Data!G5` 1–2 ms (unchanged), depth 2 60–100 ms against 2 ms before, because 200 nodes are now returned (`truncated: true`) instead of 1; after dropping the page cache, 190 ms, dominated by decoding the 200 returned nodes and their context labels. One first cold call took 8.4 s and was not reproduced. `web_report` relation counts still count formula edges only. The canvas legend reads "→ calculation dependency" for every arrow, including membership; the detail list labels it as membership. |
| S01 | **verified, not fixed** | — | `build_graph`/`read_view` on `position.xlsx`: `S!C5` shows one edge (to `J1`, outside the view) and no link to `S!C2:C11` in the same view (`validation_screenshots/lot1bis/probes/probe_s01.py`). | Same blind spot as F02 in the hierarchical view, which is the default Graph view of the application: there `S!C6` shows no link while `S!C2:C11` points to the rest of the workbook; **Back to neighbourhood** shows the fixed F02 view. Fixing it changes the aggregated edge counts and the `edgeCount` consistency check of `build_graph`; left for batch 3, which needs it for highlights. |
| F03 | **fixed** | `fe92c25` | Each reference is resolved on its own. A failure adds a diagnostic and a capability diagnostic naming the reference and the real reason (`Reference A:B!A1 could not be resolved: 3D references require …`); later references keep their edges. Calculation of the formula stays refused. | No `unresolved:` node is created for the failed reference, so it does not appear in the graph; it appears in the node's diagnostics. |
| F04 | **fixed** | `3e2d4da` | The serve extractor reuses `resolver._bound_names` (the `analyze` rule) to skip names a formula binds itself. `_xlfn.`/`_xlpm.` prefixes are removed before parsing: the `analyze` rule alone does not recognise the spelling Excel saves (probe below). Excel-spelled `LET`, nested `LAMBDA` and the stress workbook's `LET` now calculate. | A bare bound name that is also a defined name keeps the defined-name dependency (a superset, never a hidden source). The rule is formula-wide, not lexically scoped, as in `analyze`. Serve now imports a private function of the `analyze` pipeline (see issue #148). |
| F09 | **fixed** (import) | `39f74c5` | Unreadable styles become a workbook diagnostic on `xl/styles.xml`; the existing per-cell check then keeps the value of every styled cell unknown. A defined name with a `localSheetId` that is not a valid sheet position is skipped with a diagnostic naming it; formulas reading it are refused. A negative position no longer selects the last sheet. The index is marked partial; other sheets stay complete. | With damaged styles, the Data tab of that workbook answers 400 "Présentation de la feuille indisponible" (local to the tab; F08/batch 2). The workbook notice uses the generic advice "Consultez la formule et ses sources …", which does not fit an import diagnostic; the technical details name the part. |
| F05 | **partly fixed** (named refusal, semantics unchanged) | `9bac0f8` | `implicit_intersection.py` flags a multi-cell reference read in a scalar position (whole formula, operator operand, function argument not known to accept a range, outside `SUMPRODUCT`) in an ordinary formula, on the expression handed to the engine. If such a formula in the closure is not equal to its saved cache, the result is refused with a named diagnostic (`implicit_intersection_unsupported`) and no step keeps an engine value. | Measured on the generated validation workbooks: 1 verdict change in 616 formula cells (`'Cross Refs'!B11 =RevenueColumn`, engine 21.75 against a LibreOffice cache of 62.25); no "equal" became a refusal. `implicit.xlsx`: 3 "different" become refusals, `B3` keeps its refusal with a named reason. Not detected: multi-cell references returned by functions not rewritten by the index (`INDEX(…,0)`, `CHOOSE`, `IF`). Unknown functions are treated as not accepting ranges (refusal, never a value). The Excel semantics wait for batch 7. |

Other observations made during this batch (verified unless noted):

- `resolver._bound_names` (analyze) returns no bound name for
  `=_xlfn.LET(_xlpm.x,A1,_xlpm.x*3)`, the spelling Excel saves, and reports
  `_xlpm.x` as a reference (`probe_let_engine.py`). formualizer 0.14.0
  evaluates both spellings (6.0) and `_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(…),_xlpm.f(A1))`
  (3.0). The `analyze` path was not changed.
- The probe case `missing_sharedStrings` is vacuous on the generated workbooks:
  they store text inline (`t="inlineStr"`) and have no shared-string part.
- The serve application (`app.html`) has French and English interface text
  only; the nine languages of `i18n.py` cover the static report. This batch
  adds no interface text: the new diagnostics are English technical messages,
  shown verbatim in "Détails techniques" like the existing ones.

## Appendix: probes

All probes run from `validation_screenshots/lot1/probes/` with
`uv run --no-sync --project ../../.. python <script> …`. LibreOffice copies
(`soffice --headless --calc --convert-to xlsx --outdir lo <file>`) supply
saved caches; they are a second opinion, not Excel.

Position references (`make_position.py`):

```python
from openpyxl import Workbook
wb = Workbook(); ws = wb.active; ws.title = "S"
ws["A1"] = "=ROW(A1)"; ws["B1"] = "=COLUMN()"
for r in range(2, 12): ws.cell(r, 3, r)
ws["C1"] = "=ROWS(C2:C11)"; ws["D1"] = "=ROWS(D1:D10)"
ws["E1"] = "=ADDRESS(ROW(E1),COLUMN(E1))"; ws["F1"] = "=OFFSET(F2,1,0)"; ws["F3"] = 42
ws["G1"] = "=COLUMNS(A1:G1)"; ws["H1"] = "=ROW()+H2"; ws["H2"] = 5
ws["I1"] = "=SUM(C2:C11)"; ws["J1"] = "=C5*2"; ws["K1"] = "=ISBLANK(K1)"
wb.save("position.xlsx")
```

Serve-path index and targeted calculation (`probe_lazy.py <xlsx> <cell>…`):

```python
import sys, sqlite3, tempfile
from pathlib import Path
from linexcel.lazy_store import build_index, WorkbookIndex
from linexcel.lazy import evaluate_node
wb = Path(sys.argv[1]); idx = Path(tempfile.mkdtemp()) / "i.sqlite"
build_index(wb, idx)
print(sqlite3.connect(idx).execute("select source,target from edges").fetchall())
with WorkbookIndex(idx) as s:
    for c in sys.argv[2:]:
        n = s[c]; print(c, n["formula"], n["dependencies"], n["diagnostics"])
for c in sys.argv[2:]:
    r = evaluate_node(wb, c, index_path=idx)
    print(c, r["status"], r.get("value"), r.get("cachedValue"), r.get("comparison"), r["diagnostics"][:3])
```

Usages through a range (`probe_usages.py lo/position.xlsx 'S!C5'`):

```python
from linexcel.lazy_store import build_index, graph_neighborhood
# build the index as above, then:
n = graph_neighborhood(idx, "S!C5", depth=2)
print(sorted(x["id"] for x in n["nodes"]), n["pagination"]["truncated"])
# -> ['S!C5', 'S!J1'] False
```

Mid-formula extraction failure (`make_misc.py`): sheets `A`, `B`, `C` with
A1:A3 = 1..3 each, then `A!C2 = "=B!A1+A:B!A1+C!A2"`,
`A!C3 = "=A1#+C!A3"`, `A!C4 = "=SUM(Rate*A1, C!A1)"` with `Rate = C!$A$3`.

Implicit intersection (`implicit.xlsx`): A1:A5 = 1..5; `B3 = "=A:A"`,
`C2 = "=SUM(A1:A3*2)"`, `D8 = "=SUM(A1:A3*2)"`, `E3 = "=A1:A5*10"`.

Data table (`make_big.py big100k.xlsx 100000`): openpyxl write-only, header
plus 100,000 rows of id, date, region, product, qty, price, amount, flag, note,
score; sheet `Calc` with `=SUM(Data!G:G)`, `=SUMIFS(Data!G:G,Data!C:C,"North")`,
`=COUNTA(Data!A:A)`. Import timing (`prof_index.py`):

```python
import sys, time, resource, tempfile
from pathlib import Path
t = time.perf_counter()
from linexcel.lazy_store import build_index
idx = Path(tempfile.mkdtemp()) / "i.sqlite"; build_index(Path(sys.argv[2]), idx)
print(f"{time.perf_counter()-t:.2f}s",
      resource.getrusage(resource.RUSAGE_SELF).ru_maxrss // 1024, "MB",
      idx.stat().st_size // 10**6, "MB index")
```

Dependency differential (`diff_deps.py <xlsx>`): for every `kind='cell'` node,
compare the index `dependencies` (names as `NAME:<name>`, references
normalised with `linexcel.refs.parse_ref(...).to_a1()`) with
`formualizer.load_workbook(path).precedents(node_id, max_links=10000)`
(`reference.resolved or address`).

`format_display` sample:

```python
from linexcel.excel_presentation import format_display as f
f(45000.5, "dd/mm/yyyy;@", raw="45000.5")          # ('45000.5', 'unsupported')
f(0, '#,##0;-#,##0;"-"', raw="0")                   # ('0', 'unsupported')
f(45000.99999999, "yyyy-mm-dd hh:mm:ss", raw="45000.99999999")
# ('2023-03-15 23:59:59', 'formatted')  — see S02
```
