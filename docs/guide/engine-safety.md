# Native engine safety and capability evidence

The native engine can terminate its process on deeply nested formulas. A
Python exception handler cannot recover from that termination. Analysis uses
two complementary protections: the public execution policy supervises the
worker process, while the formula admission guard avoids known dangerous
expressions before importing the workbook into the engine.

## Investigating an analysis crash locally

`Analysis crashed during engine evaluation` means the isolated worker exited
without returning a result. The phase includes formula complexity checks, native
workbook import, evaluation and recovery. It does not identify a particular cell
or prove that a formula evaluation caused the exit.

Keep a small execution report without exporting the workbook graph:

```sh
linexcel analyze private.xlsx --no-html --diagnostics diagnostic.json
```

This command uses no AI or screenshots. The report includes `operation`,
`exitCode`, `failure`, runtime `versions`, budgets and native `diagnostic` output.
It excludes the graph and source workbook, but native stderr may contain private
paths, sheet names or formula text: inspect it before sharing. A normal graph JSON
export also contains this evidence under `meta.execution`.

Possible causes and the evidence that distinguishes them:

| Evidence | Interpretation |
| --- | --- |
| `memory allocation of … bytes failed`, Windows insufficient-memory or commitment-limit status | Allocation failure under the configured memory cap (2,048 MiB by default). Large arrays/ranges can use much more memory than the compressed file. The evidence alone does not distinguish the worker cap from machine-wide exhaustion. |
| `has overflowed its stack` or Windows `0xC00000FD` | Native stack overflow. The formula admission guard protects known deep expressions, but does not certify every parser/evaluator path. |
| Rust panic message or `pyo3_runtime.PanicException` | Native engine panic; retain the message and backtrace for a synthetic reproducer. |
| Windows `0xC0000005` or POSIX signal | Native access fault or termination signal; the specific input defect remains to be determined. |
| Windows `0xC0000409`, `SIGKILL`, or another exit without supporting stderr | Exit evidence only. Do not infer memory exhaustion or a particular formula. |

Rust's default [allocation error handler](https://doc.rust-lang.org/std/alloc/fn.handle_alloc_error.html)
prints to stderr and aborts; it need not raise a Python `MemoryError`. A local
Windows/formualizer 0.9.3 reproduction used `=SUM(SEQUENCE(100000,100))` with
`ExecutionPolicy(memory_mb=128)`: native allocation failed during global evaluation
and the worker exited with `0xC0000409`. Linexcel now classifies this as
`memory_limit` from the allocation message, rather than from the exit code alone.
This is a synthetic low-budget reproduction, not a diagnosis of every private
workbook reporting a crash.

For confirmed allocation failures, compare a second local run with a higher
`--memory-mb` value that fits the machine. For evaluation-specific failures, a
local `--target 'Sheet1!A1'` run can narrow the implicated dependency branch;
it still imports the workbook, so it cannot isolate an import crash. Neither
comparison disables process isolation. Increasing `--analysis-seconds` addresses
`timed_out`, not a native crash. No automatic unbounded retry is performed.

Python fault handling and Rust backtraces are enabled in the worker (an existing
`RUST_BACKTRACE` setting is respected). At most 16 KiB of stderr bytes are retained
plus a truncation marker, split between the beginning and end when necessary.
`diagnosticTruncated` identifies omitted middle content. A fatal exit may still
leave no trace, and the operation is the last checkpoint, not a native cell-level
stack trace.

For deterministic fixture checks without AI or rendering:

```sh
python validate_manual.py --workbook both --no-ai --no-vision --no-screenshots
```

This remains an incomplete acceptance run (exit 1). Its `validation.json` records
the actual analysis status separately from the intentionally missing AI and images.

## Formula admission

The guard scans worksheet XML before the first native workbook import. It
uses a lexical nesting bound before calling either `formualizer.parse()` or
`to_dict()`. That conversion itself can overflow the native stack.

Operator chains, repeated unary operators, postfix percentages, powers,
parentheses and reference unions contribute to the bound. Function arguments
are siblings: `SUM(1,1,...)` is not classified like `1+1+...`. Strings with
escaped double quotes, quoted sheet names and structured references are
treated as atomic content. The bound is conservative; rejection does not
prove the formula's mathematical result is wrong or its exact AST depth.

Unsafe formulas are removed from the engine's input copy while their source
formula and stored cache remain in the report. Uncached affected dependents
remain unavailable. Warnings identify the first affected cell and the
complexity limitation. The lexical native-safety ceiling cannot be disabled
by increasing the optional AST evaluation threshold.

The admission scan covers worksheet cell formulas, not every expression in
the package, such as formulas held only in defined names. Process supervision
remains necessary; the guard is not a proof that native crashes are impossible.

Regression cases execute in subprocesses. They include a 3,000-term chain,
unary and power chains, deeply nested parentheses, quoted punctuation, a
3,000-argument flat `SUM`, and a 700-term chain. The latter two are also
recalculated, not merely accepted by the guard.

## Targeted extraction budget

The traced closure is counted per sheet before requesting an evaluation
plan or calling `evaluate_cells`. Exceeding `max_cells_per_sheet` raises an
explicit error without starting target evaluation. Distant coordinates do
not consume the gaps between them.

This bounds the cells represented in the static trace, not all native work.
Dynamic references and an incomplete trace can lead the engine to evaluate
additional precedents. The execution policy supplies the separate process
time and memory limits.

## Capability probes

`meta.engineCapabilities` records a small reproducible synthetic matrix,
including source inputs, formulas, expected and observed values, observed
Python types, engine version and date epoch. Probes run once per worker
process. Passing a probe does not certify a workbook.

With formualizer 0.9.3, the matrix reproduces mixed number/text comparison
differences, including `10<" "` and `"0"=0`, for literals and references.
The 1900-epoch expression `DATE(2026,2,1)-46054` returns a native date where
the expected result is numeric zero. That type difference is preserved in
the evidence, not normalized away. Other tested cases include date
differences, `YEAR`, case-insensitive text equality and legacy normal
distribution functions.

When a probe fails, potentially affected formula nodes and their dependents
receive `verification="unverified_engine_semantics"` and a specific
`semanticRisks` reason. Values, cache comparisons and `valueSource` are
preserved. An annotation indicates a possible engine limitation; it does
not establish that every marked result is false. Dependency propagation is
limited by the lineage actually represented in the graph.

### Independent oracle and coverage

The initial exact two synthetic input workbooks used by the diagnostic were
recalculated in separate LibreOffice copies: 16 probes for each of the 1900
and 1904 date systems. LibreOffice 26.2.4.2 confirmed all **32 expected
results**. Formualizer 0.9.3 agreed on 19 and differed on 13: six mixed-type
comparison or comparison-guard cases in each date system, plus the
1900-epoch date-versus-number result described above. The five source inputs,
their types and the date epochs were checked after conversion. Each
LibreOffice process had a 90-second limit; raw XML, results and version
evidence remain local. LibreOffice is an independent oracle for these cases,
not a certificate of Microsoft Excel compatibility.

The runtime diagnostic now adds four defined-name probes per epoch: numeric
and text constants, a sheet-local constant shadowing a workbook constant,
and a name referring to a cell. Fresh LibreOffice recalculation of the exact
extended diagnostic inputs confirmed **40/40 expected results**. Formualizer
agreed on 21 and differed on 19: the earlier 13 differences and six constant
name failures. Native errors are retained as their raw error dictionaries
and observed Python types. The reference-name probe passes in both epochs.

For the `defined_name_constant` limitation, annotations require a matching
numeric or text constant in complete source name metadata. Sheet-local
definitions take precedence, and local `LET`/`LAMBDA` bindings are respected.
The risk propagates to represented dependents, including error guards; a
generic `#NAME?` result alone does not trigger it. Values and provenance are
preserved. Incomplete metadata, computed name definitions and unrepresented
dependencies remain limits of this diagnostic.

An additional before/after comparison of the initial 32-probe graphs found
no annotation-induced changes to values, dates, provenance, cached values,
cache verdicts, samples or decomposition steps. Twenty of the 32 formula
nodes received conservative risk annotations. That count is not a count of
demonstrated calculation errors, and making risks visible does not repair
the underlying calculations.

The initial matrix deliberately concentrates on a few known limitations. Existing
regression tests cover other contracts, but were not part of this
32-observation LibreOffice comparison:

| Area | Evidence in the 32 probes | Additional regression evidence and limits |
| --- | --- | --- |
| Comparisons | Literal and referenced numbers versus text, a comparison guard, text case and trailing spaces, and ordinary numeric comparison. Operators are `<` and `=`. | `tests/test_semantic_checks.py` checks reporting and conservative risk propagation. This is not a full operand/operator cross-product. |
| Booleans | Boolean results and the cumulative `TRUE` argument of `NORMDIST`; no boolean-operand comparison matrix. | `tests/test_analysis_contracts.py` checks that cache verdicts distinguish booleans from numbers. This tests the report's comparison contract, not general native coercion. |
| Empty values | A single-space string, which is neither an empty string nor a missing cell. | `tests/test_values.py` covers a blank arithmetic operand; `tests/test_quarantine_values.py` keeps an unavailable formula from becoming a false zero. Blank cells, empty strings, whitespace, text zero and numeric zero are not exhaustively compared. |
| Errors and guards | A successful `IF` comparison guard, not a matrix of Excel errors. | `tests/test_values.py` and `tests/test_quarantine_values.py` exercise error/cache verdicts, guarded errors and unavailable-value propagation, including `IFERROR` and `ISERROR`. They do not establish all error precedence rules. |
| Dates and distributions | Two epochs, typed date subtraction, date/serial subtraction, `YEAR`, `DATE` differences, and one case each of `NORMSDIST` and `NORMDIST`. | `tests/test_date_epoch.py`, `tests/test_time_values.py` and `tests/test_legacy_normal_functions.py` add cases; the oracle matrix does not cover all date functions, durations, timezones or distribution parameters. |
| Cycles | None. | `tests/test_iterative_values.py` checks declared iteration settings, convergence under engine criteria, and cache-only or unavailable results after nonconvergence, in full and targeted analysis. It does not prove identical fixed points across spreadsheet engines. |
| Dynamic references | None. | `tests/test_analysis_contracts.py` checks static-trace warnings; `tests/test_quarantine_values.py` checks unavailable values behind `INDIRECT` and `OFFSET`. Complete dynamic dependency closure is not guaranteed. |
| Defined names and local scope | None. | `tests/test_lineage.py`, `tests/test_targeted_boundaries.py` and `tests/test_local_scope_steps.py` cover name resolution/scope, targeted precedents and avoiding false results for local `LET` decomposition. They do not certify all name, `LET` or `LAMBDA` semantics. |

### Versioned extension and unresolved oracle results

`ORACLE_EXTENSION_VERSION = 1` in `tests/test_semantic_checks.py` records 21
additional synthetic cases, repeated in both epochs: **42 observations**.
It includes genuinely absent cells, a formula returning an empty string,
boolean comparisons, error guards and raw errors, `INDIRECT`, `OFFSET`,
defined names and a local `LET` binding. These inputs are separate from the
40 runtime probes and must not be added to them as disjoint coverage.
LibreOffice confirmed 34 of the 42 provisional expectations. Of those 34,
formualizer differed on the six constant-name observations described above.
The remaining eight expectations are explicitly unresolved or disputed:

| Formula, repeated in both epochs | Provisional expectation | Formualizer 0.9.3 | LibreOffice 26.2.4.2 |
| --- | --- | --- | --- |
| `TRUE=1` | `FALSE` | `TRUE` | `TRUE` |
| `FALSE=0` | `FALSE` | `TRUE` | `TRUE` |
| `TRUE>1` | `TRUE` | `FALSE` | `FALSE` |
| `IFERROR(INDIRECT("Missing!A1"),7)` | `7` | `7` | `#REF!` |

Neither engine agreement nor disagreement settles Microsoft Excel behavior
for these cases. Microsoft's [operator documentation](https://support.microsoft.com/en-gb/excel/calculation-operators-and-precedence-in-excel)
describes comparison results and context-dependent conversions, but does
not resolve these three boolean-versus-number examples. The provisional
expectations remain visible instead of being changed to match the oracle.
They do not introduce new runtime mismatch categories.

The first private extension used unprefixed `LET`, which LibreOffice imported
as an unknown name. A separately preserved revision using the OOXML
`_xlfn.LET` spelling produced the expected 15 in both engines. This is an
input-encoding distinction, not a numerical engine correction. Raw native
results, LibreOffice XML, version information, source hashes and both
revisions remain local. The public test cases contain only synthetic data.

Cycles remain outside this scalar oracle: convergence, iteration limits,
unavailable values and cache-only provenance are covered by the explicit
iteration regression contracts above. No arbitrary fixed point is promoted
to a universally correct scalar answer.

## Measured loader change

Error-cache extraction now skips the detailed cell regex when the sheet
contains no error type marker. The matcher also excludes self-closing cells,
which must not inherit a later cell's error value.

The inspected API paths already delegate workbook loading to python-calamine
and evaluation to formualizer. Calamine's `to_python(skip_empty_area=False)`
preserves coordinates; Python handles cache normalization and the error
overlay. Although the installed binding exposes `iter_rows`, an iterator
interface does not establish streaming native allocation or a lower peak
memory footprint. It was therefore not substituted as an unmeasured memory
optimization: the declared-dimension guard and read-only openpyxl fallback
remain necessary. Existing formualizer paths use batched `get_formulas`,
`get_eval_plan` followed by `evaluate_cells` for targeted work, and batched
scratch-expression evaluation. The measured change addresses the Python/XML
overlay rather than replacing these native paths. Python profiles expose
some native boundary calls, but this study did not count every native call
throughout analysis, profile native internals, or separately measure scratch
evaluation cost. No performance claim is inferred from those unmeasured
capabilities.

A local synthetic benchmark used exactly the same workbook before and after:
20,000 rows, four numeric constants and one `SUM` formula per row, 100,000
cells and 20,000 formulas total. Each stage was timed separately with
`time.perf_counter`, with three repetitions per version. All measurements
below are milliseconds; SD is the sample standard deviation, not a
confidence interval.

| Stage | Before median | Before min–max | Before SD | After median | After min–max | After SD |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Error-overlay extraction | 21.31 | 18.91–21.78 | 1.54 | 5.76 | 5.72–8.56 | 1.63 |
| Cache loading | 82.00 | 77.00–91.62 | 7.43 | 67.23 | 65.52–76.90 | 6.13 |
| Complexity scan | 79.11 | 78.51–80.15 | 0.83 | 78.49 | 74.93–79.80 | 2.52 |
| Native import | 489.64 | 477.06–492.94 | 8.38 | 475.11 | 473.66–484.32 | 5.78 |

The host uses an Intel Core i7-13700K, Windows 11 Home x86-64 build 26200,
and 31.69 GiB of OS-reported physical memory. Hardware identification was
recorded later on the same host, not as a contemporaneous machine-load log.
The audit environment used CPython 3.13.12, formualizer 0.9.3 and openpyxl
3.1.5. Memory consumption was **not measured** for this before/after loader
benchmark.

These three-repeat timings show a local reduction in cache-loading work;
they are not an end-to-end speed guarantee or a statistically established
global improvement. Native import was not optimized by this change, so its
observed movement is not attributed to the loader fix. Workbook structure
and machine load affect results. Generated inputs and raw profiles remain
local artifacts.

## Private corpus stability smoke

A separate sequential smoke run analyzed eight existing workbooks through
the public isolated API, with a 120-second and 2,048 MiB limit per workbook.
The sample included small, medium, large, style-heavy and financial files.
All eight workers completed without timeout, crash or memory-limit failure;
the original files' before/after hashes were identical. Five workbooks had
no cell formulas, verified against source XML, so those cases exercise
loading rather than formula-heavy recalculation. The other three contained
1,965 cell formulas, all counted in extraction.

Elapsed times ranged from 1.70 to 7.67 seconds. Peak committed memory of the
Windows worker Job ranged from 43.47 to 522.64 MiB. This is not RSS and
excludes the calling process and Ollama. A separate Qwen validation was
running concurrently: this is evidence of completion under the stated
bounds, **not a comparative performance benchmark**. These memory readings
do not fill the missing memory measurement in the loader benchmark.

The resulting graphs contained 453 nodes with semantic-risk annotations
and 76 nodes with divergent cache verdicts. These counts may overlap and
represent grouped graph nodes, not numbers of incorrect cells. Stored
caches are not an independent correctness oracle; completion does not
certify the calculations. External-cache provenance also remains distinct
from a fresh reading of linked files. Workbook identities, raw reports and
all source data remain private local artifacts.

The official [PyPI release history](https://pypi.org/project/formualizer/#history)
was checked on 2026-09-19: 0.9.3, published September 11, remained the latest
release. No dependency upgrade was made on the assumption of an unreleased
fix.
