# Power Query audit

Reviewed on 2026-10-07 against the working tree.

linexcel extracts Power Query source and some dependencies from `.xlsx` and
`.xlsm` packages. It does not execute M, refresh queries, validate their output,
or establish column-level lineage. The checked-in Excel fixture confirms the
supported extraction path. This audit corrects several scanning errors;
language scope and destination resolution still need work.

## What the reader verifies

`src/linexcel/powerquery.py` reads DataMashup custom XML,
decodes its binary stream, and opens `Formulas/Section1.m` in the embedded ZIP.
The stream begins with a version and package length. Microsoft specifies version
zero and an OPC package containing the M part. The reader omits the remaining
permissions and metadata fields. This supports source extraction without
certifying the whole stream.
[MS-QDEFF binary stream](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/22557f6d-7c29-4554-8fe4-7b7a54ac7a2b),
[package parts](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/a4c2d0b9-9a9d-452d-8802-d68339374d57).

`read_destinations()` follows worksheet relationships through tables, query
tables and connections. `GraphBuilder.build_queries()` in `src/linexcel/graph.py`
connects recognized workbook tables or names to queries, upstream queries to
downstream queries, and queries to their worksheet landing ranges. External
sources appear as opaque inputs. Stored cell values describe the last saved
workbook state.

The fixture description in `tests/fixtures/README.md` records its Excel COM
provenance. `BusyProducts` reads `SalesTable` and loads to `Loaded!A1:B3`;
`TinyConnectionOnly` has no sheet load. This validates actual package plumbing,
but represents one workbook shape.

## Corrections included

The changes in `tests/test_powerquery.py` cover:

| Problem | Corrected behavior |
| --- | --- |
| Query references and bindings were case-folded. | `Sales` and `sales` resolve separately, including graph edges. M names are case-sensitive. [M reference](https://learn.microsoft.com/en-us/powerquery-m/). |
| Connector regexes scanned comments, strings and quoted names. | Only functions appearing in code produce sources; literal URL contents survive. |
| A comment before `shared` hid the member. | Comments act as whitespace in declarations and calls; returned bodies retain their original text. |
| Line comments only recognized LF. | All M newline characters end comments. |
| Targets and quoted names retained M escape syntax. | Character escapes, combined escapes and doubled quotes decode once; Windows backslashes stay literal. |
| `File.Contents("prefix" & parameter)` reported `prefix` as a path. | Concatenated targets remain unresolved. |
| `MyExcel.CurrentWorkbook()` matched the built-in suffix. | Workbook sources require the exact function name. |
| `Web.Page("<html>…</html>")` produced an external URL input. | Inline HTML produces no external source; nested `Web.Contents` still supplies its URL. [Web.Page](https://learn.microsoft.com/en-us/powerquery-m/web-page). |
| Only compressed mashup length had a ceiling. | Uncompressed M-part size and bounded reads also obey the 32 MiB ceiling. Unknown stream versions return no section. |
| Missing worksheet loads were called “loaded nowhere (connection only).” | Warnings report no detected worksheet destination, leaving data-model loads and unresolved relationships possible. |

Comment, newline and escape handling follows the
[M lexical specification](https://learn.microsoft.com/en-us/powerquery-m/m-spec-lexical-structure).
The scanner remains a heuristic.

## Remaining gaps, in priority order

### 1. Resolve scope and distinguish field selectors

`_referenced_queries()` excludes every identifier followed by `=` anywhere in
the body. Record fields or inner bindings can hide genuine references outside
their scope. Function parameters are not registered as bindings. Field selectors
can become false query dependencies.

These cases were reproduced against the corrected reader, with `Sales` in the
set of other query names:

| M expression | Current result | Required static result |
| --- | --- | --- |
| `let r = [Sales = 1] in Sales` | No edge | Edge from query `Sales`. |
| `(Sales) => Sales` | Edge from query `Sales` | No edge: function parameter. |
| `let r = [X = 1] in r[Sales]` | Edge from query `Sales` | No edge: field selector, even though the field is missing. |
| `let a = Sales, nested = let Sales = 1 in Sales in a` | No edge | Edge from query `Sales` used by `a`. |

M assigns environments to nested expressions. Inner variables override names
within their own environments.
[M environments](https://learn.microsoft.com/en-us/powerquery-m/m-spec-basic-concepts).

Use a syntax tree with scope resolution rather than extending the binding
regex. Dotted or Unicode identifiers, section-qualified references, nested
member metadata and dynamic `#shared` access also need coverage. Self references
are excluded by `names - {name}`, so their cycles cannot appear.

### 2. Report extraction failures and load uncertainty

`read_queries()` returns `[]` when Power Query is absent and when extraction
fails. Add `absent`, `read`, `unsupported` and `invalid` statuses with
diagnostics. Missing mashup nodes should count as incomplete coverage. Keep the
list API as a compatibility wrapper if needed.

Destination keys are case-folded despite case-sensitive query names. Prefer
exact matches, with an unambiguous fallback. `_connection_queries()` can fall
back to a `SELECT … FROM [name]` command for a non-Mashup connection; require the
actual Mashup provider before accepting it. Parse quoted connection-string
values so names containing semicolons survive. Ignore external relationships
explicitly.

Distinguish worksheet, data-model, connection-only and unknown load states when
metadata permits it. `Query.loaded` currently means a detected worksheet
destination; it cannot establish the other states.

### 3. Bound each package and XML stage

The new nested M-part limit closes one expansion path. `read_section()` still
reads custom XML fully before parsing. `_unpack_mashup()` decodes the whole
base64 string before checking the package length. `read_destinations()` reads
XML without per-part budgets. Add compressed and uncompressed limits, total
expansion and part-count budgets, and bounded reads at those stages.

The default analysis worker enforces `ExecutionPolicy` time and memory limits.
Direct `read_queries()` calls and `isolated=False` do not inherit that boundary.
The reader performs no filesystem extraction, network requests or M execution;
workbook paths and URLs do not trigger refreshes.

Consider opt-in redaction for exported source and AI documentation. M can embed
credentials in SQL connection strings, URLs and header records. Redaction must
cover source code, dependency labels and warnings.

### 4. Describe connector identity and uncertainty

The database suffix rule can classify a custom transform such as
`Acme.Tables("text")` as a source. It misses other suffixes, including
`Salesforce.Data`. Prefer an explicit registry that distinguishes documented
connectors from guesses and unresolved calls.

`Sql.Database("server", "database")` stores only `server`; two databases on it
collapse into one source. Preserve both arguments in a structured identity.
[Sql.Database arguments](https://learn.microsoft.com/en-us/powerquery-m/sql-database).

`Web.Contents` stores the literal base URL; `RelativePath` and `Query` options
can change the request. Label it as a base URL. Resolve literal options only
where their semantics are proven.
[Web.Contents options](https://learn.microsoft.com/en-us/powerquery-m/web-contents).

Limited constant propagation could resolve aliases and concatenated paths.
Keep execution uncertainty visible: branches and lazy bindings may name
connectors that never run.

## Packages and performance proposals

Use Microsoft's MIT-licensed
[`@microsoft/powerquery-parser`](https://github.com/microsoft/powerquery-parser)
as an optional prototype or test oracle. Its `tryLexParse` supplies lexical and
syntax analysis, without M execution. Pair it with
[`powerquery-language-services`](https://github.com/microsoft/powerquery-language-services).
Its [scope implementation](https://github.com/microsoft/powerquery-language-services/blob/master/src/powerquery-language-services/inspection/scope/scopeInspection.ts)
handles function parameters, `let`, records, sections and field selectors. This
addresses the reproduced failures. A Node runtime, startup cost, version
management and Python interoperability add integration work. Benchmark a
worker or build-time oracle before adding runtime dependencies.

For Python extraction, share one workbook ZIP handle between section and
destination reads, cache relationship XML, and reuse lexical tokens across
parsing and scanning. The current reader traverses query bodies repeatedly.
Benchmark extraction separately from Excel calculation with many queries,
long M documents and malformed nested packages. No measured speed improvement
is claimed by this audit.

For optimization advice on workbook M, suggest early filters, appropriate
connectors and delayed expensive operations such as sorting. Microsoft
recommends these patterns. linexcel cannot establish query folding or refresh
time from source alone; validate improvements with refresh measurements in
Excel or Power Query.
[Power Query best practices](https://learn.microsoft.com/en-us/power-query/best-practices).

## Validation

`pytest tests/test_powerquery.py -q` passes all 61 tests. The broader run,
`pytest tests/test_powerquery.py tests/test_lineage.py tests/test_targeted_boundaries.py -q`,
passes all 262 tests. Ruff formatting and linting pass for both changed Python
files; `ty check src/linexcel/powerquery.py` passes. Tests retain the Excel-generated
fixture and local-binding coverage, and add case-sensitive graph edges,
comment and literal checks, escapes, dynamic-target rejection, stream versions
and nested decompression bounds.

These unit tests do not satisfy the full manual acceptance procedure in
`AGENTS.md`, which is handled separately.
