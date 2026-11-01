# What is in the lineage, and what is not

A lineage graph describes extracted relationships. Extraction, reference
resolution, calculation and agreement with Excel are separate capabilities.
Read the calculation status and declared omissions alongside the graph.

## Indexed application and batch analysis

`linexcel serve` builds a disk-backed structural index and calculates selected
dependency closures. `linexcel analyze` and `LineageResult` use the batch path.
The two paths have different loading, calculation and volatile-value policies.
The following table describes the indexed application's tested subset; it does
not extend historical batch measurements or establish full Excel equivalence.

| Feature | Indexed calculation | Remaining limits |
| --- | --- | --- |
| Tables | Source-declared headers, data, totals and row references resolve to exact rectangles | Empty data regions, disjoint selections, invalid metadata |
| Defined names | Scoped constants, absolute references and bounded static computed definitions | Relative definitions without a known anchor, dynamic/cyclic definitions; expansion is bounded |
| 3D references | Exact source sheet order in supported direct aggregate arguments | Computed arguments and other contexts |
| Supplied external files | Sparse transitive formulas, ranges, names and tables without replacing formulas by caches | Missing files, mixed-calendar formulas/dates, external arrays and dynamic contexts |
| `SUBTOTAL` / `AGGREGATE` | Source hidden rows, nested totals and typed reference aggregates; up to 20,000 positions | Active filters with modes including manually hidden rows, computed references/options, external visibility, unverified empty MEDIAN and fractional ranks |
| Arrays and spills | Declared source regions, anchor calculation and observed extent verification, including unsaved children | Undeclared/changing extents, data tables, external or cyclic regions |
| `INDIRECT` / `OFFSET` | Exact references from bounded source scalar arithmetic/concatenation and contiguous offsets, including A1/R1C1 | Volatile/runtime address expressions, dynamic names, external contexts and fractional offsets |
| Volatile functions | Recorded UTC operation clock and random seed for replay | Random sequence parity with Excel is not established |
| Iterative cycles | Source iteration settings and explicit native convergence telemetry | Disabled/malformed settings, nonconvergence and incompatible external settings |

Batch calculation also resolves declared table references and applies the bounded
visibility adapter and supported 3D aggregate expansion. Its graph still keeps
structured references as opaque nodes; previews needing `#This Row` context can
remain unavailable. `FORMULATEXT` retains the saved reading with an explicit
limitation, because engine-normalized syntax does not establish Excel's original
formula text. External ranges retain their types and dimensions, with external
reading provenance carried through their dependent calculations; this batch
path reads supplied files' saved values rather than recalculating their closure.

Excel-saved array children are removed from the engine's
private input copy so their anchor can calculate a fresh spill; the original
workbook and saved values remain available for provenance. Changed spill extents
and unsupported visibility readings invalidate dependent calculations, including
error-handler fallbacks. Active filters still prevent include-hidden modes from
distinguishing filtered rows from rows hidden manually.

The versioned `capabilities` metadata lists extraction, resolution, calculation,
verification and limitations independently. Unsupported prerequisites produce an
explicit result rather than a saved-cache substitution. A successful engine
result remains an engine observation; native Excel comparison requires an
actual Windows Excel oracle receipt for the same source and context.

An observed engine limitation in a dependency invalidates the targeted result,
including when an error handler would otherwise hide it. It is distinct from
an Excel error value such as `#N/A`, which can legitimately participate in error
handling. A dependency value that the engine cannot expose remains separately
labelled unavailable; the absence of a readable step does not identify a cause.

VBA and Power Query remain static source analysis in both paths. Neither is
executed or refreshed by linexcel.

## In the graph

| | |
| --- | --- |
| Formulas | Extracted formula source, calculation status and available decomposition. Batch analysis can group cells sharing an R1C1 pattern; the index retains individual cells with separate pattern metadata |
| Cells and ranges | Precedent and dependent edges, including across sheets |
| Defined names | Workbook-scoped and sheet-scoped, resolved to the cells they point at |
| VBA | Procedures as nodes, their internal call graph, and the ranges they read and write when the reference is written literally — `Range("A1")`, `Cells(2, 3)`, `[A1:B4]`, `Worksheets("X").Range(...)` |
| Power Query | Each query as a node with its M source, the range it loads into, the sources it names and the queries it chains from |
| Values | Source readings and available engine results, shown separately with their provenance |

## Represented, but not resolved

Unresolved dependencies remain explicit graph nodes and diagnostics. Common
examples include:

- links to a workbook that was not supplied — `'[1]Annual'!B4`
- 3-D references outside the supported aggregate contexts — `Sheet1:Sheet3!A1`
- structured table references outside the supported source metadata
- a defined name pointing outside the workbook

Showing them as nodes rather than dropping them is deliberate: a formula whose
precedent is missing is a fact about the workbook, not a gap to hide.

## Power Query (Get & Transform)

Power Query extraction is static and covers a subset of M. linexcel does not
execute M, refresh queries or verify their saved results. Recognized queries
are in the lineage. Each one is a node carrying its M source, and the
graph crosses it: `Source!A1:B4` → `BusyProducts` → `Loaded!A1:B3`, where the
landing range used to sit at the top of the graph with nothing above it.

What is read, and from where:

| | |
| --- | --- |
| The M source | the `customXml` part whose schema is `http://schemas.microsoft.com/DataMashup`: base64, then the versioned MS-QDEFF header and package ZIP containing `Formulas/Section1.m` |
| The destination | `xl/connections.xml` names the query a connection loads, `xl/queryTables/*.xml` ties that connection to a table on a sheet |
| The sources | read off the M — `Excel.CurrentWorkbook(){[Name="X"]}` is a table or defined name of this file and is linked to it, and a query reading another is an edge between the two |
| External connectors | Literal arguments to recognized calls such as `File.Contents`, `Folder.Files`, `Web.Contents`, `Sql.Database` and `*.Database`, `*.Feed` or `*.DataSource`: named, never opened |

A query that computes without loading anywhere — connection only, or straight
into the data model — is shown as loading nowhere rather than dropped.

Where it stops is the data behind an outside source. A CSV on a share, a REST
endpoint, a database: the query names them, linexcel does not read them, and
the panel says so under the source list. And M is a real language — a source
built at run time rather than written out is invisible to a static reader,
which is the same limit VBA has below. Comments and text literals are excluded
from dependency scanning, and query identifiers retain M's case sensitivity.
Dynamic connector arguments and complex scope rules can still produce missing
or ambiguous edges. See the [Power Query audit](../powerquery_audit.md) for the
tested cases and proposed parser improvements.

## Not in the graph

### Formulas the engine does not implement

The range intersection operator — `=SUM(D2:D10 D5:D20)` — is the one met in
practice. Such a cell keeps the value stored in the file, says so in its card,
and is named in the workbook warnings rather than being shown as a
recalculation that never happened.

### Volatile functions

`TODAY()`, `NOW()`, `RAND()`, `RANDBETWEEN()` and `RANDARRAY()` answer
differently every time they are computed, so there is nothing to check them
against: a value recomputed today can never agree with a file saved last week.
The batch analysis policy preserves them. The card shows the value the file stores and
says *Not recalculated (volatile)* in the other column, and the cell is not
decomposed step by step — every step would carry a figure from today's clock.

`OFFSET`, `INDIRECT`, `CELL` and `INFO` are volatile to Excel in that they
recalculate on every edit, but they return the same value for the same
workbook, so those are still computed.

In the indexed application, an explicit targeted calculation instead records
its operation clock and seed. Its value is a fresh engine reading, with a
separate comparison to the saved cache. Dynamic and metadata-sensitive
functions follow the bounded support contract above; batch behavior must not
be assumed for this path.

### Other workbooks

A formula reading `'[Budget FY26.xlsx]Annual'!B4` depends on a file that is not
being analyzed, and the calculation engine has nothing to follow. Three answers,
in decreasing order of certainty:

| | |
| --- | --- |
| **named** | always — `xl/externalLinks` carries the file name and the path the workbook declares, and the node is labelled with them |
| **cached** | when Excel saved the values it last read across the link, they are used and labelled as coming from that cache |
| **read** | with `--refs-dir` (`refs_dir=`), the referenced workbook is opened from that folder and the reference evaluates to the value it stands for |

The panel lists every linked workbook of a cell with its path and which of the
three it got, so a stale or missing dependency is visible rather than silent.
The same folder is searched for the `.xlam`, `.xla` and `.xlsm` add-ins whose
VBA a workbook calls into; their procedures join the graph, each module tagged
with the file it came from.

### VBA the parser cannot follow

Range access is detected only where it is written literally. A reference built
at run time — `Range(addr)`, `Cells(i, j)` in a loop, `Application.Evaluate` —
is invisible to a static reader, and no static reader can do better. The
procedure still appears as a node, with its source.
