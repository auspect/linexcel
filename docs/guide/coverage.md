# What is in the lineage, and what is not

A lineage tool is only useful if you know where it stops. This page lists what
linexcel reads out of a workbook, and — more importantly — what it does not, so
that a graph which looks complete can be trusted to be complete.

## In the graph

| | |
| --- | --- |
| Formulas | Every formula cell, with its computed value and a step-by-step decomposition. Cells sharing an R1C1 pattern are grouped into one node |
| Cells and ranges | Precedent and dependent edges, including across sheets |
| Defined names | Workbook-scoped and sheet-scoped, resolved to the cells they point at |
| VBA | Procedures as nodes, their internal call graph, and the ranges they read and write when the reference is written literally — `Range("A1")`, `Cells(2, 3)`, `[A1:B4]`, `Worksheets("X").Range(...)` |
| Power Query | Each query as a node with its M source, the range it loads into, the sources it names and the queries it chains from |
| Values | Read back from the workbook and recomputed, with both readings shown side by side |

## Represented, but not resolved

These appear in the graph as grey **external reference** nodes. The dependency
is visible; what sits at the other end of it is not, because it is not in the
file:

- links to another workbook — `'[1]Annual'!B4`
- 3-D references spanning sheets — `Sheet1:Sheet3!A1`
- structured table references the engine cannot resolve
- a defined name pointing outside the workbook

Showing them as nodes rather than dropping them is deliberate: a formula whose
precedent is missing is a fact about the workbook, not a gap to hide.

## Power Query (Get & Transform)

Queries are in the lineage. Each one is a node carrying its M source, and the
graph crosses it: `Source!A1:B4` → `BusyProducts` → `Loaded!A1:B3`, where the
landing range used to sit at the top of the graph with nothing above it.

What is read, and from where:

| | |
| --- | --- |
| The M source | the `customXml` part whose schema is `http://schemas.microsoft.com/DataMashup`: base64, then a ZIP whose `Formulas/Section1.m` holds every query in plain text |
| The destination | `xl/connections.xml` names the query a connection loads, `xl/queryTables/*.xml` ties that connection to a table on a sheet |
| The sources | read off the M — `Excel.CurrentWorkbook(){[Name="X"]}` is a table or defined name of this file and is linked to it, and a query reading another is an edge between the two |
| Everything else | `File.Contents`, `Folder.Files`, `Web.Contents`, `Sql.Database` and any `*.Database`, `*.Feed` or `*.DataSource` connector: named, never opened |

A query that computes without loading anywhere — connection only, or straight
into the data model — is shown as loading nowhere rather than dropped.

Where it stops is the data behind an outside source. A CSV on a share, a REST
endpoint, a database: the query names them, linexcel does not read them, and
the panel says so under the source list. And M is a real language — a source
built at run time rather than written out is invisible to a static reader,
which is the same limit VBA has below.

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
linexcel does not recompute them. The card shows the value the file stores and
says *Not recalculated (volatile)* in the other column, and the cell is not
decomposed step by step — every step would carry a figure from today's clock.

`OFFSET`, `INDIRECT`, `CELL` and `INFO` are volatile to Excel in that they
recalculate on every edit, but they return the same value for the same
workbook, so those are still computed.

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
