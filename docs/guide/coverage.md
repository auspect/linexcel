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

## Not in the graph

### Power Query (Get & Transform)

**Queries are not part of the lineage.** A workbook whose data arrives through
Power Query shows the range the query loaded into, and nothing about where that
data came from — not the query, not its M source, not the table or file it
reads. On such a workbook the graph is silently incomplete, which is the one
failure mode this page exists to warn about.

Everything needed is in the file — the `customXml` part whose schema is
`http://schemas.microsoft.com/DataMashup` carries the M source, while
`xl/connections.xml` names each query connection and `xl/queryTables/*.xml`
ties it to the destination `ListObject` and sheet range — so this is a gap to
be closed, not a limitation of the format. Tracked in
[#34](https://github.com/auspect/linexcel/issues/34) for a release after 1.0.

### Formulas the engine does not implement

The range intersection operator — `=SUM(D2:D10 D5:D20)` — is the one met in
practice. Such a cell keeps the value stored in the file, says so in its card,
and is named in the workbook warnings rather than being shown as a
recalculation that never happened.

### Volatile functions

`TODAY()`, `NOW()`, `RAND()` and `RANDBETWEEN()` recompute to something other
than what the file stores, by definition. The value card reports the difference
like any other; it is not a defect in either reading.

### VBA the parser cannot follow

Range access is detected only where it is written literally. A reference built
at run time — `Range(addr)`, `Cells(i, j)` in a loop, `Application.Evaluate` —
is invisible to a static reader, and no static reader can do better. The
procedure still appears as a node, with its source.
