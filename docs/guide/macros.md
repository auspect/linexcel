# Static VBA and add-ins

Read VBA from `.xlsm` workbooks, `.xlam` add-ins and legacy OLE `.xla` files
without opening Excel or executing their code:

```python
from linexcel import read_macros

project = read_macros("tools.xlam")
print(project["sha256"], project["warnings"])
for procedure in project["procedures"]:
    print(procedure["module"], procedure["name"], procedure["lines"])
```

```powershell
uv run linexcel macros tools.xla -o validation_screenshots/tools-source.json
uv run linexcel analyze tools.xla -o validation_screenshots/tools.html
```

`read_macros` retains complete decoded module/procedure text. The graph's source
cards show at most 6,000 characters per procedure and explicitly identify
truncation. File SHA-256, decoded module SHA-256, procedure SHA-256, stream path
and line interval have distinct meanings: decoded source is not the original
compressed binary stream. Read the saved warnings before interpreting absence
of source; damaged or partly readable projects can yield partial extraction.

Normal `.xlsx`/`.xlsm` analysis preserves worksheet lineage and embedded VBA.
`.xlam`/`.xla` analysis produces a **macro source report**: worksheet formulas, values, cache,
context, screenshots and UDF calculation are outside its scope. There are no
worksheet targets in this mode. `serve` accepts both add-in formats as primary
files or references and displays static source cards. `--refs-dir DIR` includes
the supported macro files from that folder.

Calls and literal range access are heuristic source observations. They do not
establish that a branch executed or that `ActiveWorkbook`, aliases or dynamic
references resolve to a particular workbook. The lazy viewer keeps ambiguous
literal references unresolved. This reader does not certify compiled p-code,
Excel 4.0 macros, add-in installation or runtime behavior.

In a normal workbook graph, only explicit `ThisWorkbook` can identify a range
in the source workbook. Bare sheet access, `ActiveWorkbook`, workbook aliases
and an add-in's accesses remain unresolved, even when sheet names match.
[Unqualified `Worksheets` uses the active workbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.worksheets),
while [an add-in's `ThisWorkbook` identifies the add-in itself](https://learn.microsoft.com/en-us/office/vba/api/excel.application.thisworkbook).
The recorded `referenceSemantics` and `static_vba_heuristic` edge scope describe
this conservative source interpretation; all static VBA edges are approximate,
without establishing runtime binding. Source scans ignore reference-like text
inside string literals and distinguish a conditional comparison from assignment.
Colliding module streams retain all stream origins and concatenation boundaries
with a warning; procedure lines then refer to the decoded concatenated text.

The reader uses the existing
[oletools VBA extractor](https://github.com/decalage2/oletools/blob/master/oletools/doc/olevba.md).
xlwings remains an optional separate Excel integration:
[its open-source mode requires Excel](https://docs.xlwings.org/en/stable/installation.html),
and [`Book.macro()` executes a procedure](https://docs.xlwings.org/en/stable/api/book.html#xlwings.Book.macro).
Neither is needed for static source extraction. Rendering uses a disposable
LibreOffice profile with very-high macro security and no trusted locations;
[only trusted locations can execute macros at that level](https://help.libreoffice.org/latest/en-US/text/shared/optionen/macrosecurity_sl.html).
