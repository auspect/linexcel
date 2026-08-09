# Checked-in workbook fixtures

`openpyxl` writes formulas, sheets and comments, so every other fixture in this
repository is generated at test time by `validation_workbooks.py`. It cannot
author a `vbaProject.bin` or a Power Query mashup — both are produced by Excel
itself — so those two live here as files, and `.gitignore` exempts this
directory from the blanket Excel-file ignore.

They are committed deliberately: the code that reads them is the boundary
between linexcel and a third-party parser (`oletools` for VBA), and a stub
cannot prove that boundary works on a real file.

## `power_query.xlsx`

Built once with Excel COM. Two queries, one loaded onto a sheet:

```powershell
$excel = New-Object -ComObject Excel.Application
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1); $ws.Name = "Source"
# A1:B4 filled with Product/Qty, then made a table:
$lo = $ws.ListObjects.Add(1, $ws.Range("A1:B4"), $null, 1); $lo.Name = "SalesTable"
$wb.Queries.Add("BusyProducts", $m)          # M source, see below
$wb.Queries.Add("TinyConnectionOnly", 'let Source = #table({"K"},{{1}}) in Source')
# loaded onto a second sheet through the Mashup OLE DB provider:
$conn = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=BusyProducts'
$lo2 = $dest.ListObjects.Add(0, $conn, $false, 1, $dest.Range("A1"))
$wb.SaveAs($path, 51)
```

The M source of `BusyProducts`, which reads a table back out of the workbook:

```m
let
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    Busy = Table.SelectRows(Source, each [Qty] > 1),
    Sorted = Table.Sort(Busy, {{"Qty", Order.Descending}})
in
    Sorted
```

Where it ends up in the file: `customXml/item1.xml` holds a UTF-16 XML document
whose `<DataMashup>` element is base64; decoded, that is a `uint32` version, a
`uint32` length, then a ZIP whose `Formulas/Section1.m` is the M source of every
query in plain text. `xl/connections.xml` names the query in a
`Microsoft.Mashup.OleDb.1` connection, and `xl/queryTables/queryTable1.xml` ties
that connection to the `ListObject` on the sheet.

## `macros.xlsm`

Built by hand in Excel, because writing to the VBA project through COM requires
turning on *Trust access to the VBA project object model*, and a fixture is not
worth changing a Trust Center setting for.

1. New workbook, three sheets named `Sales`, `Synthesis`, `Params`.
2. `Params!A1` = `0.2`.
3. Alt+F11 → Insert → Module (leave it named `Module1`), paste:

   ```vb
   Public Sub Refresh()
       Worksheets("Synthesis").Range("B10").Value = Rate()
       Cells(3, 2) = "ok"
   End Sub

   Private Function Rate() As Double
       Rate = Sheets("Params").Range("A1").Value
   End Function
   ```

4. Save as `tests/fixtures/macros.xlsm` (Excel Macro-Enabled Workbook).

`TestVbaOnARealWorkbook` skips itself while the file is absent, so the suite
stays green until it is added.
