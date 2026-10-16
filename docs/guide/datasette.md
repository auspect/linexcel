# Browse source data with Datasette

Start the local application and import a workbook:

```bash
linexcel serve --port 8765
```

Open **Données** to browse the workbook in the embedded Datasette interface.
Choose a worksheet, sort a column, add a filter or facet, and move through
the result pages. The interface also provides JSON and CSV exports of the
selected data. Datasette is included in the normal installation; it does not
require an AI provider or a second server process.

Each worksheet normally has a grid table named `sheet_001`, `sheet_002`, and
so on, in workbook order. Column names are Excel letters rather than guessed
headers. The `row` key is the original Excel row number, so gaps remain gaps.
Only rows containing imported source cells with provenance from this workbook
are present. Blank dependency placeholders synthesized for the graph are excluded.

## What the values mean

The grid shows decoded **stored inputs and saved formula caches** from the
primary workbook. Browsing does not recalculate formulas or execute VBA, and
it does not replace the grid values with subsequent engine results. A formula
with no saved value shows `NULL`; it is not assigned zero. Uploaded reference
workbooks contribute to the project revision but do not become additional
Datasette worksheet tables.

The `cells` table provides a row per imported source worksheet cell with its address,
Excel number format, value status, exact JSON value and type, original formula,
canonical formula, and source provenance. In particular:

| Field | Meaning |
| --- | --- |
| `value`, `value_json`, `value_kind` | Browsable SQLite value and the exact imported JSON value/type. SQLite represents booleans as `0`/`1`; large integers outside SQLite's range become text in the grid. |
| `value_role`, `value_status` | Stored input versus saved formula cache, and whether that source value is present, missing or unreadable. |
| `number_format` | Source Excel format code. The grid does not reproduce Excel's formatted display. |
| `source_formula` | Stored formula text with a presentation `=` prefix, when that text exists. |
| `formula_xml_text` | Exact decoded character content of the source `<f>` element, without adding `=`. A shared formula follower can have empty source text. |
| `canonical_formula` | Formula used by the structural analyzer; it can be translated from a shared master or normalized. |
| `formula_provenance`, `value_provenance` | Source workbook SHA-256, worksheet XML part, cell address and extraction status. |

Coverage follows the complete disk index rather than the viewer's first page
or its sample of nodes. Formatting-only blanks and synthetic blank dependency
placeholders are excluded. A formula without a saved cache remains a source
cell, with its missing cache explicitly recorded. Import
diagnostics and partial import status remain applicable to the data view. An
empty sheet has an empty grid. Sheets wider than 512 columns use a sheet-filtered
view of the complete `cells` table instead of truncating columns. `.xlam` and
legacy `.xla` imports expose VBA sources and explicitly have no worksheet data
view, matching their macro-only analysis scope.

## Isolation and storage

Datasette runs in the existing application process. Every HTML page, static
asset, table detail and export request checks the owning project through the
same authentication used by `serve`. Mounted applications must configure
`mount_path`; generated Datasette links stay under that prefix. No external
Datasette HTTP service is contacted.

The derived `source-data.sqlite` snapshot is generated on first access inside
the project's private storage directory. Its key includes the project revision,
index generation and data schema version. Reimporting the source invalidates
the snapshot. Deleting a project removes the snapshot with its other files.
Snapshot publication is atomic and observes the configured storage budget;
host filesystem quotas are still needed for a hard disk ceiling.

Datasette uses immutable SQLite connections. Arbitrary SQL queries and database
downloads are disabled; native table filters, facets, sorting, pagination, row
details, JSON and CSV exports remain available. Requests to introspection,
actor management and other databases are outside this mounted surface.

## HTTP contract

`GET /api/projects/{project_id}/data` returns:

```json
{
  "available": true,
  "revision": "project revision hash",
  "sourceSha256": "primary workbook hash",
  "url": "/api/projects/{project_id}/datasette/source",
  "cellsUrl": "/api/projects/{project_id}/datasette/source/cells",
  "sourceStatus": "complete",
  "diagnostics": [],
  "sheets": [
    {
      "sheet": "Sheet1",
      "table": "sheet_001",
      "url": "/api/projects/{project_id}/datasette/source/sheet_001",
      "rowCount": 2,
      "cellCount": 4,
      "columns": ["A", "B"],
      "layout": "grid"
    }
  ]
}
```

The response also includes `semantics` describing these bounds. URLs include
the configured mount prefix. For macro-only imports, the response is
`{"available": false, "reason": "macro_only", "sheets": [], "revision": "…"}`.
A source change on disk returns HTTP 409 instead of serving an obsolete view;
an insufficient storage budget returns HTTP 413. Datasette routes accept only
GET and HEAD.

The bridge uses Datasette's documented [internal client](https://docs.datasette.io/en/stable/internals.html#datasette-client)
and [base URL setting](https://docs.datasette.io/en/stable/settings.html#base-url).
The local URL adapter prevents duplicate prefixes in links produced by
Datasette 0.65; source cell contents are not rewritten.
