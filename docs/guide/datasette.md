# Browse source data with Datasette

Start the local application and import a workbook:

```bash
uvx linexcel serve --open --port 8765
```

Open **Données / Data** and choose a worksheet. **Excel presentation** shows
a bounded area with saved values, direct styles and worksheet geometry.
Select **Datasette filters and exports** to sort columns, add filters or facets,
and move through the result pages. The interface also provides JSON and CSV exports of the
selected data. Datasette is included in the normal installation; it does not
require an AI provider or a second server process.

Each worksheet normally has a grid table named `sheet_001`, `sheet_002`, and
so on, in workbook order. Column names are Excel letters rather than guessed
headers. The `row` key is the original Excel row number, so gaps remain gaps.
Only rows containing imported source cells with provenance from this workbook
are present. Blank dependency placeholders synthesized for the graph are excluded.

## Excel presentation and its limits

The presentation adapter reads the original OOXML styles, theme/indexed colors,
fonts, solid fills, borders, alignment, row heights, column widths, hidden rows
and columns, and merged ranges. It keeps formatting-only cells in this preview;
the raw Datasette tables retain their source-cell semantics. Values and formulas
come from the imported index, bound to the workbook hash. No workbook is modified
and no formula or macro is executed by this view.

Common decimal, grouping, percentage and numeric date/time formats have an
explicit formatter status. A complex or unsupported format displays the raw
value and is counted in the notice below the grid. A missing formula cache is
shown as `∅`; selecting that cell explains the missing saved value. Empty strings
and actual blanks remain distinct in the cell detail. Select cells with a click
or use the arrow keys, Home/End and Ctrl+Home/End to read their raw value, number
format and source formula. Datasette's `cells` table retains the fuller source
provenance and exact imported JSON.

The initial area is 60 rows by 20 columns. **First row** and **First column**
accept original Excel coordinates; row navigation moves through subsequent
areas. The API caps requests at 100 rows and 40 columns and rejects oversized
XML parts, format codes longer than 1,024 characters and responses exceeding
2 MiB. A merged range crossing the visible area is
shown as separate cells with a notice; it does not borrow an unseen anchor's
value. Scrolling stays inside the grid on narrow screens.

This is an approximate browser presentation. Column widths, fonts and borders
can differ from Excel; custom indexed palettes and some implicit cell alignments
are not reproduced. Conditional formatting, drawings, charts, text rotation,
print layout, number-format color directives and complex number formats are not reproduced. Frozen-pane metadata
is extracted but the preview pins only its own row/column headers. Hidden sheets
remain available for analysis. Use **Captures** for a separate LibreOffice render,
which can itself recalculate formulas and differ from Excel's saved values.

The source format has distinct rules for [cell styles, row/column styles and
merged cells](https://openpyxl.readthedocs.io/en/stable/styles.html). Datasette's
[render-cell hook](https://docs.datasette.io/en/0.65.5/plugin_hooks.html#render-cell)
customizes scalar display; preserving worksheet coordinates and merges while
filtering/reordering rows requires a separate presentation. The application keeps
both views available through the same authenticated project.

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
  "presentationAvailable": true,
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

`GET /api/projects/{project_id}/data/presentation?sheet=Sheet1&row=1&column=1&rows=60&columns=20`
returns `cells`, a reusable `styles` map, `rowDimensions`, `columnDimensions`,
`merges`, `extent`, `defaults`, `freezePane`, the actual `window`, explicit `limits`
and `notices`, with the same workbook hash and project revision. Each cell has
`display`, `formatStatus`, `styleId`, `numberFormat`, its raw `value` and source
`formula`, with `formulaRole` distinguishing original source text from a
canonical fallback; `sourceFormula` and `canonicalFormula` remain separate.
This endpoint has the same owner check, stale-source rejection and
read-only policy as the data browser. Oversized presentation sources return 413;
Datasette remains available independently.

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
