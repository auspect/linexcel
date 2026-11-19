# Read the lineage report

A user opens the single HTML file written by `linexcel analyze` straight from
disk, finds a cell, and reads its source formula, its value, the step-by-step
decomposition and what it depends on, with no server and no network.

## Sub-features

- `report-open` loads from a `file://` URL and shows the workbook name and
  totals.
- `report-search` finds a cell by reference and lists it as a result.
- `report-node` shows formula, values, decomposition and precedents of the
  selected node.
- `report-sheets` shows each worksheet with its metadata and a cell preview.
- `report-optional-tabs` hides the views that have no content.
- `report-offline` requests nothing outside the file.
- `report-mobile` keeps the same path usable at 390×844.

## How to get to it (user POV)

- Open the `*_lineage.html` file that `linexcel analyze` wrote, in a browser.
- Open a file written by `result.save_html(...)` from Python.
- Display `result` in a marimo or Jupyter cell (embedded copy of the report).
- Open an investigation exported from the local application.

## Driving it with ctl

Preconditions:

- `{evidence}/reports/sales.html` exists, written by the **Analyze** step of
  [Analyze a workbook](./analyze-workbook.md) in this run.

Run one session per viewport. The steps below are for
`ctl browser --run <run> --open "{evidence}/reports/sales.html" --name report-desktop --viewport desktop`
followed by these `--step` values in order. Repeat with `--name report-mobile
--viewport mobile` for `report-mobile`.

- **Open.** The page loads. Step `expect-text h1 :: fy26_commercial_review.xlsx`.
  The banner also reads `50 formulas · 20 nodes · 21 edges`.
- **Landing proof.** Steps `aria landing` and `shot landing`. The snapshot
  lists the tabs `Graph` and `Sheets` and the `Sheet filter` combobox.
- **Search.** Type a cell reference and submit. Steps
  `fill #lin-search :: Summary!C10` and `press #lin-search :: Enter`. A
  `Search results` group appears with one button starting `Summary!C10`.
  Step `expect-count #lin-result-list button :: 1`.
- **Select.** Choose the result. Step `click #lin-result-list button`. The
  `Node details` region shows the node.
- **Node panel.** Assert what the user reads. Steps
  `expect-text #lin-panel :: =ROUND(C8-C6*RebateRate-C8*Assumptions!$D$8, 2)`,
  `expect-text #lin-panel :: 816,668.1` and
  `expect-text #lin-panel :: Precedents (4)`. Then `text node :: #lin-panel`
  and `shot node`. The text lists `RebateRate`, `Summary!C8`, `Summary!C6` and
  `Assumptions!D8`, and the decomposition down to `C6 * RebateRate = 46,897.9`.
- **Sheets.** Switch view. Steps `click #lin-tab-sheets`,
  `expect-text #lin-sheets-main :: 18 rows × 7 columns`, `aria sheets :: #lin-sheets-main`.
  The sidebar has one button per sheet: Summary, Sales, Catalog, Assumptions.
- **Optional tabs.** Steps `expect-hidden #lin-tab-overview` and
  `expect-hidden #lin-tab-screenshots`. A report made without AI documentation
  and without `--screenshots` offers neither view.
- **Offline.** No step. The session exits 1 if the page requests any
  non-loopback address or raises a JavaScript error; `session.json` records
  `"externalRequests": []` and `"pageErrors": []`.

The step with `$D$8` contains `$`. In PowerShell wrap that step in single
quotes: `--step 'expect-text #lin-panel :: =ROUND(C8-C6*RebateRate-C8*Assumptions!$D$8, 2)'`.

## Gotchas

- The Overview and Visual preview tabs exist in the page but stay hidden
  unless the report carries a workbook overview (AI) or sheet screenshots
  (LibreOffice). Clicking them times out; that is not a defect of the tab.
- The search box searches the visible graph by default. A cell that is folded
  into a group of stretched formulas is found only with the `Scope` combobox
  set to `Workbook`.
- The report formats numbers for reading (`816,668.1`); the JSON holds the raw
  value (`816668.1`). Assert each in its own format.
- Section headings of the panel are upper-cased by the stylesheet. A `text`
  dump shows `PRECEDENTS (4)`, but `expect-text` matches the text of the page
  source, `Precedents (4)`.
- The graph layout is not deterministic. Never assert node positions or
  compare screenshots pixel by pixel.
- On mobile the node panel opens as a sheet over the lower half of the graph.
  A viewport `shot` shows only its top; use `text node :: #lin-panel` for the
  rest, and close the panel with its `✕` button before a step that needs the
  canvas.
- A notebook embeds the report in a sandboxed frame. That entry point is not
  driven here; say so if your change is about embedding.
