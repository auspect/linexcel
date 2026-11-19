# Follow a calculation

In the local application, a user selects a formula cell of an imported
workbook. The application shows its source formula and saved value at once,
starts a targeted calculation of that cell and what it depends on, and shows
the result beside the dependency cells.

## Sub-features

- `calc-search` finds a cell from the graph search box.
- `calc-auto` starts the targeted calculation on selection, without a click.
- `calc-result` shows the engine result next to the saved Excel value.
- `calc-dependencies` lists the dependency cells with their own values.
- `calc-task` records the calculation as a completed operation.
- `calc-from-nodes` reaches the same cell from the Nodes list.

## How to get to it (user POV)

- Open the `Graph` tab, type a cell reference in the search box, choose the
  result.
- Open the `Nodes` tab, search or filter by sheet, choose a node.
- Open the `Sheets` tab, choose `Explore this sheet`, then choose a node.
- Click a node on the graph canvas.
- Choose `Recalculate` on an already selected cell.

## Driving it with ctl

Preconditions:

- The workspace `Verify sales` exists and is the last one opened in the saved
  browser state: run the **Import** steps of
  [Import a workbook](./import-workbook.md) first, in the same run.

Run `ctl browser --run <run> --open app --name calc-graph --viewport desktop`
with these steps in order.

- **Open the graph.** Steps `expect-text #project-title :: Verify sales`,
  `click #tab-graph` and `expect-text #view-graph :: Workbook graph ·`. The
  last one waits for the workbook overview, which the application builds the
  first time the Graph tab is shown.
- **Search.** Step `fill #graph-search :: Summary!C10`. A result for the cell
  appears.
- **Select.** Step `click #graph-results [data-related='Summary!C10']`. The
  reader on the right shows `Selected cell · Summary!C10`.
- **Wait for the calculation.** Steps
  `wait #graph-calculation-steps .step-card` and `wait-tasks`. No click on
  `Recalculate` is needed.
- **Result.** Steps
  `expect-text #graph-preview :: Linexcel result · 816668.1`,
  `expect-text #graph-preview :: Saved cache: Not available` and
  `expect-text #graph-preview :: Comparison unknown`. The value equals the one
  `linexcel analyze` reports for the same cell.
- **Dependencies.** Steps
  `expect-text #graph-calculation-steps :: 100 / 226 dependencies shown` and
  `expect-text #graph-calculation-steps :: Sales!J10`, then
  `text calculation :: #graph-preview`, `aria reader :: role=region[name='Selected node reader']`
  and `shot calculation`.
- **Second view.** The operation is recorded. Steps `click #tab-tasks`,
  `expect-text #task-list :: Targeted recalculation`,
  `expect-text #task-list :: Summary!C10 · Completed`, `text tasks :: #task-list`.
- **Nodes entry.** Run a second session,
  `ctl browser --run <run> --open app --name calc-nodes`, with steps
  `click #tab-nodes`, `fill #search :: Summary!C10`,
  `expect-count #node-list [data-node] :: 1`,
  `click #node-list [data-node='Summary!C10']`, `wait #node-detail`,
  `expect-text #node-detail :: =ROUND(C8-C6*RebateRate-C8*Assumptions!`,
  `text detail :: #node-detail`, `shot node-detail`.
- **Mobile.** Repeat the graph session with `--name calc-graph-mobile
  --viewport mobile`. The reader is stacked under the graph, and its value
  chips are shortened with an ellipsis on screen; the same assertions on
  `#graph-preview` hold because they read the full text.

## Gotchas

- Until 2026-11-19, a cell selected in the same session as an import, while
  the workbook overview was still being built, lost its selection when the
  overview arrived: the view returned to the overview with an empty search
  box. Measured then: 5 of 5 sessions lost it once the session waited three
  seconds after the click. Fixed in `app.html`; regression test
  `tests/test_app_large_browser.py::test_selection_survives_delayed_workbook_overview_after_import`.
  The recipe still waits for `Workbook graph ·` so that it starts from a
  settled page. If your change touches selection, the overview or task
  polling, drive the immediate path on purpose: import, then search and select
  in the same session without that wait, add `wait-tasks` and `wait-ms 3000`,
  and assert `Selected cell · Summary!C10` again.
- A cell selected before the Graph tab is shown (from the Nodes tab, or in the
  import session) no longer starts the workbook overview on its own. Choose
  `Overview` to get it.
- The Nodes search filters after a short delay. Clicking the first node right
  after `fill` selects the first row of the unfiltered list (the defined name
  `RebateRate`). Wait with `expect-count … :: 1`, then click by `data-node`.
- Selecting an input cell or a defined name starts no calculation, and its
  panel never gets a `.step-card`. Use a formula cell.
- `Saved cache: Not available` is specific to the fixture, which openpyxl
  writes without cached values. A workbook saved by Excel shows a cache and a
  comparison. That path is not covered by this recipe.
- The dependency table shows at most 100 rows of the 226. Assert the counter,
  not a row beyond the first hundred.
- The values in the table are cell values after the targeted calculation. The
  application says so itself: they are not a trace of evaluated branches.
- `wait-tasks` reads the server. The page catches up on its next poll, so
  assert with `expect-text` rather than dumping text immediately.
- The AI documentation tab is present and disabled without `--ai`. A proof of
  AI documentation needs a launch with a model and its own feature file.
- The graph canvas is drawn by React Flow and lays itself out. Selecting a
  node by clicking the canvas is not scripted here; do not report it verified.
