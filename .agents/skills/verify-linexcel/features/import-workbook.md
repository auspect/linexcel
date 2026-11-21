# Import a workbook

A user starts the local application with `linexcel serve`, opens the access
link it prints, and imports a workbook. The application indexes its formulas,
dependencies and saved values into a named workspace, without recalculating
the workbook.

## Sub-features

- `import-access` opens the application through a single-use access link.
- `import-upload` imports a workbook chosen in the browser.
- `import-index` shows the indexed node and relationship counts.
- `import-return` reopens the last workspace for a returning browser.
- `import-server-path` imports a workbook from a path on the server's disk.

## How to get to it (user POV)

- Run `linexcel serve` and open the `Linexcel : http://127.0.0.1:<port>/#session=…`
  link it prints.
- Run `linexcel session-link --data-dir DIR` to open the running application
  in another browser.
- On the welcome screen, choose `Choose an Excel workbook`. It opens the
  `Import a workbook` form of the sidebar.
- In the sidebar (the `Menu` button on a narrow screen), open
  `Import a workbook`, choose `+ Choose workbook`, then `Import`.
- In the same form, set `Workbook source` to `File on the server` and type an
  absolute path.

## Driving it with ctl

Preconditions:

- `ctl doctor --run <run>` exits 0.
- For `import-upload`, no workspace is named `Verify sales`.

- **Open through a link.** A new browser enters with a fresh link. Run
  `ctl browser --run <run> --open app --name import-upload --fresh` with the
  steps below. The helper obtains the link from `linexcel session-link`, the
  same link `serve` prints. Step `expect-text #connection :: Server connected`.
- **Landing proof.** Steps `aria landing` and `shot landing`. On a new data
  directory the sidebar reads `No workbook imported.`
- **Choose the workbook.** Steps `upload #workbook :: {sales}` and
  `fill #project-name :: Verify sales`.
- **Import.** Step `click #import-button`. The heading becomes the workspace
  name: step `expect-text #project-title :: Verify sales`.
- **Index.** Wait for the index, then read it. Steps `wait-tasks`,
  `expect-text #stats :: 531 nodes`, `expect-text #stats :: 225 relationships`
  and `expect-text #stats :: 0 active operation(s)`.
- **Second view.** The Operations tab lists what ran. Steps `click #tab-tasks`,
  `expect-text #task-list :: Structural import`, `text tasks :: #task-list`.
  Then `click #tab-diagnostics`, `expect-text #view-diagnostics :: Import complete`,
  `aria imported :: main`, `shot imported`.
- **Returning browser.** Run a second session without `--fresh`:
  `ctl browser --run <run> --open app --name import-return --step "expect-text #project-title :: Verify sales" --step "expect-text #stats :: 531 nodes" --step "shot returned"`.
  The workspace reopens without any click.
- **Server path.** Import the second fixture from disk. Run
  `ctl browser --run <run> --open app --name import-server-path` with steps
  `expect-text #project-title :: Verify sales` (the reopened workspace),
  `click #import-panel > summary`, `select #import-source :: local`,
  `fill #local-path :: {stress}`, `fill #project-name :: Verify stress`,
  `click #import-button`, `expect-text #project-title :: Verify stress`,
  `wait-tasks`, `expect-text #stats :: 0 active operation(s)`,
  `text stats :: #stats`, `shot stress`. The sidebar lists both workspaces.

## Gotchas

- An access link works once and expires after ten minutes. Never paste the
  link from `scratch/serve.log` into a step: it may already be used. The
  helper mints a new one for each `--fresh` session.
- Without a session, protected routes such as `api/projects` and `api/health`
  answer 401, while `api/config` stays readable. A 200 on `api/config` proves
  the server is up, not that a browser is signed in.
- Right after `Import`, `#stats` reads `0 nodes · 0 relationships · 1 active
  operation(s)`. Assert the final counts with `expect-text`, which waits; a
  `text` dump taken immediately records the transient state.
- `Choose an Excel workbook` exists only on the welcome screen, while no
  workspace is open. Once one is open, use `#import-panel > summary`.
- The import form collapses when the last workspace finishes loading, so a
  click on `Import a workbook` in the first second is undone. Wait for the
  reopened workspace with `expect-text #project-title :: …` before opening it.
- On `mobile` the sidebar is closed. Add `click #workspace-menu` before any
  step that touches the import form.
- Each import adds a workspace to the sidebar, newest first. A recipe that
  expects `Verify sales` to be open must run before one that imports another
  workbook, or reopen it with `click #projects [data-project] >> text=Verify sales`.
- The stress fixture indexes to 1945 nodes and 1014 relationships.
- The node counts of the application (531) and of `analyze` (20) differ on
  purpose: the application indexes cells one by one, the report groups them.
- `File on the server` is offered only by the standalone launcher on loopback.
  `linexcel serve --no-local-import` removes it.
