# linexcel verification map

This directory is the maintained source for verifying what a linexcel user can
do. Read this index before driving anything, then use the matching feature file
as the recipe. When a feature gains or loses an entry point, change its file in
the same commit.

## Baseline preconditions

- `ctl` is `uv run --extra screenshots python .agents/skills/verify-linexcel/scripts/control_linexcel.py`,
  run from the repository root.
- `ctl launch --run <run>` succeeded, and `ctl doctor --run <run>` exits 0.
- The fixtures are the synthetic workbooks built at launch from
  `validation_workbooks.py`: `{sales}` (`fy26_commercial_review.xlsx`, sheets
  Summary, Sales, Catalog, Assumptions, 50 formulas) and `{stress}`
  (`validation_stress.xlsx`). `{macros}` is `tests/fixtures/macros.xlsm`.
- No model is configured. AI tabs and buttons are present but disabled.
- Never drive an instance this run did not launch, and never use `~/.linexcel`.

## Driving conventions

- Start each recipe from the baseline unless its preconditions say otherwise.
  The application recipes build on "Import a workbook".
- Treat every command as literal, after replacing `<run>`. Keep `{sales}` and
  the other placeholders as written: the helper expands them.
- Command-line actions go through `ctl cli`. Browser actions go through
  `ctl browser`, one `--step` per user action, in the order given.
- Prefer element ids and ARIA roles. Never use coordinates or DOM position
  beyond "the only result".
- `wait-tasks` waits on the server. The page refreshes on its own poll, so
  follow it with an `expect-text` on what the user reads.
- Step strings use no inner quotes except single quotes inside a selector. In
  PowerShell, a `$` inside a double-quoted step is interpolated: none of the
  recipes needs one.
- Leave what a recipe created in place. `ctl cleanup` removes the whole
  instance and keeps the evidence.

## Proof and skip reporting

- Capture the user action and the resulting state, not only the final screen.
  `session.json` and the `cli/*.json` records hold the actions.
- UI proof includes an `aria` snapshot and a `shot` with the workbook or
  project name visible, at `desktop` and at `mobile` for a layout change.
- CLI proof is the recorded command, stdout, stderr and exit code.
- Proof of a written file includes a second view of it: `ctl graph` for JSON,
  a browser session for HTML.
- Name the feature id and the entry point with every artifact you cite.
- Report an unreachable path with the attempted command and the unmet
  precondition. Do not report a skipped entry point as verified through
  another one.
- When the expected value in a recipe no longer matches, find out which
  changed: the fixture in `validation_workbooks.py`, or the product. Derive the
  new expectation from the workbook's formulas, never from the new output.

## Feature entry contract

Each feature file starts with an H1 title and one paragraph describing what the
user can do. It then has exactly four H2 sections, in this order:

1. `Sub-features` lists short ids, one line per behaviour.
2. `How to get to it (user POV)` lists every entry point a user has.
3. `Driving it with ctl` starts with `Preconditions:` and pairs each user
   action with an exact command and the observable result.
4. `Gotchas` lists traps that waste or invalidate a run.

Keep implementation details out. Name user paths, stable handles, required
state, commands and observable proof.

## Features

- [Analyze a workbook](./analyze-workbook.md): `linexcel analyze`, its HTML and
  JSON outputs, targets, dry run and exit codes.
- [Read the lineage report](./lineage-report.md): the standalone HTML file,
  its search, node panel, sheets view, hidden tabs and offline guarantee.
- [Import a workbook](./import-workbook.md): `linexcel serve`, the access link,
  import from the browser or from a server path, the indexed project.
- [Follow a calculation](./follow-calculation.md): selecting a cell in the
  application, the targeted calculation, its dependencies and its task.
- [Read macros](./read-macros.md): `linexcel macros`, static VBA source as
  JSON, never executed.

Not mapped yet: sheet captures (LibreOffice), the data browser, AI
documentation and questions, library search, backup and restore, the
investigation export, external references with `--refs-dir`. A change to one
of those needs its own feature file before it can be called verified here.
