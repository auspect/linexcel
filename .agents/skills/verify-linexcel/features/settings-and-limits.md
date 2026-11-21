# Settings and limits

In the local application, a user sees the limits that stop an operation next
to the button that starts an import, sees how much of the server's storage
quota the workspaces use, and changes limits, AI documentation settings,
backup restore and feature checks in one settings dialog. An operation stopped
by a limit names the setting that changes it.

## Sub-features

- `settings-open` opens the settings from the sidebar's `Settings` entry, and
  from `Change` on the limits line of the import form, on the limits section.
- `settings-limits` sets the duration of long operations in minutes, the
  targeted recalculation in seconds and the memory per operation in GB; the
  limits line follows.
- `settings-advised` puts the three fields back to the server's recommended
  values.
- `settings-refuse` refuses a value outside its range, under the field and
  again when an operation starts.
- `storage-line` shows the storage in use and the server quota under the
  workspace list.
- `limit-failure` shows, in Operations, a stopped operation with the value it
  reached and the setting that raises it.

## How to get to it (user POV)

- The `Settings` entry pinned at the foot of the sidebar (`Menu` first on a
  narrow screen). It opens on the section shown last, `Operation limits` at
  first.
- In the open `Import a workbook` form, the line `Limits: 60 min · 120 s · …`
  above `Import`, then `Change`.
- In the dialog, the sections `Operation limits`, `AI documentation`,
  `Backup and recovery` and `Feature availability`: a side list on a desktop,
  a row of pills on a phone. `Done`, the close button and Escape close it.
- `Local storage` under the workspace list.
- A stopped operation: the `Operations` tab of its workspace.

## Driving it with ctl

Preconditions:

- `ctl doctor --run <run>` exits 0.
- The workspace `Verify sales` exists: run the **Import** steps of
  [Import a workbook](./import-workbook.md) first, in the same run.
- `Summary!C6` has not been calculated in this run (no other recipe selects
  it), so selecting it starts a new calculation.

Run `ctl browser --run <run> --open app --name settings-desktop` with these
steps in order.

- **Storage line.** Steps `expect-text #project-title :: Verify sales`,
  `expect-text #storage :: Local storage`, `expect-text #storage-value :: GB of`
  and `aria storage :: #storage`. The meter's value text repeats the line.
- **Open from the sidebar.** Steps `click #settings-open`,
  `expect-text #settings-limits :: Operation limits`, `aria settings :: #settings`
  and `shot settings-limits`. Then `press Escape` and `expect-hidden #settings`.
- **Open from the limits line.** Steps `click #import-panel > summary`,
  `expect-text .limits-line :: Limits: 60 min`,
  `expect-text .limits-line :: 120 s`, `shot import-limits`,
  `click #limits-edit` and
  `expect-text #settings [aria-current='true'] :: Operation limits`.
- **Refuse a value.** Steps `fill #limit-memory-gb :: 999`,
  `press #limit-memory-gb :: Tab` and
  `expect-text #limit-memory-gb-error :: choose a value from 0.5 to`. The
  message names the field, `Memory per operation`, and the server ceiling.
- **Advised values.** Steps `click #limits-reset`,
  `expect-value #limit-long-minutes :: 60`,
  `expect-value #limit-targeted-seconds :: 120` and
  `expect-hidden #limit-memory-gb-error`.
- **Change a limit.** Steps `fill #limit-targeted-seconds :: 1`,
  `click #settings-done` and `expect-text .limits-line :: 60 min`, then
  `expect-text .limits-line :: 1 s`.
- **Read a limit failure.** Steps `click #tab-nodes`,
  `fill #search :: Summary!C6`, `expect-count #node-list [data-node] :: 1`,
  `click #node-list [data-node='Summary!C6']`, `wait-tasks`, `click #tab-tasks`,
  `expect-text #task-list :: Time limit of 1 s reached. Raise`,
  `expect-text #task-list :: in Settings`,
  `expect-text #task-list :: Operation limits (up to 3600 s), then run it again`,
  `text tasks :: #task-list` and `shot limit-failure`. The full sentence names
  the field, `Targeted recalculation`, between curly quotes.
- **Second view of the storage.** Steps `click #settings-open`,
  `click #settings [data-settings='health']`, `click #health-check`,
  `expect-text #health-status :: Storage` and `text health :: #health-status`.
  The storage check lists the same usage and quota in MiB.
- **Mobile.** Run `ctl browser --run <run> --open app --name settings-mobile --viewport mobile`
  with steps `expect-text #project-title :: Verify sales`,
  `click #workspace-menu`, `expect-text #storage :: Local storage`,
  `click #settings-open`, `expect-text #settings .settings-nav :: Features`,
  `shot settings-mobile`, `click #settings-done` and `shot sidebar-mobile`.
  The dialog fills the screen and `Done` spans its width.

## Gotchas

- The memory value depends on the machine: the server ceiling is half of its
  memory in whole GB, and each operation gets an equal share of it in 0.5 GB
  steps. Assert minutes and seconds, or read the GB value from the help line
  under the field (`Server ceiling: … Recommended: …`).
- The helper opens the application with `?lang=en`. French text needs
  `select #ui-language :: fr` first.
- The fields hold minutes, seconds and GB; the server receives seconds and
  MiB. Nothing is saved on the server: a reload shows the advised values again.
- A cell already calculated at this revision reuses its result: no task starts
  and nothing fails. Use a cell no other recipe selected.
- With the two fixtures the storage line reads `< 0.1 GB of … GB`. It changes
  visibly only when gigabytes move, so compare it with the health check rather
  than before and after a deletion.
- A 1 s targeted recalculation fails because the isolated worker takes longer
  than that to start. A recipe that needs a successful calculation afterwards
  must put the advised values back first.
- Windows PowerShell 5.1 passes non-ASCII step text to Python mangled (`é`
  arrives as `?`), so these steps assert ASCII fragments. In Git Bash the full
  sentences work.
