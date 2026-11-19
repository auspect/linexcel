---
name: verify-linexcel
description: >-
  Drive the real linexcel from this checkout and capture proof that a behaviour
  works: the `linexcel` command line (analyze, macros, dry-run), the standalone
  HTML lineage report it writes, and the local web application started by
  `linexcel serve`. Use after changing analysis, the report viewer, the app or
  the CLI, when asked to "check it in the real app", to reproduce a bug as a
  user would, or to attach screenshots and transcripts to a PR. It launches a
  disposable instance, never the user's own workspace.
---

# Verify linexcel by driving it

A user reaches linexcel in three ways. This skill drives all three from
the working tree and records what happened:

| Entry | What the user does | Driven with |
| --- | --- | --- |
| Command line | `linexcel analyze book.xlsx`, `linexcel macros book.xlsm` | `ctl cli` |
| HTML report | opens the `*_lineage.html` file that `analyze` wrote | `ctl browser --open FILE` |
| Local application | `linexcel serve`, opens the printed link, imports a workbook | `ctl browser --open app` |

The Python API (`from linexcel import analyze`) is a fourth way in. It is not
driven here: `tests/` covers it, and the command line calls the same functions.

`ctl` below stands for this exact prefix, run from the repository root:

```
uv run --extra screenshots python .agents/skills/verify-linexcel/scripts/control_linexcel.py
```

`uv run` makes the working tree the code under test (the package is installed
editable). The `screenshots` extra brings Playwright. If Chromium is missing,
run `uv run --extra screenshots playwright install chromium` once.

Read [`features/README.md`](features/README.md) before driving anything. It
indexes one recipe file per feature, with every entry point a user has. A
proof that drives one convenient entry point is incomplete when the map lists
others for the behaviour you changed.

## What this does not replace

`AGENTS.md` requires a full `validate_manual.py` run, with an independent
visual review at 1440×900 and 390×844, before delivering a change to analysis,
evaluation, documentation, screenshots or the viewer. This skill is the fast,
targeted proof of one behaviour. It does not satisfy that acceptance run, and
a pass here says nothing about the features you did not drive.

## Launch

```
ctl launch --run <run>
```

`<run>` is a short id you choose, such as `pr160-target`. Omit it for a
timestamp. Launch:

1. creates `validation_screenshots/verify-linexcel/<run>/` (git ignores it),
   with `scratch/` and `evidence/`;
2. builds the two synthetic workbooks from `validation_workbooks.py` into
   `scratch/fixtures/`: `fy26_commercial_review.xlsx` (sales) and
   `validation_stress.xlsx` (stress);
3. starts `python -m linexcel serve --port 0 --data-dir scratch/data`, on a
   free loopback port, with an empty private data directory;
4. waits until the server has published itself in `scratch/data/.access/` and
   `GET api/config` returns its limits. A listening port alone is not ready.

It prints `ready: http://127.0.0.1:<port>/ (pid …, linexcel <version>)`. No
migration or seed step exists: the application starts usable from an empty
directory, and each recipe imports the workbook it needs.

Nothing contacts a model. To verify AI features, pass serve options after
`--`, for example `ctl launch --run <run> -- --ai --ai-model qwen3.8`, and say
in your report which model answered. A requested AI stage that fails is a
failure, not a skip.

The command line and the report need no server, but `launch` is still how a
run gets its fixtures and evidence directory.

## Doctor

```
ctl doctor --run <run>
```

Read-only. Run it first whenever something looks off. It answers "is this
instance worth driving?":

- `process`: the server pid recorded at launch is alive;
- `ownership`: the data directory is still published by that pid;
- `http`: `api/config` answers 200 with the local profile;
- `session`: the profile that `linexcel session-link` needs is present;
- `build`: `src/linexcel` is byte-identical to what it was at launch. The
  server does not reload code. After an edit this check fails: run cleanup,
  launch again, then drive. Evidence from a stale server proves the old code;
- `fixture:*`: both workbooks exist.

Exit 0 means healthy. `ctl browser --open app` runs the same checks and
refuses to drive an unhealthy run.

## Drive

Command line, recorded with its output and exit code:

```
ctl cli --run <run> --name <label> [--expect-exit N] -- <linexcel arguments>
```

`{sales}`, `{stress}`, `{macros}`, `{fixtures}`, `{evidence}` and `{scratch}`
expand to paths of the run. Write reports under `{evidence}/reports/`, which
launch creates (`analyze` does not create parent directories).

A second, read-only view of what `analyze --json` wrote:

```
ctl graph --run <run> --json "{evidence}/reports/sales.json" --node "c:Summary!C10"
```

A browser session, as a list of steps:

```
ctl browser --run <run> --open app --name <label> [--viewport desktop|mobile] --step "<verb> <args>" --step …
ctl browser --run <run> --open "{evidence}/reports/sales.html" --name <label> --step …
```

A step is a verb, then arguments separated by ` :: `:

| Step | Effect |
| --- | --- |
| `click SELECTOR` | click the first match |
| `fill SELECTOR :: TEXT` | replace the content of an input |
| `press KEY`, `press SELECTOR :: KEY` | press a key on the page or an element |
| `select SELECTOR :: VALUE` | choose an option of a `<select>` |
| `upload SELECTOR :: FILE` | set the files of a file input |
| `wait SELECTOR`, `wait-hidden SELECTOR` | wait for visibility |
| `wait-tasks` | app only: wait until no server task is queued or running |
| `expect-text SELECTOR :: TEXT` | wait until the element contains the text, or fail |
| `expect-value SELECTOR :: VALUE`, `expect-count SELECTOR :: N`, `expect-hidden SELECTOR` | assertions that wait |
| `shot NAME`, `shot-full NAME` | screenshot to `NAME.png` |
| `aria NAME`, `aria NAME :: SELECTOR` | ARIA snapshot to `NAME.aria.txt` |
| `text NAME :: SELECTOR` | visible text to `NAME.txt` |
| `download FILENAME :: SELECTOR` | click and save the download |

Selectors are Playwright selectors. Use the element ids the application and
the report already carry (`#tab-graph`, `#graph-search`, `#lin-search`), or a
role (`role=button[name='Explore this sheet']`). Do not use coordinates or tab
order. Take an `aria` snapshot when you need to discover a handle.

The first `--open app` session of a run enters through a fresh
`linexcel session-link`, as a new browser does. Later sessions reuse the saved
browser state and reopen the last project, as a returning browser does. Pass
`--fresh` to enter as a new browser again. Viewports are 1440×900 (`desktop`)
and 390×844 (`mobile`).

A session fails, with exit 1 and `FAILED.png`, when a step fails, when the
page raises a JavaScript error, or when it requests a non-loopback address.
The last two are product promises (the report is offline and self-contained),
so do not pass `--allow-page-errors` or `--allow-external` to get a pass.

## Evidence

Everything lands in `validation_screenshots/verify-linexcel/<run>/evidence/`:

- `run.json`: URL, serve command, version and commit of the instance;
- `cli/<label>.json` with `<label>.stdout.txt` and `<label>.stderr.txt`: the
  command, its exit code and its output;
- `reports/`: files the command line wrote;
- `browser/<label>/`: `session.json` (every step with its result, page errors,
  external requests), screenshots, ARIA snapshots and text dumps.

Proof standards:

- Drive the path a user takes: the command, the file, the link. Do not call
  internal functions, post to `api/` by hand, or edit the data directory.
- Capture the action and the state it produced. A final screenshot alone does
  not show what was done; `session.json` and the CLI record do.
- Check the side effect in a second view. After `analyze`, read the JSON with
  `ctl graph` and open the HTML. After a calculation in the app, the
  Operations tab must list the task as completed.
- Compare against the workbook, not against the last output. Expected values
  in the recipes come from formulas in `validation_workbooks.py`.
- UI changes need both viewports.
- `analyze --dry-run` reads the workbook and prints to stdout only. It writes
  no report, no JSON and starts no worker. If your change touches it, confirm
  that by listing the output directory before and after.
- Evidence can embed workbook content. Keep it under `validation_screenshots/`
  and never commit it. For a PR, attach or quote the files you chose.

Report what you drove, what you did not, and every step that failed, with the
path of its evidence. Do not report an entry point you skipped as verified
through another one.

## Cleanup

```
ctl cleanup --run <run>
```

It stops the server pid this run started, with the analysis workers under it,
only after confirming that the data directory is still published by that pid.
It never kills by process name. It then deletes `scratch/` (data directory,
fixtures, server log, browser state) and keeps `evidence/`. Run it after a
failed attempt too, so a broken run does not strand a server and a port.

To discard old evidence, delete `validation_screenshots/verify-linexcel/<run>/`
yourself once nobody needs it.

On Windows the data directory nests owner, project and task ids past the
260-character path limit. The server and `ctl cleanup` handle that with an
extended path prefix; `Remove-Item`, Explorer and a plain `shutil.rmtree` on
`scratch/` can fail with "directory not empty" or "path not found". Always
remove a run's scratch state with `ctl cleanup`.

## Isolation

- Each run has its own port and data directory, so several runs can exist side
  by side. Use a different `--run` id per agent.
- Never point anything at `~/.linexcel`, the default of `serve`, `library`,
  `doctor` and `session-link`. It is the user's real workspace. `ctl cli` adds
  the run's `--data-dir` to those commands when you omit it.
- Never drive a server you did not launch, including one on port 8765.
- One data directory has one owner. While the server runs, `linexcel library …`
  on the same directory exits 2 with "This data directory already has a
  coordinator". Verify library search in the application, or stop the server.

## Helper

`scripts/control_linexcel.py` is the only script. Its commands are `launch`,
`doctor`, `cli`, `graph`, `browser` and `cleanup`, all shown above;
`ctl <command> --help` lists the options. It needs nothing beyond the
`screenshots` extra.
