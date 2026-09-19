# Command line

Everything the Python API does deterministically, without writing any Python:

```bash
uvx linexcel analyze workbook.xlsx
```

`uvx` runs it without installing anything. Installed in a project, the same
command is available as `linexcel`, or as `python -m linexcel` when the
console script is not on `PATH`:

```bash
uv add linexcel                 # pip install linexcel
uv run linexcel analyze workbook.xlsx
```

Deterministic by default: nothing contacts a model unless you ask for it.

## Output

Writing `<workbook>_lineage.html` next to the workbook is the default.

| Option | |
|---|---|
| `-o`, `--output PATH` | Where the HTML goes. `-` writes to stdout. |
| `--json PATH` | Also write the graph as JSON. `-` writes to stdout. |
| `--no-html` | Skip the viewer — use with `--json`. |
| `--language {en,fr,es,de,it,pt,nl,ja,zh}` | Interface and AI prompt language (default `en`). See [Languages](languages.md). |

Progress, statistics and warnings go to **stderr**, so `-` is safe to pipe:

```bash
linexcel analyze workbook.xlsx --json - --no-html | jq '.meta.stats'
```

## Before committing to a long run

`--dry-run` inspects declared dimensions, external links and extraction ceilings.
Declared dimensions are not occupied-cell counts, and file size cannot predict
execution time. No ETA is displayed.

```bash
linexcel analyze workbook.xlsx --dry-run
linexcel analyze workbook.xlsx --analysis-seconds 120 --memory-mb 2048 -v
```

Normal analysis is isolated with a configurable 120-second budget. AI,
screenshots and export have separate limits. `--time-budget` remains the
cooperative decomposition limit, not the hard analysis deadline. See
[execution budgets](execution.md) for statuses, compatibility and platform
memory semantics. Interrupted CLI analyses export diagnostics with exit code 3;
cancellation exits 130 without starting optional stages.

## Watching it run

`--verbose` draws progress per sheet through the two phases that take the
time — reading the values the file stores, and sweeping it for formulas —
then prints what each phase cost. It writes to stderr, so it never mixes into
`-o -` or `--json -`, and it draws bars only when stderr is a terminal: a CI
log gets one plain line per phase instead of thousands of redraws.

Install `linexcel[progress]` for the bars. Without it the same phases and the
same timings are printed unadorned; nothing else changes.

The exit status is `0` on success, `2` on a failure the tool recognises (an
unreadable workbook, a contradictory pair of options), and `130` on Ctrl-C.

## What goes into the graph

| Option | |
|---|---|
| `--refs-dir DIR` | Folder holding the workbooks this one links to, and the add-ins whose VBA it calls. Without it a cell reading another file is named, never resolved — see [Other workbooks](coverage.md#other-workbooks). |
| `--screenshots DIR` | Render each sheet to a PNG and show it in the report. Needs LibreOffice and Poppler; see [Screenshots](context.md#screenshots). |
| `-v`, `--verbose` | Progress while it runs, and per-phase timing, on stderr. |
| `--time-budget SECONDS` | Ceiling on the step-by-step decomposition (default 300). Past it, cells keep their values and lose only their breakdown, and the report says so. |
| `--target SHEET!A1` | Trace the static upstream lineage of this cell without requesting global recalculation. Repeatable; comma-separated cells are accepted, with Excel quotes around sheet names containing commas (`'Costs, FY26'!A1`). Only relevant names, VBA and Power Query context remain. Dynamic references (INDIRECT/OFFSET) or a truncated trace can cause the engine to evaluate dependencies omitted from the graph; the report states this limitation. |
| `--dry-run` | Say what the file declares — sheets, declared size, linked workbooks, the ceilings that will apply — and stop without analysing it. |

## AI documentation

Opt-in, and it needs the `ai` extra — which `uvx` can add on the fly:

```bash
uvx --from "linexcel[ai]" linexcel analyze workbook.xlsx --ai-docs \
    --base-url http://localhost:11434/v1 --model <tag>
```

No provider is chosen for you. Name the endpoint on the command line, or set
`LINEXCEL_AI_BASE_URL`, `LINEXCEL_AI_MODEL` and `LINEXCEL_AI_API_KEY` — see
[Choosing an AI provider](providers.md#environment-variables).

| Option | |
|---|---|
| `--ai-docs` | Document the workbook and its nodes. |
| `--deterministic-only` | Lineage only. The default, and worth passing explicitly in a script. |
| `--base-url`, `--model`, `--api-key` | The endpoint. Prefer the environment variable for the key. |
| `--max-workers N` | Concurrent requests (default 4). |
| `--max-tokens N` | Cap the output of each individual response. |
| `--token-budget N` | Cap the whole run, in and out. See [Capping the bill](ai.md#capping-the-bill). |
| `--no-workbook-doc` | Document the nodes, skip the workbook overview. |

## Describing the screenshots

Separate from `--ai-docs`, and deliberately so: this is the only option that
puts a picture of a sheet in a request, and a picture shows every row on it.

```bash
uvx --from "linexcel[ai]" linexcel analyze workbook.xlsx \
    --screenshots shots/ --vision-docs --vision-model <a vision model>
```

| Option | |
|---|---|
| `--vision-docs` | Describe each rendered sheet with a multimodal model. Requires `--screenshots`, and is refused with `--deterministic-only`. |
| `--vision-model` | The model that looks at the images, when it differs from the one that writes. |

Read [Describing the screenshots](ai.md#describing-the-screenshots) before
trusting the result, and [Data handling](data-handling.md) for what leaves the
machine.

## Version

```bash
linexcel --version
```

A build that is not exactly a release tag reports a development version, such
as `1.3.0+dev.4.g7f5caf5`.
