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

The exit status is `0` on success, `2` on a failure the tool recognises (an
unreadable workbook, a contradictory pair of options), and `130` on Ctrl-C.

## What goes into the graph

| Option | |
|---|---|
| `--refs-dir DIR` | Folder holding the workbooks this one links to, and the add-ins whose VBA it calls. Without it a cell reading another file is named, never resolved — see [Other workbooks](coverage.md#other-workbooks). |
| `--screenshots DIR` | Render each sheet to a PNG and show it in the report. Needs LibreOffice and Poppler; see [Screenshots](context.md#screenshots). |
| `-v`, `--verbose` | Per-phase timing on stderr, for finding where a large workbook spends its minutes. |

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
