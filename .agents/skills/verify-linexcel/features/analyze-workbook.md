# Analyze a workbook

A user points `linexcel analyze` at an Excel workbook and gets a standalone
HTML lineage report, optionally the dependency graph as JSON, without any
model being contacted.

## Sub-features

- `analyze-report` writes the HTML report and prints a summary on stderr.
- `analyze-json` writes the graph as JSON beside or instead of the report.
- `analyze-target` limits the lineage to what feeds one cell.
- `analyze-dry-run` describes the file and stops before analysing it.
- `analyze-failure` exits 2 with a one-line reason for an unreadable workbook.
- `analyze-version` reports the version of the checkout.

## How to get to it (user POV)

- Run `linexcel analyze WORKBOOK` in a terminal (console script).
- Run `python -m linexcel analyze WORKBOOK` when the script is not on `PATH`.
- Call `linexcel.analyze(path)` then `result.save_html(...)` or
  `result.save_json(...)` from Python.

## Driving it with ctl

Preconditions:

- `ctl doctor --run <run>` exits 0 (the fixtures exist).
- `{evidence}/reports/` holds no `sales.*` or `target.*` file yet.

- **Analyze.** Run the analysis with both outputs. Run
  `ctl cli --run <run> --name analyze-sales --expect-exit 0 -- analyze "{sales}" -o "{evidence}/reports/sales.html" --json "{evidence}/reports/sales.json"`.
  Exit code `0`. Stdout is empty. Stderr names the `JSON:` and `HTML:` paths
  and ends its summary with `Nodes: 20  Edges: 21  Sheets: 4`.
- **Second view of the JSON.** Read the node behind the KPI "Net Commercial
  Contribution". Run
  `ctl graph --run <run> --json "{evidence}/reports/sales.json" --node "c:Summary!C10"`.
  `executionStatus` is `completed`, `totalFormulas` is `50`, the node formula
  is `=ROUND(C8-C6*RebateRate-C8*Assumptions!$D$8, 2)`, its value is `816668.1`
  with `valueSource` `engine`, and `precedents` are exactly `c:Summary!C6`,
  `c:Summary!C8`, `i:Assumptions!D8` and `n:RebateRate`: one per reference in
  the formula.
- **Second view of the HTML.** Open the report as in
  [Read the lineage report](./lineage-report.md).
- **Target.** Trace one cell only. Run
  `ctl cli --run <run> --name analyze-target --expect-exit 0 -- analyze "{sales}" --target "Summary!C10" --no-html --json "{evidence}/reports/target.json"`.
  Exit code `0`, stderr reports `Nodes: 13  Edges: 13  Sheets: 4` and a
  `warning: Targeted analysis of Summary!C10` line stating that the lineage is
  limited to the static upstream trace. No HTML file is written.
- **Dry run.** Ask what the file declares. Run
  `ctl cli --run <run> --name analyze-dry-run --expect-exit 0 -- analyze "{sales}" --dry-run`.
  Stdout lists `4 sheet(s), 747 cells declared`, one line per sheet and the
  ceilings. Stderr is empty and no `fy26_commercial_review_lineage.html`
  appears in `{fixtures}`.
- **Failure.** Give a path that does not exist. Run
  `ctl cli --run <run> --name analyze-missing --expect-exit 2 -- analyze "{fixtures}/missing.xlsx"`.
  Exit code `2`, stdout empty, stderr is one line starting with `linexcel:`.
- **Version.** Run `ctl cli --run <run> --name version --expect-exit 0 -- --version`.
  Stdout is `linexcel <version>`; on a commit that is not a release tag it
  carries a `+dev` segment.
- **Console script entry.** `ctl cli` uses `python -m linexcel`. To prove the
  console script itself, run `uv run linexcel --version` and compare.
- **Python entry.** Not driven by `ctl`. Run
  `uv run python -c "from linexcel import analyze; r = analyze(r'<path to the sales fixture>'); print(len(r.nodes), len(r.edges))"`
  and expect `20 21`.

## Gotchas

- `-o` and `--json` do not create parent directories: a missing folder gives
  exit `2` and `No such file or directory`. `{evidence}/reports/` exists.
- Without `-o`, the report is written beside the workbook as
  `<stem>_lineage.html`, here inside `scratch/fixtures/`, which cleanup deletes.
  Pass `-o` when the file is evidence.
- Progress, statistics, warnings and hints go to stderr. Only `--dry-run`,
  `-o -` and `--json -` write to stdout.
- The fixture is authored by openpyxl, so it stores no cached values: every
  formula value comes from the engine (`valueSource` `engine`) and nothing can
  be compared with an Excel cache. Use a workbook saved by Excel to verify
  cache comparison.
- The hint about `--screenshots` appears on every run with more than one sheet.
  Sheet rendering needs LibreOffice and Poppler and is not part of this recipe.
- Exit `3` (incomplete analysis) and `130` (cancelled) cannot be reached with
  these fixtures within the default 120-second budget. Do not claim them.
- The counts above describe `build_sales_workbook` at the commit where this
  file was last checked. If `validation_workbooks.py` changed, recount from
  the workbook.
