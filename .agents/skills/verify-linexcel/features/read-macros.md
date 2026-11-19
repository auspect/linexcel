# Read macros

A user asks for the VBA source stored in a macro-enabled workbook or an
add-in. linexcel prints it as JSON, with provenance and stated limits, and
never executes it.

## Sub-features

- `macros-stdout` prints the report on stdout.
- `macros-file` writes the report to a file with `-o`.
- `macros-not-executed` states, in the report, that nothing was run.
- `macros-not-excel` refuses a file that is not a workbook, with exit 2.

## How to get to it (user POV)

- Run `linexcel macros WORKBOOK` in a terminal.
- Run `linexcel macros WORKBOOK -o report.json`.
- Call `linexcel.read_macros(path)` from Python.
- Analyze the same workbook: its VBA modules appear as nodes of the report.

## Driving it with ctl

Preconditions:

- `ctl launch --run <run>` created the run. No server is needed.
- `{macros}` is `tests/fixtures/macros.xlsm`, a checked-in file built with
  Excel.

- **Print.** Run
  `ctl cli --run <run> --name macros-stdout --tail 12 --expect-exit 0 -- macros "{macros}"`.
  Exit code `0`, stderr empty. `cli/macros-stdout.stdout.txt` is one JSON
  object with `"mode": "static_vba_source"`, `"execution": "not_executed"`,
  `"status": "read"` and a `modules` entry `Module1` whose text contains
  `Public Sub Refresh()`.
- **Write a file.** Run
  `ctl cli --run <run> --name macros-file --expect-exit 0 -- macros "{macros}" -o "{evidence}/reports/macros.json"`.
  Stdout is empty and `{evidence}/reports/macros.json` holds the same object.
- **Limits are stated.** The `limitations` list of the report includes
  `No macro execution, UDF calculation, p-code audit or Excel 4.0 macro lineage.`
- **Not a workbook.** Run
  `ctl cli --run <run> --name macros-not-excel --expect-exit 2 -- macros "pyproject.toml"`.
  Exit code `2`, stdout empty, stderr is one line starting with `linexcel:`.

## Gotchas

- Exit `3` means a partial read (`"status": "partial"`). The checked-in
  fixture reads completely, so that path is unreachable with the seed data.
  Report it as not verified.
- The report quotes VBA source verbatim. For a real workbook that is private
  content: keep it in the evidence directory and out of a PR.
- `"execution": "not_executed"` is a statement in the output. The stronger
  proof that nothing runs is that the command needs no Excel and finishes in
  about a second on a machine without Office.
- The module text is decoded source. It is not a byte-for-byte copy of the
  binary stream; compare procedures, not hashes of the file.
