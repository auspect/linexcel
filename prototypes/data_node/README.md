# Data node prototype (batch 2)

An isolated prototype, not part of the product: nothing in `src/linexcel`
imports it, it is not in the wheel, and the source distribution excludes this
folder. The findings, measurements and the integration proposal are in
[`docs/reading_resilience_data_node.md`](../../docs/reading_resilience_data_node.md).

## What it shows

- `data_node.build_store` streams a workbook with serve's own reader
  (`lazy_store._read_source`, unchanged) and only replaces its storage: literal
  values go to one SQLite grid per sheet (`row` + one typed column per Excel
  column), with exceptions for booleans, errors, ISO dates and unreadable
  values, and number formats kept as runs of rows per column. Formula cells
  stay per-cell nodes, as today.
- `data_node.detect_regions` recognises a data region: an Excel table (header
  row and column names from the table part, formula columns split off), or a
  run of formula-free columns with a detected header row. Smaller blocks
  (under 10,000 cells) and columns that mix formulas and literals stay per
  cell. Each region becomes one node with its bounds, headers, schema (type
  per column), number formats, row count, a SHA-256 fingerprint of its values
  and a bounded preview.
- `data_node.link` points the formulas' range references to the region with
  the sub-range they read.
- `bench.py evaluate` runs serve's own targeted calculation
  (`lazy._evaluate_index`) twice: unchanged, then with an index whose range
  selection loads region sub-ranges into formualizer in row chunks with
  `set_values_batch`. Closure, refusal and cache comparison are serve's.
- `data_node.window` serves a presentation window by a range query.

No value is recalculated or substituted: the store holds the saved file values.
An unreadable value stays an `unreadable` exception and makes a calculation
that reads it refuse.

## Running it

From the repository root, with the corpus outside the repository:

```powershell
$W = "<linexcel-lab>\workbooks"
uv run python prototypes/data_node/bench.py build $W\public\dgt-elections-2017-2020-tour1.xlsx <scratch>\dgt.sqlite dgt-build.json
uv run python prototypes/data_node/bench.py windows <scratch>\dgt.sqlite Feuil1 dgt-windows.json
uv run python prototypes/data_node/bench.py verify <scratch>\dgt.sqlite <serve index.sqlite> dgt-verify.json 3000
uv run python prototypes/data_node/bench.py link <scratch>\fpna.sqlite <serve index.sqlite> fpna-link.json
uv run python prototypes/data_node/bench.py evaluate $W\generated\large\fpna-20mb-seed42.xlsx <scratch>\fpna.sqlite <serve index.sqlite> "Forecast!D2" fpna-d2.json
```

A serve index is built with `linexcel.lazy_store.build_index` (see
`probes/measure_import.py`). Outputs, stores and indexes belong in the ignored
`validation_screenshots/` directory or outside the repository.

`probes/` holds the measurement scripts used for the batch 2 note: import cost
(`measure_import.py`, `profile_index.py`), readers (`reader_bench.py`),
formualizer loading (`probe_fz_*.py`), the serve path end to end
(`serve_path.py`, `probe_data_lock.py`, `probe_views.py`) and damaged copies of
public workbooks (`make_damaged.py`). `reader_bench.py` with `fastexcel` or
`polars` runs in a separate throw-away environment; neither is a linexcel
dependency.

## Known limits

- Detection is a heuristic: a header must be all text over at least half the
  columns, within the first 20 rows, followed by a row with a non-text value.
  Two-row headers (names, then codes) are not recognised as such.
- A column with any formula in the sheet is excluded from a region, even when
  the formula lies outside the region rows.
- The raw XML spelling of a number (`"5.10"`) is not kept, only its value;
  presentation that needs the spelling must read the part.
- Style ids, row heights, column widths and merges are not indexed yet; only
  number formats are.
- Grids wider than 512 columns keep their extra cells in a long `wide` table;
  regions beyond column 512 are not created.
