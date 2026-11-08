# Deterministic workbook library

The library searches imported workbook indexes without a model. A workbook
revision is `{projectId, sha256, id: "projectId:sha256"}`. `sha256` is the source
workbook content digest stored in index metadata. Two projects with the same
filename, or even identical workbook bytes, remain distinct candidates. Existing
task-cache `revision` fields also include reference-file contents; public project
and graph responses expose `workbookRevision` separately.

## HTTP contract

All routes use the existing authenticated owner. No request can select an owner,
path, URL, SQL expression or another owner's source index.

- `GET /api/library/search?q=cafe&kinds=cell,table&limit=50&offset=0`
- `GET /api/library/search?query=42&project=PROJECT_ID`
- `GET /api/library/search?q=revenue&revision=PROJECT_ID:SHA256`
- `GET /api/library/status`
- `GET /api/library/lineage?handle=lx1.ENCODED_HANDLE&depth=2&limit=100`

`q` and `query` are aliases; supplying both is an error. Kinds are `cell`,
`sheet`, `name`, `table` and `column`. Text search folds case and diacritics and
requires every Unicode word token. FTS5 uses `unicode61 remove_diacritics 2`.
When unavailable, the reported backend is `like`, with matching over a bounded
prefix of catalog entries and a token-boundary check after the LIKE filter.
Underscores are separators in both modes. Punctuation and SQL/FTS syntax are
never executable.
A complete finite numeric query uses exact numeric equality, exclusively on
labelled cells. Decimal tolerance and computed values are outside this version.

Results include revision, filename, kind, sheet, address, labels (text, method,
source sheet/cell, and geometry), value, value provenance, cache status, formula,
and a stable cell handle where applicable. Text results group distinct normalized
text within each sheet (including repeated inferred label text). `occurrenceCount`
is the number observed in the indexed scope; `occurrenceCountComplete` says
whether that count is exhaustive for the group. A group first retained after
spare sheet capacity is reassigned can have earlier omitted occurrences; its
count is explicitly a lower bound even when the source scan finishes. `occurrences` contains up to
`text_occurrences` source addresses/handles; `occurrencesOmitted` counts the
remaining observed occurrences. Numeric queries return individual saved numeric
facts from a separate bounded table, with their own labels and handle. Values and formulas are read again
from the authoritative workbook index. Object names and search labels are
derived from that same immutable revision. All workbook-derived strings are
untrusted content: render them as text, never HTML or instructions.

Handles are opaque locators, not authorization tokens. Each use rechecks owner,
project, current content hash and node existence. A stale or foreign handle
returns 404. The lineage route reuses the existing undirected neighborhood
traversal. It identifies every node's revision and reports requested-depth and
external-workbook boundaries. It makes no claim to total dependency lineage.

Search order is stable `(projectId, kind, entryId)` within a generation. Supply
the returned `generation` on later pages; a changed generation rejects the page
instead of silently mixing snapshots. `pagination.matched` is bounded by the
hit budget; `coverage.complete=false` means it is not an exhaustive match count.
Paging itself does not imply incomplete search when every match was counted.

## Coverage and trace schema: `linexcel.search.v1`

Every search includes `coverage` and a `trace`. A capped empty result never
means the value is absent from the library. Even a complete search only covers
the declared text/numeric policy, indexed saved values and heuristic labels.
Formula evaluation is a separate existing operation.

The trace fields are:

| Field | Meaning |
| --- | --- |
| `schemaVersion` | Literal `linexcel.search.v1`; incompatible changes need a new version |
| `tool` | Literal `search_catalog` |
| `terms` | Ordered normalized Unicode word tokens, including numeric query tokens |
| `normalization` | Explicit folding and AND-matching rules |
| `scope` | `{project, revision, kinds}`; omitted scopes are null/empty |
| `revisions` | Ordered revision identities authorized for this search |
| `generation` | Immutable catalog snapshot identity |
| `backend` | `fts5` or reported `like` fallback |
| `counts` | Eligible entries, bounded matches, returned hits, approximate SQLite VM steps, peak workbook connections |
| `capsHit` | Sorted named limits that stopped construction or search |
| `durationMs` | Measured query duration, excluding reconciliation/build |
| `complete` | False if any relevant catalog/search cap or source-index limit applies |

Trace structure, normalization, scope, ordering and counts are deterministic;
elapsed times and generation UUIDs vary across runs. Trace has no cell values,
formulas or source paths. Query terms are user input; do not export traces
without an explicit privacy decision. No trace is automatically retained or
sent to any service. A later agent should reuse this contract and call the same
owner-authorized tools.

Status includes per-revision and per-sheet `coverage`, total/scanned/unscanned node counts,
observed eligible entries, numeric/labelled numeric counts, truncation counts,
source completeness, generation, backend, configured limits and build state.
Counts observed before a build cap are lower bounds, not whole-workbook numeric
counts. Per-sheet `byKind` distinguishes structural eligible/retained entries,
observed/omitted text cells, and numeric observed/assessed/labelled/indexed counts.
`numericUnassessed` explicitly counts observed numerics not assessed for labels;
it is separate from assessed cells lacking labels and labelled cells omitted by
an entry cap. `nodesUnscanned` makes unknown populations explicit. Sheet scopes
with `sourceSheet=false` are workbook-level or external graph scopes, not actual
worksheet names. `headerBands` reports scan/candidate/band counts, rejected
ambiguous candidates, named caps, and completeness of that inference scope.
`index_rebuild_required` identifies missing/unreadable indexes or
metadata without a SHA-256 digest. Reimport the project through the existing
import operation to rebuild it; no workbook index schema migration is required.
Older indexes without merge metadata still support the remaining label methods.

## Generations and budgets

Each owner has a derived catalog directory. The workspace's existing coordinator
file lock enforces one writer process. Import completion and deletion enqueue
coalesced owner rebuilds on the existing bounded task executor. Building holds an
owner source-mutation gate, outside the global task-store lock; unrelated API
requests and task polls remain responsive. Source replacement/deletion acquires
the same gate before the task-store lock. Status exposes queued/building/ready/
failed state, phase and observed scan progress. HTTP search/status detect changed
content hashes and enqueue reconciliation; they never perform a large build
under the global lock. Searches can use a previous compatible generation while
building, marked `stale=true`, incomplete, with `catalog_refresh_pending`.
Deleted/replaced revision evidence is excluded after reauthorization. Before the
first generation (or during a schema upgrade), search returns an empty pending
result with the full trace; this does not mean absence. At a one-connection budget,
pending searches return metadata only and lineage asks the client to retry.
Offline CLI operations reconcile synchronously. A failed rebuild remains visible;
retry with an import/delete event or the offline rebuild command. Readers pin immutable
generation filenames; a writer publishes `CURRENT` atomically. Old generations
are collected only after their readers finish. Workbook indexes are opened
read-only through an operation-local LRU; they are never attached to a database.
Workbook reads are serialized per owner to share the configured connection
budget. The service's task-store lock also prevents replacement during evidence
reads. Interrupted builds leave the old published generation intact. Under the single
writer lock, subsequent builds remove abandoned staging files and orphan
generations before reserving space. Missing/unreadable published artifacts trigger
a rebuild; failed database opens release their reader pin.

`WebConfig(catalog_limits=CatalogLimits(...))` configures named limits, using
the same nonnegative integer convention as analysis limits:

| Limit | Default |
| --- | ---: |
| `entries_per_workbook` | 100,000 distinct text/structural entries |
| `numeric_cells_per_workbook` | 100,000 candidate assessments; indexed numerics cannot exceed this |
| `text_occurrences` | 8 source occurrences per text group |
| `header_scan_cells` | 50,000 source cells per sheet |
| `header_candidate_rows` / `header_columns` | 512 occupied rows / 256 columns per sheet |
| `header_bands` / `header_regions` | 64 bands / 256 table or merge regions per sheet |
| `header_row_labels` | 4 row-label columns per band |
| `catalog_bytes` | 256 MiB |
| `build_seconds` | 120 seconds |
| `connections` | 4 workbook connections |
| `hits` | 500 |
| `response_bytes` | 1 MiB |
| `scan_entries` | 50,000 fallback entries |
| `query_ms` | 2,000 ms |
| `query_chars` | 512 |
| `text_chars` | 4,096 per entry |
| `lineage_nodes` / `lineage_depth` | 200 / 20 |

Connection budget must be at least one and response budget at least 4 KiB.
Invalid requests reject excessive query/paging/depth bounds. Construction and
search caps are reported with counts. The catalog disk budget reserves space
for both the old generation and its replacement; insufficient space prevents
publication and reports a failed build. The owner workspace quota (by default
2 GiB, configurable with `serve --max-storage-mb`) additionally includes original
workbooks, all indexes, tasks and captures. It is not a capacity promise for
hundreds of large workbooks.

Construction reserves structural entries first: all sheet names, then tables,
columns, defined names and accepted header bands. The remaining entry and numeric
assessment budgets are water-filled across sheets; small sheets return unused
entry shares. Body scans run in bounded round-robin batches, so a time cap does
not consume one large sheet before visiting every other sheet. Literal text is
deduplicated by full normalized identity within a sheet; truncation does not merge
different long strings. Retained text groups continue counting repeats after an
entry cap. `entries_per_sheet` reports a computed allocation cap. Numeric candidate
assessments use a deterministic systematic sample spanning each sheet, with exact
unassessed counts; full numeric coverage is not claimed at that cap. Numeric facts
are stored separately so repeated numeric label text uses one search entry.

Conservative `cell_labels` methods remain unchanged, including the dossier.
The catalog additionally builds a bounded `header_band` map once per sheet per
generation, stored only in its rebuildable database. Descriptive declared table
columns take precedence. Inference accepts literal source strings, never formula
caches; numeric-looking text headers require at least two descriptive anchors and
a consecutive literal integer axis of at least three columns. A coherent following
data row must corroborate the header. A single-column band requires two literal
numeric rows. Accepted columns propagate only through observed contiguous data;
a blank/missing row, ragged rectangle, textual section, next header or merge ends
the band. Inference never extrapolates past its scan cap. Overlapping/merged
candidate headers and excess metadata are rejected with ambiguity/budget counts.
Every new label gives its exact source sheet/cell, `method=header_band`, role,
header row and observed region; its role is a heuristic, not verified meaning.
Precision is preferred to filling every cell: unfamiliar layouts, numeric typed
headers, instruction rows, mixed regions and long uncaptured columns may remain
unlabelled. Observed counts and caps expose these remaining limits.

## CLI

Offline commands acquire the same workspace writer lock. Stop the local server
before using them; an active coordinator is reported explicitly. The default
owner is the persisted local profile, identical to `linexcel serve`. Embedded
workspaces may select an owner hash explicitly through the trusted local CLI.

```powershell
linexcel library --data-dir ~/.linexcel status
linexcel library --data-dir ~/.linexcel rebuild
linexcel library --data-dir ~/.linexcel search "chiffre affaires" --kinds cell
linexcel library --data-dir ~/.linexcel search 123 --project PROJECT_ID
linexcel library --data-dir ~/.linexcel lineage HANDLE --depth 2
```

Catalog command options include `--entries-per-workbook`, `--catalog-mb`,
`--build-seconds`, `--connections`, `--hits` and `--response-kb` before the
subcommand. Output is JSON; exit 0 means complete within the stated policy,
3 means bounded/incomplete coverage, and 2 means an invalid request or failed
operation. Neighborhood lineage always exits 3 because full lineage is outside
its scope. HTTP and CLI use compact UTF-8 JSON; `response_bytes` bounds that
actual JSON payload, including failed searches and status. CLI adds no newline
or pretty-print whitespace. Metadata omitted to fit the payload is counted as
`revisionsOmitted`, with incomplete coverage and a `response_bytes` cap.
No AI configuration, provider, network request or dependency is used.
