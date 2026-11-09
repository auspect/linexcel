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
- `GET /api/library/lineage?handle=lx1.ENCODED_HANDLE&direction=upstream&depth=10`
- `GET /api/library/links?status=ambiguous&project=PROJECT_ID&limit=50&offset=0`
- `POST /api/library/pins` with `{linkId,targetProject,targetSha256}`
- `DELETE /api/library/pins/LINK_ID`
- `GET /api/library/occurrences?generation=GENERATION&group=ENTRY_ID&after=0&limit=50`
- `POST /api/library/attachments/SCOPE_ID/keep`
- `DELETE /api/library/attachments/SCOPE_ID`

Upload requests may set `searchScoped: true` when starting an existing
`/api/uploads` transfer. The returned project remains absent from the normal
project list and catalog until the user calls the keep route. Its opaque upload
ID is the scope ID. Supply `scope=SCOPE_ID` to search, status, lineage, links
and occurrence requests; this authorizes that one attachment alongside the
owner's ordinary loaded revisions, so external links can still be resolved
without exposing the attachment to library-wide searches. Scope pin requests
include `scope` in the JSON body; scope unpin requests include it in the query.
Deleting the scope removes the temporary project. A kept project becomes an
ordinary library member and no longer needs `scope`.

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

Every grouped cell result includes `occurrenceGroup={generation,id}`. Supply those
fields to `occurrences` to reach cells omitted from the inline list. The lookup
rechecks owner, project and source revision, then scans the authoritative index in
source sequence order. It matches full normalized text, or the same label text
using the generation's bounded header map and current source row context. The
result contains addresses and reusable cell handles. It can include matching
cells outside the catalog's assessment sample; its scope is explicitly separate
from `occurrenceCount`. A missing/foreign/replaced group returns 404; a changed
generation rejects the page. No request supplies a path or chooses an owner.

Use `pagination.nextAfter` as `after` on the next request with the same generation
and group; null means the scan reached the end. Each page assesses at most
`scan_entries` source cells and returns at most `limit` (bounded by `hits`).
`query_ms` interrupts both SQLite scanning and Python processing; `response_bytes`
bounds the entire UTF-8 envelope. Named caps and `coverage.complete=false` expose
an unfinished scan. `page_limit` identifies an ordinary page boundary. A page may
contain no matches and still advance its cursor. If the budget cannot examine
even one source cell, the request fails explicitly rather than returning a
non-advancing cursor. Complete applies to the remaining scan after `after`, not
the whole workbook or global lineage. No full occurrence membership table is
stored, so large groups do not add unbounded catalog disk usage.

Handles are opaque locators, not authorization tokens. Each use rechecks owner,
project, current content hash and node existence. A stale or foreign handle
returns 404. Without `direction`, the lineage route preserves M1's undirected
same-workbook neighborhood envelope. With `direction=upstream|downstream`, it
returns directed lineage across loaded owner revisions, as described below.

## External links and directed lineage

Index schema 4 retains external reference facts independently of calculation:
the `[N]` workbook relationship position, relationship Target exactly as written
(including URI escaping), decoded basename, target sheet/address/rectangle and
any cached values. The immutable catalog generation derives its link table from
those facts. Old indexes need reimporting: `external_facts_rebuild_required` in
link status and `index_rebuild_required` in lineage disclose missing facts.
External A1 cells and rectangular ranges are supported. External defined names,
structured references, whole rows/columns, dynamic and 3-D references are typed
`unsupported_reference` boundaries, with the original operand and subtype.
Local structured columns such as `Sales[2020]` or `Sales[inputs.xlsx]` remain
local: a bracketed column name alone is never an external workbook marker.
Declared but unused external books remain in source index `externalBooks`
metadata; `links` lists formula/name reference operands.

Resolution uses a case-insensitive basename comparison (`lower()`, as in the
existing external resolver), across all currently loaded revisions of the
authenticated owner. Exactly one candidate is `resolved_by_name`, several are
`ambiguous`, and none is `unresolved`. Every same-name project is a candidate,
including projects with identical bytes. Import dates, loading order and cached
values never select a target. `resolved_by_name` explicitly warns that this
may not be the revision the link was saved against. A moved file with the same
basename can resolve; a renamed file needs a pin.

Pins are explicit user actions binding a stable link ID to a project and source
content hash. Both source link and target revision are authorized against the
current owner. Pins persist in the owner's `link-pins.sqlite`, outside the
disposable catalog directory. Rebuilds and process restarts retain them. A
replacement or deletion leaves `status=pinned`, `stalePin=true` and
`targetRevision=null`; it never silently falls back to a same-name candidate.
Unpinning returns to name resolution. Pins for deleted source revisions remain
durable but do not appear as current links.

Directed responses contain `nodes`, `edges`, `boundaries`, `revisions`,
`stronglyConnectedComponents`, `coverage` and `trace`. Node IDs and handles are
revision-qualified; `sourceNode` retains the source index ID. Each node includes
sheet, address, kind, labels, saved value/provenance and formula. Range nodes
describe stored members and implicit blank cells, without enumerating every
blank. Their handles and verified blank-cell handles can be traversed again.
This reads saved values and source formulas; it does not recalculate formulas.

Edges always point from precedent to consumer, regardless of traversal direction.
Kinds include `formula`, `range` and `external_link`; the latter includes
`resolutionStatus`, original path and version caveat. Upstream resolves into the
target revision and follows its own formulas. Downstream searches the generation's
reverse link table, including range containment, plus local ranges. An ambiguous
potential reader creates a boundary, never a fabricated consumer edge.

Downstream also checks every identity-unresolved link in the loaded owner's
generation, including missing targets, ambiguous targets and stale pins. A
different filename cannot exclude a reader: the target may have been renamed.
A verified A1 reference can be excluded when its sheet differs from the selected
cell's sheet or its rectangle does not overlap the selected cell/range.
Unsupported forms without verified bounds remain possible readers.

Each reached cell/range aggregates the remaining possibilities in a
`possible_unresolved_reader` boundary with an exact `count`, a deterministic
`linkIds` sample, and `sampleOmitted`. The count describes link operands, not
distinct workbooks. `lineage_reader_samples` bounds the sample (default 8);
zero retains the count with an empty sample. Known ambiguous candidates can also
have an `ambiguous_external` detail boundary. These possibilities never create
edges or select a target revision, and always make `coverage.complete=false`.
An exhaustive count can still have `coverage.exhaustive=true`; a time or other
traversal cap makes it false. A time cap during sampling preserves any completed
count and records omitted samples.

Upstream external edges include `cacheDiagnostic` entries such as
`{address:"A1",cached:500,current:600,agreement:false}`. Numeric comparisons reuse
the existing relative/absolute tolerance of 1e-9. Missing current values yield
unknown agreement. Diagnostics are bounded and report omissions; they never
change resolution. An unresolved link with cache data reaches `cache_only`,
labelled "cached value, source unavailable", with no target edge.

Every boundary identifies its reached node and typed reason. Reasons include
`input_constant` (a legitimate origin), `unresolved_external`,
`ambiguous_external`, `unsupported_reference`, `source_unavailable`, `macro`,
`power_query`, `cycle`, `cache_only`, `source_index_incomplete`,
`possible_unresolved_reader`, and budget stops.
Power Query loaded table destinations are detected through the existing OOXML
connection reader; their cells stop at `power_query`, not an input constant.
VBA procedure nodes stop at `macro`; no macro or query is executed. Iterative
strongly connected component detection identifies local and cross-workbook cycles.

`coverage.exhaustive` says whether traversal exhausted its defined loaded-revision
scope without a budget/missing-index stop. `coverage.boundaries` separately names
what it reached. `coverage.complete` additionally requires all reached boundaries
to be legitimate input origins: an exhaustive graph with an unresolved external
link, macro, query or cycle is still incomplete. This makes a bounded or unresolved
diagram unsuitable for claiming total lineage. The scope cannot identify files
that have never been loaded or prove the historical version saved against a link.
Both `coverage.scope` and `trace.scope` explicitly state: "loaded owner revisions
only; never-loaded workbooks are not visible". Downstream completeness therefore
does not establish that no never-loaded workbook reads the selected cell.

`trace.schemaVersion=linexcel.lineage.v1` records root, direction, generation,
crossed revisions, limits, counts, peak source connections, completeness and
duration. Crossed revisions remain in the trace when response trimming removes
graph nodes; an envelope that cannot retain that audit metadata fails explicitly.
`preparationMs` separately reports synchronous catalog reconciliation;
it is outside traversal's time allowance. Traversal defaults are 200 nodes,
1,000 edges, 32 workbooks, depth 20 and 2,000 ms; the existing response budget is
1 MiB. Request parameters `limit`, `edges`, `workbooks`, `depth`, `milliseconds`
can lower the configured ceilings. Zero edges/time/depth are real limits; root
node/workbook capacity must be at least one. Budget boundaries include `depth`,
`node_budget`, `edge_budget`, `workbook_budget`, `time_budget`, `response_bytes`,
`link_budget` and `link_build_time`. The generation link budget is 100,000 operands
and shares catalog build time/disk limits. SQLite progress callbacks interrupt
long scans; Python loops check deadlines. These are cooperative limits, not
hard real-time guarantees. Oversized root evidence fails explicitly.

Traversal reuses the operation-local bounded LRU and reserves a connection for
background building. It never attaches indexes to SQLite. Immutable generations
are pinned while read; builds/pin changes serialize with the existing owner
mutation gate. Added, replaced or deleted revisions and changed pins invalidate
the generation. `status.links` exposes counts by resolution status and separate
link coverage. Search coverage continues to describe the search index policy.

```powershell
linexcel library --data-dir PATH links --status ambiguous
linexcel library --data-dir PATH pin LINK_ID TARGET_PROJECT TARGET_SHA256
linexcel library --data-dir PATH unpin LINK_ID
linexcel library --data-dir PATH lineage HANDLE --direction upstream --depth 10
linexcel library --data-dir PATH lineage HANDLE --direction downstream --workbooks 8
```

The external-link contract tests generate OOXML externalLink parts and
relationships with openpyxl, with hand-computed graphs. No Excel-authored
external-link file is claimed as validation; the existing Excel-authored VBA
and Power Query fixtures exercise those distinct stop boundaries.

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
Older indexes without merge metadata retain conservative labels; new bands are
disabled with `header_merge_metadata_unavailable` until reimport restores that
geometry. Missing metadata must not be interpreted as proof of no merges.

## Generations and budgets

Each owner has a derived catalog directory. The workspace's existing coordinator
file lock enforces one writer process. Import completion and deletion enqueue
coalesced owner rebuilds on a dedicated single-worker catalog executor, preserving
all workbook task slots for imports and other tasks. Building holds an
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
retry with an import/delete event or the offline rebuild command.
Source enumeration and snapshot exceptions also terminate in a failed build with
an incomplete coverage reason (`catalog_source_snapshot_failed` for enumeration).
Queued reconciliation skips enumeration and building after the store closes;
shutdown cancels remaining queued catalog futures and waits for the bounded active build.
Readers pin immutable
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
| `lineage_reader_samples` | 8 link IDs per possible-reader boundary |

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

The dossier's `cell_labels` methods remain unchanged. The catalog applies a
stricter policy: nearby `above_text` alone is omitted; a column label requires
declared table metadata or a corroborated header band. Within a band, its inferred
roles supersede nearby `left_text`. Declared table nearby text is `row_context`,
not a proven key. The catalog builds a bounded `header_band` map once per sheet per
generation, stored only in its rebuildable database. Descriptive declared table
columns take precedence. Inference accepts literal source strings, never formula
caches; numeric-looking text headers require at least two descriptive anchors and
a consecutive literal integer axis of at least three columns. A coherent following
data row must corroborate at least two adjacent populated header columns.
A single-column title above repeated numbers never creates a band; a declared
single-column table can still supply its source column metadata. Accepted columns
propagate only through observed contiguous data;
a blank/missing row, ragged rectangle, textual section, next header or merge ends
the band. Inference never extrapolates past its scan cap. Overlapping/merged
candidate headers and excess metadata are rejected with ambiguity/budget counts.
Every new label gives its exact source sheet/cell, `method=header_band`, role,
header row and observed region, and `certainty="heuristic"`; it never asserts
`ambiguous=false`. Only the leftmost textual key column receives `role="row"`.
Other textual columns before the numeric data receive `role="row_context"`.
`field` is that column's exact literal heading and `fieldCell` its source address
on the same sheet, so context remains identifiable without treating it as a key.
For example, ONS retains Code as the row key and Name/Geography as context;
`estimated` in an ID/Comment/Amount band is context, not a row key. These roles
describe geometry and remain heuristic, not verified business meaning.
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
linexcel library --data-dir ~/.linexcel occurrences ENTRY_ID --generation GENERATION --after 0 --limit 50
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
