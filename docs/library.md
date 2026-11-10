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

## Optional questions

Before preparation, workbook/revision selections are intersected with literal
workbook filenames in the question. Matching uses NFC and case folding, exact
filename boundaries and longest-name precedence. Quoted and Excel-bracketed
filenames are resolved as whole names before recognizing known substrings.
Unknown quoted names and bare
filename tokens, duplicate names and conflicts fail with HTTP 409 before creating
a token or task. A structured project/revision choice disambiguates same-name
files. No natural-language intent classifier or model chooses this boundary.
An explicit or named workbook is resolved from the owner-local inventory before
opening its index; other workbook indexes are not opened by preparation. The
frozen scope and preparation response expose
`sourceSelection: {kind: "explicit"|"question_filename"|"library", filenames}`.
Only these revisions enter the tools' private catalog and directed graph;
unloaded or excluded external sources remain lineage boundaries.

Direct search works with no AI configuration. **Recherche avancée** progressively
reveals scope and object-type filters. The adjacent **Poser une question à l’IA**
action appears only when the server's existing AI configuration and the `ai`
extra are available. It carries the direct query into an editable question and
does not contact the model until submission. **Retour à la recherche** restores
direct search; **Retour au classeur** restores the application workspace.
The endpoint and model are named in the question view. The separate `ask_provider.ToolProvider`
protocol returns structured tool calls; it does not implement `aidoc.LLMProvider`.
Its OpenAI-compatible adapter reuses the documentation client's HTTP connection
setup and M0 parameter negotiation. A text-only or tool-rejecting endpoint fails
explicitly, with a trace and no prose fallback.

The application's evidence-search instructions and strict tool schemas guide
the model; no externally installed skill is needed. The question interface
explains this under **Comment l’IA recherche**. Opening this view, inspecting
configuration and deterministic search do not themselves make a model request.

Questions are prepared through `POST /api/library/ask/prepare` with
`{question, project?, revision?, scope?}`. The response discloses the exact
loaded revisions, question, endpoint, model and categories (values, formulas,
labels, sheet/workbook names). The 10-minute, owner-bound token freezes this
scope. `POST /api/library/ask` consumes `{token, remoteConsent:false}` locally;
a remote endpoint requires `remoteConsent:true` for that individual question.
Only literal loopback IPs and `localhost` count as local. A changed endpoint,
expired token or reused token requires fresh preparation. No model request is
made before required consent. An empty/unimported scope fails clearly; scopes
above 32 revisions or 64 KiB of revision metadata require narrower selection.
Question HTTP clients never follow redirects, including from loopback to remote
destinations. The approved endpoint and model are frozen in task input; the
provider uses that snapshot and rechecks live configuration before each request.
A queue-time or later configuration change fails closed without sending to the
new destination or silently choosing a different model.

`GET /api/library/ask/config` exposes availability and retention. A started run
is an ordinary `ask` task on the existing executor, capacity queue and cancel
event. `GET /api/library/ask/RUN_ID` returns task progress and its authorized run;
`POST /api/tasks/RUN_ID/cancel` cancels it. Model calls occur outside
`TaskStore.lock`. Cancellation is cooperative: an in-flight HTTP call can take
up to the smaller of 30 seconds and the remaining run time to settle. Queued
cancellation and interrupted-server recovery also persist terminal traces.
Question results are intentionally available only from the ask route, where
source authorization is rechecked; generic task polling exposes task metadata.
Completed proposals retain a stable content signature of their source cells and
traversed lineage. Result reads recompute that lineage under the owner mutation
gate. Relevant link changes suppress stale completeness verdicts; unrelated pins
and equivalent catalog rebuilds preserve valid proposals.

The explicit state machine validates tool names, strict argument objects,
finite numbers, call IDs (unique across the run), response finish reasons and
terminal-call placement. There is no code, path, URL, SQL, file, network, pin or
upload tool. Workbook text is inside JSON marked `untrusted_workbook_data`;
all text is rendered inertly. Every tool reacquires the owner mutation gate,
re-enumerates authorized current sources and checks frozen content hashes.
Newly loaded workbooks cannot enter an existing run. Pins and file loading remain
user actions, and a missing source is a lineage boundary.

Tools are `list_workbooks`, `list_revisions`, `search_catalog`, `describe_sheet`,
`read_range` (at most 25 indexed cells), `get_cell`, `find_value`, `trace_lineage`,
`list_external_links`, `compute` and `submit_answer`. Search/links return at most
20 results, sheet descriptions sample 12 cells, and numeric lookup assesses at
most 5,000 source cells/20 hits with explicit coverage. Range omissions are
reported as blank-or-unindexed; they are not invented evidence. Search and
directed lineage call the M1/M2 engines in a private, disposable run catalog,
so a narrow question cannot replace the library's shared catalog generation.
The private catalog reads the same durable owner-local link pins as the library;
the model cannot create or change them. User pin changes invalidate its derived
generation and remain constrained by the question's frozen authorized revisions.
Its disk ceiling is 16 MiB, build allowance 15 seconds, additionally bounded by
the configured catalog limits. These are cooperative limits, not hard real-time
or process-memory isolation guarantees.

`submit_answer` accepts `{status:"proposed", handles?, compute?, rationale?}`
or `{status:"not_found"}`. Every proposed cell and computation operand must have
been returned to the model by a tool in this run. The server rereads its value,
formula and labels, checks index existence and computes upstream M2 lineage.
Outcomes are `answer_proposed`, `not_found`, `failed`. No valid handle or an
exhausted search budget gives `not_found`, displaying only the trace. Provider,
call-validation and cancellation failures terminate explicitly. The trace's last
step states whether operations completed or were truncated and that this does
not establish absence. A model's decision to stop is not an exhaustive search
of every cell or proof of relevance.

The **Réponse proposée** card marks source facts as verified while keeping
relevance **à confirmer**, shows source revision/coordinates, heuristic label
provenance, saved values/formulas and source lineage with its boundaries. No
formula is recalculated. A short rationale is shown only as **Commentaire de l’IA,
non vérifié**. Any rationale containing a digit is omitted. Digit-free
text can still contain false claims, including numbers written as words. This
presentation rule does not validate numbers, units, meaning or relevance and is
not a defense against laundering claims. Rationale text is never retained in
trace arguments, including `not_found`. Confirm/reject calls to
`POST /api/library/ask/RUN_ID/feedback` with `{decision:"confirm"|"reject"}`
record up to 20 timestamped decisions without blocking initial display.

`compute` accepts only sum, difference, ratio, min, max, mean and count over
issued operand handles. Difference/ratio require two finite numerics. Sum,
min/max/mean and count ignore explicit blank values, preserve all operands and
labels, propagate source errors, and reject nonnumeric operands. Count counts
numeric operands. An absent saved formula cache is unavailable, not an arithmetic
blank. Results are synthetic nodes labelled **Calculé par linexcel pour cette
réponse, présent dans aucun classeur**. Their operands retain server-computed
source lineage. Arithmetic correctness does not establish that a calculation
answers the question, matches units or uses the right period.

Each calculation includes a `linexcel.compute.v1` receipt: allowlisted engine and
version, operation, canonical operand snapshots (revision, coordinates, saved
value, formula and provenance), executable Python and `codeSha256`. The server
authors this recipe, never the model. `replay_computation` validates the version,
operation and recipe/snapshot consistency and reproduces the same arithmetic,
including blank, source-error and missing-cache outcomes. Production computes
through the typed function; it never evaluates model code. An offline replay
uses historical snapshots and does not confer current source authorization.

The model transport replaces canonical cell handles with short run-local IDs
such as `c1`. Only evidence actually delivered by a successful tool call creates
these capabilities. Incoming aliases must match exactly; they resolve to the
canonical handles before authorization and computation. Workbook text is never
rewritten as an alias. Published answers, graphs, receipts and evidence signatures
retain canonical revision/cell identities; aliases are not persisted or shared
between questions.

Each proposed cell has its own visible source graph, relationships and traversal
boundaries. Source-to-reader arrows follow the stored graph direction. The
model's operation trace is distinct from this server-computed dependency proof.

`WebConfig(ask_limits=AskLimits(...))` configures default ceilings: 16 tool steps,
64,000 cumulative tokens, 120 seconds, 512 KiB of tool results delivered to the
model, 2,048 output tokens per call and 30 seconds per HTTP call.
`linexcel serve --ai-question-token-budget` sets the token ceiling. Each request
is estimated before it is sent, at no less than one token per three bytes of
context and tool schemas, then corrected by the provider's count for the
previous request. A request that would pass the ceiling is not made. Observed
or estimated provider usage includes parameter-negotiation attempts. Each trace
is bounded to 256 KiB, with argument/handle-detail omissions disclosed if
necessary. Provider prose and reasoning streams are neither answers nor trace
content.

The model receives a compact view of each tool result, at most 24 KiB: cell
identity, sheet, address, kind, saved value, formula, up to four labels and the
value's provenance. Revision hashes, XML provenance, per-sheet catalog coverage
and the path an external link had on its author's machine are not sent. Text is
cut at 300 characters, a formula at 600 and a lineage at its 60 nearest nodes.
A shortened result states what was left out, the trace records it, and the
search is then reported as incomplete. Only handles in the delivered view
become capabilities. The verified answer is read again in full from the source
indexes. It never returns to the model, so it does not count against the
512 KiB; it is refused above 8 MiB.

A call that names an unknown sheet, asks for more than 25 cells, or proposes
something other than a cell is answered with a fixed code and hint, and the
model may correct it. The fourth such call ends the search. Every other invalid
call still ends it at once: an unknown tool or argument, malformed JSON, a
project outside the scope, or an unissued handle in a read, lineage or
calculation.

A catalog rebuild after an import holds the owner's library gate, for up to two
minutes on a large workbook. Preparing a question does not wait for it. A
submitted question waits before its 120 seconds start, and reading the run
meanwhile returns `libraryBusy: true` instead of blocking.

Runs persist under the private owner's `ask/RUN_ID/`, separately from source
indexes. `linexcel.ask.v1` records scope/revisions, tool and argument metadata, bounded
result counts/handles, duration, tokens, budget use and verdict. Cell values,
formulas and labels are excluded from default step summaries.
Tool argument values and call IDs are never retained in the question trace.
Only argument byte counts and schema field names remain; no value-bearing digest
is stored. Metadata is not anonymization.
Input questions and authorized factual answers remain separate private artifacts.
Do not export these private artifacts by default. Retention defaults to seven days and 100
settled runs per owner; preparation and reads sweep expired/excess runs, excluding
active tasks. `AskLimits.retention_days` and `retained_runs` configure it. Reading
a run after source replacement/deletion or attachment expiry suppresses its
evidence and old scope details. Trace publication is application-owned, with no
global tracer. See [optional content-free Phoenix export](phoenix.md).

## Deterministic library routes

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
- `GET /api/library/attachments`
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
`GET /api/library/attachments` lists recoverable temporary projects after a
reload. Each has an opaque scope ID and an expiry time. Temporary projects
expire after 24 hours; scoped reads reject them and keeping an expired scope
returns HTTP 410. Listing attachments and starting an upload advance a bounded
owner-local cleanup sweep (32 manifests per call), freeing expired projects and
their files before upload quota admission. Repeated calls continue the sweep;
active tasks are cancelled and retried by a later sweep. Kept projects have no
scope ID or attachment expiry and are never candidates for this TTL cleanup.
Deleting a scope removes the temporary project, including after an
import task failed. A kept project becomes an ordinary library member and no
longer needs `scope`.

The no-AI view binds results and trace steps to each submitted scope. Superseded
requests are aborted and marked as superseded, rather than completed; their
responses cannot replace the current evidence. The persistent trace control is
before the search controls, including on mobile, and exposes the full audit.

The shared graph starts with a collapsed workbook/revision overview. Each card
represents its listed member cells and counts internal relationships; expanding
a revision shows those cells while other revisions stay aggregated as context.
Cross-revision dependencies retain their direction and aggregate their counts.
Origins and stops are labelled on cards and in the detailed evidence, not added
as synthetic dependency edges. Display counts distinguish drawn relationships,
the source dependencies they represent, and grouped internal relationships;
the API lineage and its coverage remain unchanged by this presentation filter.

`q` and `query` are aliases; supplying both is an error. Kinds are `cell`,
`sheet`, `name`, `table` and `column`. Text search folds case and diacritics and
requires every Unicode word token. FTS5 uses `unicode61 remove_diacritics 2`.
When unavailable, the reported backend is `like`, with matching over a bounded
prefix of catalog entries and a token-boundary check after the LIKE filter.
Underscores are separators in both modes. Punctuation and SQL/FTS syntax are
never executable.
A numeric query matches the labelled cells whose saved number rounds to the
query at the number of decimals typed: `6600` finds a saved
`6600.000000000001` and `967.5` finds `967.538`, while `967.60` does not.
Spaces, apostrophes, dots or commas group thousands; a comma or a dot is the
decimal separator, whichever comes last. `1,234` is ambiguous and matches both
1.234 and 1234. `34.3%` matches 34.3 and 0.343. Other spellings Python reads,
such as `1e6`, match exactly. The trace reports the bounds as `numericMatch`.
Number formats are not applied, and computed values are outside this version.

A query that is one cell address, such as `B5`, `$B$5` or `Synthèse!B5`, is
searched as text first, because `Q4` or `FY26` are labels far more often than
addresses. When no text matches, the result lists the cell at that address in
each scoped sheet, or in the named sheet. Those hits carry `matchedBy:
"address"` and the trace reports `addressMatch`. Empty cells are not listed.

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

The external book's declared `r:id` selects its relationship, including when
Excel writes `xxl21:alternateUrls` and lists alternate relationships first.
Drive-less paths, file URLs, UNC paths, OneDrive and SharePoint URLs use the same
basename rule; URI escapes are decoded for name matching, while
`relationshipPath` preserves the selected Target as written. No target URL is
fetched, and no absolute path is opened to identify a loaded revision.
For HTTP(S) Targets, name matching uses the URL path without its query or fragment;
local paths retain literal question marks. Conflicting declarations of one
relationship ID leave identity unresolved; identical type/Target/mode duplicates
remain valid. Older external-fact indexes require a structure refresh or reimport;
until refreshed, lineage reports `index_rebuild_required` instead of trusting
previously extracted identity facts.

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

A cell's external readers must contain that cell. A local range reached through
one of its members continues only to its exact local consumers; its other
members' external readers and uncertainties are not inherited. A range selected
as the traversal root includes external readers overlapping its full span.
Verified blank-cell handles use the same containment rule as stored cells.
Blank members inside declared array bounds also follow their array anchor. This
is array-group provenance, without inventing a formula or value for the member.
Read-time lookup does not materialize the output rectangle. Downstream expands
implicit members only within the node ceiling; a larger implicit output region
stops with an explicit budget boundary and incomplete coverage.

Downstream also checks every identity-unresolved link in the loaded owner's
generation, including missing targets, ambiguous targets and stale pins. A
different filename cannot exclude a reader: the target may have been renamed.
A verified A1 reference can be excluded when its sheet differs from the selected
cell's sheet or its rectangle does not overlap the selected cell/range.
Unsupported forms without verified bounds remain possible readers.

Each reached cell (or selected range root) aggregates the remaining possibilities
in a `possible_unresolved_reader` boundary with an exact `count`, a deterministic
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

The committed external-link contract tests generate OOXML externalLink parts
and relationships, with hand-computed graphs, including Excel's three-relationship
alternate-URL shape. Local trials on 2026-11-09 authored external-link workbooks
with real Excel 16.0 through COM and through GUI actions, then compared saved
values and lineage. The GUI trial includes single-cell and multi-cell legacy CSE
arrays and dynamic spills. Those workbooks are kept outside Git because they
contain absolute local paths. Excel-authored shared external formulas and live
cloud links remain unverified. The committed
Excel-authored VBA and Power Query fixtures exercise those distinct stop boundaries.

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
Search-scoped builds use the same owner mutation gate and catalog executor and
never reconcile synchronously while holding the task-store lock. The first
scoped search/status response is explicitly pending and incomplete; retry after
the build reaches a terminal state. Scope reconciliation serializes with normal
owner builds because both publish into the owner's immutable catalog.
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
