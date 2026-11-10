# Search, follow links and ask questions

Use **Search** to search the loaded workbook library. The search field is ready
immediately; **Advanced search** reveals scope and object-type filters. Use
**Back to workbook** to return to the application workspace. Search and lineage
work without an AI provider. Import the workbooks first, then search their labels,
sheet names, names, table columns or labelled numbers. Accents and case are folded;
all search terms must match. Try a short label if a sentence finds nothing.
Type a number as the sheet shows it: `967,5` or `967.5` finds a saved 967.538,
and `1 455 900` finds 1455900. A cell address such as `B5` or `Synthèse!B5`
finds that cell when no text matches it.
Check the scope, revision, coordinates and coverage beside each result. A partial
search cannot establish absence.

Select a source cell to follow its upstream or downstream lineage. External A1
links cross into another loaded workbook only when its name identifies one
revision, or when you explicitly pin a revision in **External links**. Same-name
files and revisions remain ambiguous. Never assume that the most recently loaded
file is the intended one. A pin becomes stale after replacement. Saved link
caches are historical evidence; cache agreement never establishes identity.

Read the boundaries as carefully as the graph. Missing files, ambiguity, unsupported
references, macros, Power Query, cycles and budgets can stop traversal. A complete
supported lineage describes loaded evidence, not every source that ever existed.
Large graphs need zoom; the relationship list is an accessible alternative.

When a provider is enabled, **Ask AI a question** appears beside direct search.
It carries your search text into an editable question; opening it sends nothing
to a model. **Back to search** returns to direct search. Without AI configuration,
direct search remains available and the AI action is hidden.

Choose a workbook in **Authorized question sources** to constrain the question.
The existing revision and attachment scopes remain authoritative. A literal
workbook filename in the question also narrows the scope before any model call:
for example, `Use Budget FY26.xlsx only`. Case and Unicode composition are
normalized; accents remain significant. Multiple named files form an explicit
allowlist. An unknown filename, a conflicting selected source or a duplicate
filename stops preparation and asks you to choose a source. The resolver matches
literal filenames; it does not interpret natural-language exclusions or infer
which unnamed workbook you meant. Use the selector for those cases.

The question uses the same sources and the chosen scope. The interface discloses
the configured endpoint, model, question, scope
and categories of workbook data. Remote processing requires consent for each
question. All AI clients refuse HTTP redirects: configure the final endpoint URL.
The model cannot load files, choose a pin, execute code or widen the scope. Attach
an additional workbook yourself if needed; question attachments remain scoped to
that question unless you add them to the library.

The model follows the application's workbook-evidence instructions and calls
bounded tools: catalog and value search, small range reads, source-cell checks,
lineage traversal and arithmetic. It must submit source handles rather than
invent an answer. **How AI searches** explains this workflow in the interface;
the trace records the actual operations. This does not require an external skill
installation. The configured model must support structured tool calls; a
text-only response fails explicitly.

A **Proposed answer** contains server-read cells and server-computed lineage. The
saved value and formula can be correct while the selected cell answers the wrong
question. Relevance remains yours to confirm or reject. Labels are often inferred.
No workbook formula is recalculated by the question tool. Arithmetic uses the
bounded `compute` tool and produces a synthetic result with all operands, units
and periods still requiring review. A model rationale is unverified, even if it
contains no digits. A model that misnames a sheet, asks for too large a range or
proposes a range instead of cells is told so and may correct itself, up to three
times; the trace shows each refused call. Any other invalid call, or a provider
failure, ends the search without a proposal.

The model does not receive whole tool results. It gets the cell identities,
labels, saved values and formulas it needs to choose its next step, without
revision hashes, file-format details or the folder an external link pointed to
on its author's machine. A long lineage is cut to its nearest cells for the
model, and the trace says so. The answer you read is not shortened: linexcel
rereads every cited cell and its full lineage from the indexes.

Every proposed cell shows its directed dependency relationships, saved values,
formulas, provenance, origins and traversal limits beside the answer. Arrows run
from a source to the cell that uses it. This lineage is recomputed by linexcel
from its database and graph, independently of the model's search steps. The
operation trace describes how the model searched; the dependency proof describes
where the selected value comes from. Neither establishes relevance.

For a new calculation, **Calculation code and sources** shows the typed
operation, verified operand snapshots, a versioned Python recipe and its digest.
Linexcel authors this recipe and computes the result; the model supplies only a
tool call selecting an allowed operation and previously returned sources.
The recipe calls the same allowlisted arithmetic implementation and can be
replayed with `linexcel.ask_tools.replay_computation`. Replay uses the recorded
saved-value snapshots; it does not reopen or authorize current workbooks.
No arbitrary model-generated code is executed. An optional AI comment remains
separate from the source facts and calculation proof.

The operation trace is always available, including on failure or abstention. It
shows counts, budgets, durations, finish reasons and omissions. When a question
stops on its token, step, time or result-size limit, the answer area names that
limit. Such a stop never establishes that the information is absent. A server
started with `linexcel serve` gives each question 64,000 cumulative tokens by
default. Every tool step resends the history. Change the allowance with
`--ai-question-token-budget`. A request that would pass the allowance is not
sent. Right after a large import the library rebuilds its search index; a
question asked then says it is waiting and starts when the rebuild ends. A
completed search
is not proof that no answer exists. Stored proposals are checked against current
source content and freshly traversed link state on read. An unrelated pin or an
equivalent catalog rebuild preserves a valid proposal; relevant changes suppress
its historical completeness claim.

For provider limits, HTTP contracts and retention see [Library reference](../library.md).
For optional local tracing see [Phoenix](../phoenix.md). Model quality must be
measured against frozen gold evidence, with wrong cells, failures and unanswerable
questions retained in the denominator; source fact verification is a separate test.
