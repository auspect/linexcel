# Search, follow links and ask questions

Open **Library search** to search the loaded workbook library. Search and lineage
work without an AI provider. Import the workbooks first, then search their labels,
sheet names, names, table columns or labelled numbers. Accents and case are folded;
all search terms must match. Try a short label if a sentence finds nothing.
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

Choose **Ask a question** only if you want the optional model to search those same
sources. The interface discloses the configured endpoint, model, question, scope
and categories of workbook data. Remote processing requires consent for each
question. All AI clients refuse HTTP redirects: configure the final endpoint URL.
The model cannot load files, choose a pin, execute code or widen the scope. Attach
an additional workbook yourself if needed; question attachments remain scoped to
that question unless you add them to the library.

A **Proposed answer** contains server-read cells and server-computed lineage. The
saved value and formula can be correct while the selected cell answers the wrong
question. Relevance remains yours to confirm or reject. Labels are often inferred.
No workbook formula is recalculated by the question tool. Arithmetic uses the
bounded `compute` tool and produces a synthetic result with all operands, units
and periods still requiring review. A model rationale is unverified, even if it
contains no digits. No proposal is shown after an invalid call or provider failure.

The operation trace is always available, including on failure or abstention. It
shows counts, budgets, durations, finish reasons and omissions. A completed search
is not proof that no answer exists. Stored proposals are checked against current
source content and freshly traversed link state on read. An unrelated pin or an
equivalent catalog rebuild preserves a valid proposal; relevant changes suppress
its historical completeness claim.

For provider limits, HTTP contracts and retention see [Library reference](../library.md).
For optional local tracing see [Phoenix](../phoenix.md). Model quality must be
measured against frozen gold evidence, with wrong cells, failures and unanswerable
questions retained in the denominator; source fact verification is a separate test.
