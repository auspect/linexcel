# Agent workflow proposals using linexcel

These are proposed complementary skills, not installed integrations. Each skill
should start with a concrete user objective, work on immutable input copies,
retain private evidence outside Git, and distinguish source facts, engine
calculations and AI interpretations. Use `linexcel --help` and each subcommand's
help from the installed version before assembling commands.

| Proposed skill | Trigger and useful result | Tools and evidence | Completion gate |
| --- | --- | --- | --- |
| `audit-workbook` | “Find the fragile areas in this workbook.” Produce a prioritized list of formula errors, unresolved references and structural risks. | `linexcel analyze`, graph export, diagnostic counts and a reproducible sample across sheets. | Every finding names its source cell and evidence; unsupported coverage remains explicit. |
| `explain-kpi` | “Explain how this KPI is calculated.” Trace one output to its inputs and provide a concise business explanation. | Targeted analysis, source formulas, defined names, bounded upstream graph, cached and calculated values. | Reconcile the stated arithmetic with the actual formula; distinguish unknown inputs and decomposition limits. |
| `compare-workbook-versions` | “What changed between these files?” Compare formulas, source values, formatting and dependency relationships. | Hash both inputs, analyze both independently, query source-cell exports and presentation windows. | Report changed addresses and provenance; a changed cache or presentation alone is not proof of a regression. No built-in diff command is assumed. |
| `plan-excel-to-sql` | “Turn this workbook into a maintainable data pipeline.” Produce a table mapping and a calculation migration plan. | Datasette CSV/JSON, complete source-cell tables, formulas and dependency groups. | Validate joins, null/error handling, date epochs and expected outputs on agreed fixtures before proposing deployment. |
| `check-reconciliation` | “Check these explicitly defined totals and balances.” Produce a reproducible control ledger. | User-defined control cells, targeted calculations, source values, tolerances and provenance. | Every control records pass/fail/unknown and the tested version. The result is a scoped control check, not an audit certification. |
| `document-handover` | “Prepare this workbook for a new owner.” Produce an overview and source-linked calculation notes. | Structural analysis, optional AI documentation and screenshots with retained evidence dossiers. | Independently review formulas and image claims; missing captures or truncated responses prevent a complete claim. |
| `investigate-calculation-gap` | “Why does this value differ from Excel?” Isolate a minimal reproducible discrepancy. | Cache versus engine readings, external references, date system, unsupported functions and exact source formulas. | Explain what is established, obtain an agreed independent expected result, and keep unresolved differences explicit. |

Implement `audit-workbook`, `explain-kpi` and `investigate-calculation-gap` first:
they reuse existing analysis and provenance without requiring a new comparison
engine. Version comparison and SQL migration need explicit adapters and numerical
contracts. Their first deliverable should be a reviewed plan and reproducible
fixtures, followed by an implementation when authorized.

A shared skill contract should record the installed linexcel version, workbook
hashes, selected sheet/cell scope, limits, command outcomes and evidence paths.
Avoid silent truncation or fabricated missing values. Sending workbook content to
an AI endpoint, publishing results and changing an operational workbook must
follow the user's chosen destinations and authorization. Keep generated
workbooks, captures, model answers and private reports in ignored local storage.
