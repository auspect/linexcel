# swarm v1.8 — Research consult (AXIS C): reuse excel-skill & OfficeCLI

Model consulted: **z-ai/glm-5.3-flash** (OpenRouter). Read-only recommendation; no code changed.
Inputs: excel-skill provable-contracts (veracity invariants) + OfficeCLI (iOfficeAI) surface.

## Q1 — Veracity invariants → linexcel viewer value-verdict display

Transfers cleanly (already in place, keep):
- Verdict decided in Python (`cachedAgreement`) and cached on the node — viewer cannot reinterpret; single authority.
- Per-side provenance source (file/engine/volatile/fallback) = "every value carries provenance".
- Explicit `not stored` / `not recalculated` states = the "both readings exist" precondition.

Gaps / risks to fix (mapped to falsifications):
(a) **Float tolerance / date-part coarsening.** If `same` is emitted after silent tolerance or a date-part-only compare, a real divergence can be masked (false coherence). A date cell differing only by time-of-day would show `same`. Fix: make the comparison rule an attribute — `same-exact` vs `same-within-ε` (ε shown), and date-part matches render `same (date-part)`, never bare `same`.
(b) **Fallback & volatile echoed in the "calculated" column.** A guarded fallback (IFERROR branch) or an uncomputed `#SPILL!` fallback shows the *stored value echoed back*, not an engine product. Displaying it under "calculated" is the hors-graphe misattribution falsification. Worse: comparing stored vs a fallback echo compares the same bytes → `same` while the engine computed nothing. Fix: fallback side renders `not recalculated (fallback echo: stored value)`; verdict = `no-engine-reading`, never `same`. Volatile needs `differ-volatile` (NOW()/RAND() drift is intrinsic, not staleness) — else users suppress "differ" globally and unflag real divergences.
(c) **Display without honest provenance.** Every rendered cell — including placeholders, blanks, `0`, error fallback — must carry a provenance label; echoes visually distinct from engine output. Blank vs `""` vs `0` must not render identically.
(d) **Absence-of-divergence claims.** Summary counters must treat `not stored`/`not recalculated`/fallback as "no verdict", never agreement. A no-formula workbook must not render a calculated column at all (no-formula postcondition).

Recommended verdict enum: `same-exact | same-within-ε | format | differ | differ-volatile | no-engine-reading | no-stored`. Mandatory provenance label per rendered value. Counts only over cells where both readings exist.

## Q2 — OfficeCLI borrowings (analysis scope only)

- Ship a **SKILL.md** for linexcel describing read-only usage + verdict/provenance semantics so agent hosts interpret output correctly.
- Publish a **JSON Schema** for graph/JSON output with fixed enums (provenance, verdict, states); the HTML's embedded data validates against it.
- **Deterministic output contract**: same input → byte-identical HTML/JSON (stable ordering, no timestamps) → run-to-run diffing as audit evidence.
- **Runnable examples**: one example workbook + golden output snapshot (regression test + doc in one).
- **Per-format docs**: short page per output format (HTML fields, JSON schema).
- **Single-binary ethos**: linexcel's zero-dependency standalone viewer already matches OfficeCLI's — document it as a stated contract.

Skip: editing, plugin/mutation protocols (out of linexcel's read-only scope).
