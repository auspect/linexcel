# linexcel v1.8 — swarm SPEC (worktree-shared)

Target: make the lineage **viewer** clean/professional, make the **value display
(file vs calculated) provably coherent**, and improve the **terminal** UX with
English recommendations + rich output. UI language in the viewer stays **English
only** (i18n keys exist per-language; viewer default en). All new terminal text
is **English**.

## Disjoint work axes (merge-safe)

### AXIS A — viewer.html (OWNER: kimi builder, single editor)
File: `src/linexcel/assets/viewer.html` (~2046 lines, CSS+DOM+JS inline) + its
tests `tests/test_viewer_ui.py`, `tests/test_viewer_values.py`.
Goals (keep the contract — see §Contract):
1. Professional finish: spacing, hover/focus states, consistent chrome,
   legibility at dense graphs, empty/error states per panel.
2. Value coherence: the file-vs-calculated card (`readings()` l.1824, `valueSection`
   l.1870, `pairGrid`/`sampleTable` l.1916/1940) must show coherent, truthful states:
   - "not stored" vs "not recalculated" vs "guarded fallback" vs "volatile" are
     distinct, correctly attributed, never an empty cell that reads as a bug.
   - verdict (`same`/`differ`/`format`) is decided in Python (`cachedAgreement`) and
     only displayed here — do not re-decide in JS.
3. Screenshots: when a report has no sheet screenshots (technical: LibreOffice
   missing / render failed / `--screenshots` not passed), say so plainly in the
   viewer instead of an empty-looking gallery, and point to the CLI flag. Promote
   the screenshot feature when screenshots ARE present (it helps trust).

### AXIS B — terminal / CLI (OWNER: kimi cli)
Files: `src/linexcel/cli.py`, `src/linexcel/progress.py`, `src/linexcel/structure.py`,
`src/linexcel/result.py` (only `save_screenshots`/render paths), tests that exercise
CLI output. Disjoint from AXIS A (no viewer.html).
Goals:
1. English recommendations to the user at the end of a run (stderr), e.g.:
   - if `--screenshots` was NOT passed and the file would benefit → suggest it.
   - if `--screenshots` WAS passed but render failed / LibreOffice missing → say
     why and how to install/enable, don't silently drop shots.
   - if external workbooks are unread → suggest `--refs-dir DIR`.
   - reuse existing recommendation strings already present (dry-run has good ones).
2. Rich print: add `rich` as an OPTIONAL, opt-in pretty summary (list of nodes by
   kind / counts / a table of top-level values). Must degrade to plain text when
   rich is not installed. Do not force rich on the default path. (rich is currently
   avoided; ship it opt-in via `linexcel[progress]` extra — already declared.)
3. Keep determinism: rich must never change the HTML/JSON output, only stderr.

### AXIS C — research consult (NO code)
Read `~/excel-skill` (the "spécialisé Excel skill" repo — veracity contracts:
provenance, hors-graphe XML, divergence invariants) and the iOfficeAI/OfficeCLI
repo (github.com/iOfficeAI/OfficeCLI — Office suite CLI for AI agents, schema-driven,
SKILL.md embedded). Write `docs/swarm_v1.8_research.md`: what transfers to linexcel's
viewer value-verdict model / CLI (provable contracts, provenance, divergence) and
what OfficeCLI offers linexcel could borrow. Recommendation only; no source edits.

## §Contract (ALL axes must not break)
- Public API unchanged: `render_html`, `analyze`, `result.save_html/to_html/node/find/document`, graph-JSON structure.
- graph JSON: additive-only changes. New fields ride along (`valueSource/cachedValue/valueDate/cachedAgreement` already exist).
- The 49 viewer tests + value tests ARE the contract (chrome + value semantics). `uv run pytest tests/test_viewer_ui.py tests/test_viewer_values.py` must stay green; full `uv run pytest` at the end.
- Viewer template is ONE file. Do NOT split into separate .css/.js — `viewer.py` reads the single template and `test_viewer_ui.py` parses its script block.
- i18n: any NEW user-visible viewer string must be added in ALL 9 languages (`src/linexcel/i18n.py`); keep each line ≤88 chars (E501). Let `uv run ruff format` decide CJK joins; confirm `ruff check` + `ruff format --check` both pass.
- Lint gate = CI scope: `uv run ruff check src/ tests/` AND `uv run ruff format --check src/ tests/` (both, not just src/).
- Verify with `uv run linexcel`, NEVER `uvx` (uvx pulls the published package, not this branch).
