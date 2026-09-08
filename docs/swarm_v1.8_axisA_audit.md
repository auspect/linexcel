# swarm v1.8 — Axis A: qualitative value-coherence audit (runtime DOM, not just tests)

Method: rendered real reports with `uv run linexcel analyze`, drove the viewer under
headless Chromium (Playwright), selected cells via the address search box, read the
rendered `.lin-vcard` DOM text. Confirms the value card is TRUTHFUL in every state glm's
consult flagged. No source change needed — the model is already coherent.

| Case | Fixture | Rendered card | Verdict |
|---|---|---|---|
| stored≠calc divergence | excel-skill divergence.xlsx S!A3 | file `99` / recalc `3` | `is-diff`, verdict "The recalculated value differs from the file" ✅ |
| IFERROR fallback | iferror_nosheet S!E2 | file "Not stored" / recalc `-1` | `is-guarded`, no false agreement ✅ |
| IFERROR→"absent" | iferror_nosheet S!E3 | recalc "absent" | `is-guarded`, echo not mislabelled as engine ✅ |
| volatile TODAY() | volatiles S!K1 | file `—` / "Not recalculated (volatile)" | viewer abstains honestly, no fake same ✅ |

Also verified: zero JS console errors across all rendered reports; the screenshots tab +
per-sheet gallery render cleanly when `--screenshots` present (no double-render).

## Conclusion
The file-vs-calculated provenance model (verdict decided in Python `cachedAgreement`,
viewer only displays) is coherent. glm's Q1 recommendations (add `same-exact` vs
`same-within-ε`, `differ-volatile`, no-engine-reading classes) are enhancements, not
fixes for existing incoherence — parked as future hardening, not needed for v1.8 coherence.

## Screenshots promotion — already covered by AXIS B (CLI, per Philippe "terminal suffices")
`linexcel analyze` now tells the user when a multi-sheet report lacks screenshots:
`hint: 2 sheets would render as pictures — rerun with --screenshots DIR ... (needs LibreOffice)`
and explains loudly when a requested render fails (LibreOffice missing / pdftoppm), keeping
the report written. Verified live.
