---
name: check-group-precedents
description: >-
  Exhaustive bounded check that a group of stretched formulas is wired to
  exactly the cells its formulas read. Use before and after changing
  `src/linexcel/refs.py` (`stretch_ref`, `parse_ref_detailed`, `ref_to_r1c1`,
  `cell_to_r1c1`), `canonical_r1c1` in `rewrite.py`, `FormulaGroup` in
  `sweep.py`, or how `build_formula_nodes` in `graph.py` turns a reference into
  edges; when a report shows a formula group with missing or surprising
  precedents; and when asked to verify, prove or model-check lineage grouping.
---

Read and follow the agent-neutral skill at `.agents/skills/check-group-precedents/SKILL.md`
(repository root). It is the single maintained copy, shared with Codex, Copilot
and other coding agents; its scripts and the record of known findings are in
the `scripts/` folder beside it.
