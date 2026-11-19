---
name: verify-linexcel
description: >-
  Drive the real linexcel from this checkout and capture proof that a behaviour
  works: the `linexcel` command line (analyze, macros, dry-run), the standalone
  HTML lineage report it writes, and the local web application started by
  `linexcel serve`. Use after changing analysis, the report viewer, the app or
  the CLI, when asked to "check it in the real app", to reproduce a bug as a
  user would, or to attach screenshots and transcripts to a PR. It launches a
  disposable instance, never the user's own workspace.
---

Read and follow the agent-neutral skill at `.agents/skills/verify-linexcel/SKILL.md`
(repository root). It is the single maintained copy, shared with Codex, Copilot
and other coding agents; its helper script is in the `scripts/` folder beside it
and its feature map in `features/`.
