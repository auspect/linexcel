---
name: formualizer
description: >-
  Use before writing, changing, reviewing or explaining linexcel code that relies on
  formualizer (formula parsing, tokens, references, evaluation, recalculation,
  dependency traces, spills, external links, names, tables), before upgrading or
  pinning formualizer, and whenever anyone — you, a comment, a doc, a memory or the
  developer's request — claims what formualizer can or cannot do. Treats every
  such belief as a hypothesis to verify against the installed version, the latest
  release and linexcel's actual use of the API.
---

# formualizer: verify before you believe

formualizer is the Rust spreadsheet engine under linexcel. It changes quickly:
releases 0.8 to 0.14 shipped within a few months and kept adding trace,
inspection, spill, provenance and recalculation APIs. Whatever you remember about
it, and whatever this repository's comments and docs say, describes *some*
version. linexcel may also use only part of what the engine offers.

So your prior knowledge, the developer's assumptions in the request, code
comments and earlier plans are **leads, not facts**. Confirm them against the
installed package and the latest release before acting on them, and say so
plainly when one turns out wrong, including one stated by the developer.

## Protocol

1. **Inventory first.** From the repository root:

   ```powershell
   uv run python .agents/skills/formualizer/scripts/inventory.py --online
   uv run python .agents/skills/formualizer/scripts/inventory.py --compare <latest>
   ```

   It prints the installed, declared (`pyproject.toml`), locked (`uv.lock`) and
   latest PyPI versions, the stub files to read, public names and class members
   that `src/linexcel` never mentions, comments that assert engine limits, and the
   API added or removed in another release. If latest is newer than locked, answer
   "does that release change what I am about to do?" with evidence.

2. **Write the beliefs down.** Before coding, list each belief your plan relies
   on ("the engine cannot evaluate X", "linexcel must parse Y itself", "trace()
   omits Z"), with where it came from: memory, comment, doc, test, request.

3. **Verify each belief, strongest evidence first.**
   - **Probe the behaviour.** Run a minimal script against a tiny workbook, on the
     locked version and on the latest one, without touching the lock:
     `uv run --no-project --with formualizer==<version> python probe.py`.
     Write probe files to the scratchpad or `validation_screenshots/`, never to
     the repository. A file saved by real Excel beats a hand-built one when the
     question is about what Excel stores.
   - **Read the installed stubs:** `formualizer_py.pyi` and `__init__.pyi`, for
     signatures and docstrings. The inventory gives their paths.
   - **Read upstream:** the
     [changelog](https://github.com/PSU3D0/formualizer/blob/main/CHANGELOG.md),
     [releases](https://github.com/PSU3D0/formualizer/releases) and docs for every
     version between locked and latest.
   - **Read the repository's own evidence, as history:**
     `tests/test_formualizer_contract.py`,
     `tests/test_formualizer_release_contracts.py`,
     `.github/workflows/formualizer-compat.yml` (floor, locked and latest
     channels) and `docs/formualizer_upgrade_plan.md`. Dated addenda record what
     was true then.

4. **Challenge the other direction too.** For the area you touch, look for
   linexcel reimplementing in Python what the engine now provides: parsing,
   reference extraction, dependency walking, spill handling, function semantics
   or provenance. Also look for workarounds and bridges that a newer release no
   longer needs. The inventory's "never mentioned" lists are where to start.

5. **Make beliefs self-checking.** Do not write undated claims such as
   "formualizer does not support X". Write "formualizer 0.13.0 returns #SPILL!
   for a loaded spill (see test_...)" and back it with a contract test that fails
   when upstream behaviour changes, so the belief retires itself.

6. **Report.** End with a short **Beliefs checked** list: belief, evidence (probe
   output, stub line or changelog entry), verdict (confirmed, refuted, or changed
   in version N).

## linexcel principles that still apply

- Saved workbook values are the facts. An engine result never silently replaces
  a cached value. Recalculation is explicit and labelled as such.
- Lineage is deterministic. A model may search or phrase, but never supplies a
  fact.
- Excel is the reference for semantics, not the engine's documentation.
- A floor raise in `pyproject.toml` is only needed when linexcel calls a newer
  API. The compat workflow's floor channel must keep passing.

## Red flags: stop and verify

- "formualizer can't / doesn't / always / never …" with no version and no probe.
- A Python reimplementation of anything in the list in step 4.
- `try`/`except` around engine calls that swallows errors or falls back silently.
- Comments naming a version older than the locked one.
- Upper bounds on formualizer, or skipping the compat contracts.

## Upgrading the lock

```powershell
uv lock --upgrade-package formualizer
uv run pytest tests/test_formualizer_contract.py tests/test_formualizer_release_contracts.py -q
uv run --isolated pytest -q -m "not ai"
```

Then add a dated addendum to `docs/formualizer_upgrade_plan.md` and a CHANGELOG
entry that states what changes for users.

## Worked example (2026-11-10, a dated snapshot to re-verify)

`src/linexcel/values.py` said "formualizer does not compute dynamic-array
overflow, so a spilled formula reads back as #SPILL!". The inventory showed
0.13.0 locked and 0.14.0 published. A probe loaded a saved `SEQUENCE(3)` spill
with `SUM(A1#)`:

| Version | Spill A1:A3 after `evaluate_all` | `SUM(A1#)` |
| --- | --- | --- |
| 0.13.0 | `#SPILL!`, 2, 3 | `#REF!` |
| 0.14.0 | 1, 2, 3 | 6 |

The comment was true for 0.13.0 and false for 0.14.0. A free in-memory spill
already worked in both. The 0.14.0 comparison also listed new `Source*`
dependency-evidence and `Recalc*` job APIs that linexcel did not use yet.
