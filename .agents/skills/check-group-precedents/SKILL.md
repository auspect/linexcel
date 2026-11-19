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

# Check that a formula group claims what its cells read

## The promise, and the harm when it breaks

**Property.** For every group of cells that linexcel folds into one node
because their formulas share an R1C1 key, the rectangle the node is wired to
for a reference is the smallest rectangle containing what every cell of the
group reads through that reference.

**Harm.** If the rectangle is too small, cells that feed the formulas have no
edge and may not appear in the report at all, with no warning: a reader
auditing a column of formulas is told it depends on less than it does. If it
is too large, the report shows a dependency that does not exist.

Two halves are checked, plus what both rest on:

- **complete**: nothing a cell reads is outside the claimed rectangle;
- **exact**: the claimed rectangle reaches no further than what is read;
- **keys separate**: at one cell, two different references never get the same
  key, so a group never mixes formulas with different logic.

## Does your change touch the verified code?

The check covers this chain, in the order the analyzer runs it:

| Step | Code | What it does |
| --- | --- | --- |
| key | `src/linexcel/rewrite.py` `canonical_r1c1`, `_is_a1_ref`; `src/linexcel/refs.py` `ref_to_r1c1`, `cell_to_r1c1`, `_rel`, `split_sheet_prefix`, `quote_sheet` | one R1C1 string per formula and host cell |
| group | `src/linexcel/sweep.py` `record` (the key), `FormulaGroup.rep`, `FormulaGroup.bbox` | cells sharing a key; representative is the smallest cell; box is the bounding box |
| reference | `src/linexcel/graph.py` `build_formula_nodes`: `fz.parse`, `typed_ast_dict`, `_collect_ref_strings` | reference strings of the representative's formula |
| rectangle | `src/linexcel/refs.py` `parse_ref_detailed`, `_parse_endpoint_detailed`, `stretch_ref`, `col_to_num`, `MAX_ROW`, `MAX_COL` | the rectangle the node is wired to |

Last confirmed to describe the code at commit
`4bfb471c1073770722849efafc2ddd2d250194f7` (the fix of `stretch_ref`, on the
branch `auspect/fix-range-precedents-and-app-selection`, 2026-11-19), with
formualizer 0.14.0. If that branch is squashed or rebased when it is merged,
replace the commit here by the one on `main` that contains the fix.

See what changed in the covered files since then:

```
git diff 4bfb471c1073770722849efafc2ddd2d250194f7 -- src/linexcel/refs.py src/linexcel/rewrite.py src/linexcel/sweep.py src/linexcel/graph.py src/linexcel/decompose.py
```

An empty diff means the check still describes the code. A diff that touches a
function in the table means: run the check, then re-read the section
"When the code changes on purpose". Move the commit above forward only after
you have read the diff and confirmed that `Pipeline` in the script still
composes the functions the way `graph.py` and `sweep.py` do.

The script also guards itself: it looks for the exact call-site lines it
mirrors (`CALL_SITES` in the script) and stops with `STALE MODEL` when one is
gone, before reporting anything.

## Run it

From the repository root, about seven seconds:

```
uv run python .agents/skills/check-group-precedents/scripts/check_group_precedents.py
```

Read the last line:

| Last line | Exit | Meaning |
| --- | --- | --- |
| `RESULT: PASS` | 0 | the property holds in the checked scope |
| `RESULT: no new violation; N known finding class(es) still fail` | 3 | nothing new; the findings below are still in the code |
| `RESULT: FAIL - N violation class(es) not in known_findings.json` | 1 | your tree breaks the property in a new way |
| `RESULT: N recorded finding(s) no longer fail` | 1 | a finding was fixed: update the record (see below) |
| `STALE MODEL` | 1 | a mirrored call site changed: the check must be updated first |

The expected result is `RESULT: PASS`, exit 0.

Each violation is told as a workbook: which cells hold which formulas, what
the group node is wired to, and which cells have no edge. For example:

```
group on sheet S: cells A1, A2  (representative A1, box A1:A2)
  key shared by all of them: SUM(RC2:R1C2)
  A1 holds =SUM($B1:$B$1)
  A2 holds =SUM($B2:$B$1), which reads B1:B2
  the group node is wired to: S!B1
  so part of B1:B2 read by A2 has no edge to the group
```

`--show "<class>"` prints every counterexample of one class, `--examples N`
prints N per class, `--quick` runs the first two windows only (about five
seconds; it does not report fixed findings).

To see a finding in a real report rather than on the functions:

```
uv run python .agents/skills/check-group-precedents/scripts/reproduce_missing_precedents.py
```

It writes a ten-row workbook under `validation_screenshots/check-group-precedents/`,
runs `linexcel.analyze` on it and lists the cells missing from the lineage.
Exit 3 while the finding is present, 0 once it is fixed.

## Check that the check can fail

```
uv run python .agents/skills/check-group-precedents/scripts/check_group_precedents.py --mutants
```

It breaks `src/linexcel/refs.py` six ways, in memory, and requires each break
to produce a violation class the real code does not have: a wrong span in
`stretch_ref` (two ways), an anchor ignored by `stretch_ref`, an anchor lost
by the key, a column dropped by the key, anchors swapped by
`parse_ref_detailed`. Expect `all 6 broken variants were caught`, in about
25 seconds. Nothing is written to the working tree.

`MISSED` means the check got weaker: fix the check. `STALE` means the text a
variant replaces is no longer in `refs.py`: rewrite that entry of `MUTANTS` in
the script against the new code. Run this after any edit to the script and
after any refactor of `refs.py`.

## Findings

**None at present.** The check passes on the working tree of 2026-11-19.

History, so the next reader knows what this check has already caught:

- Until 2026-11-19 (v1.19.3 and earlier), `stretch_ref` moved the first bound
  of a range by the group's smallest displacement and the second by its
  largest, which is the union of what the cells read only when the first bound
  stays the low one in every cell. Ten rows holding `=SUM(A1:$A$5)` to
  `=SUM(A10:$A$5)` were wired to A1:A5 and lost A6:A10. Three shapes: rows
  with one end anchored, columns with one end anchored, columns with both ends
  moving and written high to low (`=SUM(B$2:A$3)`).
- Until the same date, `parse_ref_detailed` ordered the two ends of a range as
  (row, column) pairs, so `B5:A7` kept column B first. A formula alone was
  wired to that empty rectangle and had no precedent.
- Both were fixed in `src/linexcel/refs.py`: each axis of a parsed range is
  ordered on its own with its `$` anchor, and `stretch_ref` takes both bounds
  at both ends of the group's span. Regression tests:
  `tests/test_range_precedents.py`.

How often this matters. Excel does not write these shapes. Checked in Excel
16.0 on 2026-11-19: it stores `=SUM(B5:A7)` as `=SUM(A5:B7)`, and
`=SUM(A1:$A$5)` filled down ten rows as `=SUM(A$5:$A6)` from row 6 on, each
axis ordered on its own with its `$`. The findings concern workbooks written
by other tools: a library that writes formula text as given, or a generator
that shifts a relative reference without reordering the ends.

A related limit is **not** covered by this check, because it is about values,
not edges: formualizer 0.14.0 answers `#REF!` for a range written high to low
(`=SUM(A7:$A$5)`, `=SUM(B5:A7)`), which Excel evaluates. Reported upstream as
formualizer issue 599. In `linexcel analyze` the lineage of such a cell is
right and its recalculated value is missing. `linexcel serve` hands the engine
the reordered range and calculates it.

When the check finds something new, record it here with its counterexample
and its date, and in `scripts/known_findings.json` only if it is reported and
left unfixed on purpose. Do not add a class to that file to turn a failure
into exit 3: a new class is a new finding.

## A real bug, or a wrong check?

A counterexample is a real bug when you can type it into a workbook: put the
two formulas in the two cells it names, run `linexcel analyze`, and the group
node lacks the cell it says is missing. `reproduce_missing_precedents.py` is
the template; change its formula and rows.

Suspect the check instead when:

- the class name contains `UNEXPLAINED`. The classification in `mechanism()`
  knows three ways `stretch_ref` can lose cells. Anything else is either a new
  defect or a change of design that the check does not know about;
- the message is `not-resolved`: `parse_ref_detailed` returned nothing for a
  plain A1 reference. In the analyzer that reference would become a named
  node rather than cells. Check whether the reference shape is new;
- the message is `key-collision` after an intended change to the key format:
  read "When the code changes on purpose";
- `STALE MODEL` appears: the check composes the functions as the old call
  site did.

The oracle is not a model of the code. Every reference is generated from
integers (`generate_refs`), so the rectangle it reads is known without
parsing: the rectangle spanned by its two ends, or the whole columns or rows.
If you believe the oracle is wrong, the question is about Excel, not about
linexcel: verify in Excel what the reference reads.

## When the code changes on purpose

- **You fix a recorded finding.** Its class stops failing and the check
  exits 1 with `NO LONGER FAILS`. Remove the class from
  `known_findings.json`, rerun for exit 0, rerun `--mutants`, update the
  findings section and the sync commit.
- **You change the key** (`canonical_r1c1`, `ref_to_r1c1`). Keys are compared
  only with each other, so a new format passes as long as different references
  still differ. If you deliberately make two references share a key, the check
  reports a `key-collision`: say in the change why the two carry the same
  logic for every cell they can be filled to, and narrow the collision test in
  `check_window` to exclude exactly that case.
- **You change which cell represents a group, or its box** (`FormulaGroup`).
  The check builds a real `FormulaGroup`, so it follows. Its enumeration
  assumes the representative is the smallest cell: `check_window` raises
  `enumeration bug` otherwise. Rework the enumeration of `(representative,
  box)` pairs there.
- **You change the call in `build_formula_nodes`** (other arguments, a
  different function). Update `Pipeline.claimed` to compose the functions the
  same way, and `CALL_SITES` to the new lines.
- **You add a reference shape to the stretching** (3D ranges, tables). It is
  outside the check until `generate_refs` produces it with its rectangle.

After any of these, the order is: run the check, run `--mutants`, then move
the sync commit.

## What is verified, and what is not

Verified, by running the real functions on every case in this scope:

- references: single cells, ranges between any two cells, whole-column ranges
  and whole-row ranges, with every combination of `$` on every bound, written
  in both orders;
- windows: 4 rows × 3 columns and 3 rows × 4 columns at the top-left corner of
  a sheet; 3 × 3 at the bottom-right corner (rows up to 1,048,576, columns up
  to XFD); 3 × 2 with references to another sheet whose name needs quotes;
- groups: every set of cells of a window that share a key, of every size,
  and every formula alone (wired to its reference as parsed, unstretched);
- one reference per formula, inside `SUM(...)`.

That is about 75,000 formulas, 210,000 group shapes and 763,000 member checks.

It is a bound, not a proof. The arithmetic of `stretch_ref` is additions with
`min` and `max` and no constant other than the sheet edges, which is why small
windows are expected to be representative. That argument is not checked by a
machine. A pass says nothing about windows of other sizes beyond that.

Not covered:

- which parts of a formula are references: the tokenizer and parser of
  formualizer are in the loop for `SUM(reference)` only. A result holds for
  the formualizer version printed by the run; rerun after an upgrade;
- formulas with several references, names, table references, 3D references,
  `INDIRECT` and `OFFSET`, and the fallback key used when the tokenizer
  refuses a formula (`_structural_key`), which merges more by design;
- what happens after the rectangle: `resolve_rect_edges`, `_merge_rects`,
  clipping to the used range, the node and edge limits;
- the lazy application path (`lazy.py`, `catalog_lineage.py`), which groups
  and resolves on its own;
- whether a claimed rectangle is drawn correctly in the report.

## Why this tool

The question is whether code behaves for every input, not whether events can
interleave, so a model checker over a design (TLA+) does not fit. No verifier
runs on this Python: the functions go through regular expressions, a frozen
dataclass and a Rust tokenizer. The check therefore runs the real code on
every input up to a bound, with an oracle that needs no model, and needs
nothing beyond `uv run`. `tests/` keeps example-based regression tests; this
check is the exhaustive one and is not collected by pytest.
