# Bounded local audits

The default API and CLI run analysis in an isolated process with a **120-second
budget** and a **2,048 MiB memory limit**. This is a limit, not a prediction.
Reading the source into memory, AI requests, screenshot rendering and HTML
export are outside that analysis budget. Workbook inputs are never rewritten.
Process startup adds overhead, particularly for small workbooks; isolation is
a safety boundary, not a claim that every analysis becomes faster. After an
interruption, bounded process and temporary-file cleanup can extend the return
time beyond the analysis deadline.

```python
from linexcel import analyze, ExecutionPolicy

result = analyze("book.xlsx", execution=ExecutionPolicy(seconds=120, memory_mb=2048))
print(result.graph["meta"]["execution"])
result.save_html("audit.html")
```

```console
linexcel analyze book.xlsx --analysis-seconds 120 --memory-mb 2048 -v
```

Progress reports observed phases and elapsed time. A phase can remain unchanged
while a native call runs; there is no invented percentage or ETA. The old
`estimatedSeconds` inspection field remains present as `null`, accompanied by
`estimateStatus="deprecated_unavailable"`. `estimate_seconds()` is deprecated
and returns `None`. `--time-budget` / `step_seconds` retain their narrower,
cooperative decomposition budget; they cannot interrupt an active native call.

## Results and compatibility

`meta.execution.status` distinguishes `completed`, `timed_out`, `cancelled`,
`memory_limit` and `crashed`. Completion means the pipeline returned, not that
every formula was evaluated or that its values match Excel. Native exit codes,
bounded error diagnostics, completed phases and observed phase durations are
retained. A native initialization failure can remain `crashed` when the cause
cannot be identified reliably; a low memory budget alone is not proof of cause.

On interruption, only completed source evidence is retained. No partial native
calculation is promoted to a recalculated value. Unknown formula counts remain
`null`. CLI interruptions do not start AI or screenshot work. An incomplete
diagnostic export exits with code 3; cancellation exits with code 130.

The isolated result has **`result.engine is None`**. Graph access, documentation
and exports remain available. Code that needs to interact with a live native
workbook must explicitly opt into the legacy process behavior:

```python
result = analyze("book.xlsx", execution=ExecutionPolicy(isolated=False))
value = result.engine.get_value("Sheet1", 1, 1)
```

That mode has no hard time/memory guarantee and can expose the caller to native
crashes. Linexcel never switches to it automatically after an isolated failure.

Windows starts the worker suspended, attaches a Job Object before initialization,
then resumes it. The Job caps aggregate committed memory and kills descendants
on closure. POSIX installs `RLIMIT_AS` before importing linexcel and cleans up the
worker process group: this is a per-process virtual-address limit, **not an
aggregate RSS limit**. The `memoryLimitKind` field records this distinction.

## Coverage and semantic evidence

`meta.analysisCoverage` records the requested scope, extracted formula count,
known omissions and their phase/reason. Unknown omitted counts are never
estimated. Static traces and grouped values do not certify complete dependencies.

`meta.coverage` counts **graph nodes**, not all workbook cells. The engine/cache/
unavailable/other categories partition nodes; divergences are an overlapping
category and include disagreements found in group samples. Constants and range
samples are not counted as recalculated formulas. Names, code and other nodes
without a scalar are not counted as failed calculations.

The [engine safety notes](engine-safety.md) describe capability probes and their
limits. Potentially affected formulas preserve their observed values but carry
an explicit verification limitation, propagated to represented dependents.
Neither matching a cache nor passing synthetic probes proves Excel equivalence.

## Separate optional work

The built-in AI client allows 300 seconds per HTTP request by default, configured
through `LINEXCEL_AI_TIMEOUT_SECONDS`, with no automatic retries. This transport
timeout is separate from the analysis budget. Token budgets are checked between
requests; already submitted responses can exceed the token budget as documented.
Custom providers remain responsible for their own timeout behavior. Screenshot
rendering retains its separate per-command timeout (default 180 seconds).

Private inputs, prompts, responses and reports belong in ignored local storage.
Acceptance requires a complete `validate_manual.py` run plus independent review
of the exact outputs; successful generation alone is not a factual verdict.
