# Local validation

For changes to analysis, evaluation, documentation, screenshots or the viewer,
run `validate_manual.py` against the changed working tree before delivery.
The minimum local acceptance run includes both generated workbooks, analysis,
workbook context, sheet screenshots and all report languages. AI node
documentation, workbook overviews and image analysis are optional. Run without
AI when no model validation is requested:

```powershell
uv run --extra screenshots validate_manual.py --workbook both --no-ai
```

For optional AI validation, omit `--no-ai` and use local Ollama with `qwen3.8`
for documentation and vision unless the user selects another model. Explicitly
requested AI stages must complete; an unavailable model is a failure, not a skip.

Keep generated workbooks, screenshots, model responses and real-world corpus
files in the ignored `validation_screenshots/` directory or a local parent
directory. Never commit private corpus files or reports to GitHub.

Inspect every dashboard tab, the graph, and a reproducible sample of nodes
spread across sheets and edge cases. Compare documentation and decomposition
with the source formulas and value provenance. Capture the dashboard itself
at desktop and mobile widths and inspect those screenshots too. Review AI
image descriptions against the actual images; generation success is not proof
of factual correctness. Record discrepancies and remaining limits honestly.

A unit-test-only run, `--max-nodes`, missing captures or a partial language run
does not satisfy this acceptance requirement. `--no-ai` and `--no-vision` are
valid optional-stage selections; record the skipped stages without claiming AI
coverage. Empty/truncated responses fail when that AI stage was requested.
Always verify visually at desktop (1440×900) and mobile
(390×844) viewports with an independent AI reviewer agent, checking bounding
boxes and notice scoping across tabs. Preserve each run separately and inspect its
`validation.json`. Include relevant changelog entries and independent review
before delivering a PR. If a required local dependency prevents completion,
report the failed stage explicitly instead of declaring validation successful.
