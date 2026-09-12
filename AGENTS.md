# Local validation

For changes to analysis, evaluation, documentation, screenshots or the viewer,
run `validate_manual.py` against the changed working tree before delivery.
The minimum local acceptance run includes both generated workbooks, AI node
documentation, workbook overviews, sheet screenshots and AI analysis of those
screenshots. Use local Ollama with `qwen3.8` for documentation and vision unless
the user selects another model.

```powershell
uv run --extra ai --extra screenshots validate_manual.py --workbook both --vision --model qwen3.8 --vision-model qwen3.8 --token-budget 1000000
```

Keep generated workbooks, screenshots, model responses and real-world corpus
files in the ignored `validation_screenshots/` directory or a local parent
directory. Never commit private corpus files or reports to GitHub.

Inspect every dashboard tab, the graph, and a reproducible sample of nodes
spread across sheets and edge cases. Compare documentation and decomposition
with the source formulas and value provenance. Capture the dashboard itself
at desktop and mobile widths and inspect those screenshots too. Review AI
image descriptions against the actual images; generation success is not proof
of factual correctness. Record discrepancies and remaining limits honestly.

A unit-test-only run, `--no-ai`, `--no-vision`, `--max-nodes`, missing captures,
empty/truncated AI responses, or a partial language run does not satisfy this
acceptance requirement. Preserve each run separately and inspect its
`validation.json`. Include relevant changelog entries and independent review
before delivering a PR. If a required local dependency prevents completion,
report the failed stage explicitly instead of declaring validation successful.
