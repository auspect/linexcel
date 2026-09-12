# Local acceptance with Ollama

Run the real pipeline on the current working tree before delivering changes to
analysis, evaluation, AI documentation or the viewer. Automated tests complement
this run; they do not replace it.

```powershell
uv run --extra ai --extra screenshots validate_manual.py --workbook both --vision --model qwen3.8 --vision-model qwen3.8 --token-budget 1000000
```

Ollama must serve the selected model locally. LibreOffice and Poppler must be
available for workbook rendering. The default endpoint is
`http://localhost:11434/v1`; there is no cloud fallback. Both generated workbooks
and image analysis are enabled by default. Calls use one worker by default to
avoid competing for a local model. A token budget is checked between requests;
it includes documentation and image descriptions across languages for each
workbook. Increase it explicitly when the manifest reports missing responses.

Each invocation creates a new directory under the ignored
`validation_screenshots/manual/`. `--output-dir PATH` chooses another new
directory, such as a private sibling workspace. Existing directories are
refused. `--file PATH` reads a real workbook without rewriting it. Keep real
workbooks and all reports embedding their contents outside GitHub.

The run saves graphs, PNG files, AI responses, HTML reports and `validation.json`.
The manifest records working-tree source hashes, configuration, input and image
hashes, durations, token usage and coverage by language. A successful process
requires every requested node card, overview, image and image description.
Missing providers, incomplete languages and empty or truncated model responses
fail. Diagnostic options `--no-ai`, `--no-vision`, `--no-recalc` and `--max-nodes`
produce partial evidence and return a nonzero exit status. A node limit uses
reproducible sampling across sheets; `--seed` controls that sample.

Generation coverage does not establish factual correctness. The manifest leaves
manual review pending. Open every dashboard tab, inspect graph connections and
sample nodes across sheets and difficult formulas. Check the original formula,
each evaluated step, cached/recalculated value provenance and the AI's claims.
Check each image description against its source image. Capture the dashboard
at desktop and mobile widths and inspect text, overflow, selection, navigation
and image readability. Record discrepancies with node or sheet identifiers,
the exact run directory and the evidence that supports the finding.

The stress fixture includes intentionally invalid references, circular formulas
and a formula longer than Excel's 8,192-character limit. It tests resilience;
it is not a fully Excel-compatible correctness oracle. LibreOffice is a second
implementation, so disagreement with its saved values requires investigation.
An AI description is another observation to review, not an independent numeric
oracle.
