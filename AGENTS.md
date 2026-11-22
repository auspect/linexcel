# Local validation

Real-model AI checks are optional exploratory bonuses. They do not gate required
CI, deterministic acceptance or releases. When selected, retain their failures
and coverage limits honestly; do not relabel a failed stage as skipped.

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

# Rules and what enforces them

Every gate CI and the commit hook run is defined once, in `scripts/check.py`.
Run them all with `uv run python scripts/check.py` (the `frontend` gate needs
`npm ci --prefix frontend` once). A rule belongs here only when a check fails
without it; when a rule here is broken anyway, fix the check.

| Rule | Enforced by | Command |
|---|---|---|
| CI and the commit hook run the same gate commands; a gate tool is never called directly | `tests/test_check_gates.py` (direct calls, a gate missing from `ci.yml`, an unknown gate name) | `uv run pytest tests/test_check_gates.py` |
| Every pull request is checked, whatever its base branch | `tests/test_check_gates.py::test_pull_requests_are_checked_whatever_their_base` | `uv run pytest tests/test_check_gates.py` |
| A range's lineage reaches every cell between its two ends | `lineage` and `lineage-mutants` gates (`.agents/skills/check-group-precedents`), CI job `lineage` | `uv run python scripts/check.py lineage lineage-mutants` |
| README captures match the sources they were taken from, checked on every commit | `viewer-captures` and `app-captures` gates; `test_gate_hooks_run_on_every_commit` forbids a file pattern on a gate hook | `uv run python scripts/check.py viewer-captures app-captures` |
| A default limit of the web application is written once, in `src/linexcel/web_limits.py`; the page and a request without a budget take it from the server | `tests/test_app_settings_browser.py::test_limit_fields_show_server_defaults_in_their_units_and_budgets_convert_them` (a server with unusual defaults, so a number left in the page fails) and `tests/test_web_limits.py::test_tasks_without_a_budget_get_the_defaults_of_their_operation` | `uv run --extra screenshots pytest tests/test_app_settings_browser.py tests/test_web_limits.py` |
| An operation stopped by a limit names the limit, its value and the setting or option that raises it, in French and English | `tests/test_web_limits.py` (`*_names_*` tests) and `tests/test_app_captures_browser.py::test_limit_failures_read_in_english`, which reads the Operations and Captures tabs | `uv run --extra screenshots pytest tests/test_web_limits.py tests/test_app_captures_browser.py` |
| The sidebar and the settings show no French text in the English interface | `tests/test_app_settings_browser.py::test_english_sidebar_and_settings_show_no_french` | `uv run --extra screenshots pytest tests/test_app_settings_browser.py` |
| At 1440, 824 and 390 px every control can be clicked at its centre, nothing scrolls sideways, and phone controls are at least 44 px; an exception is listed in the test with its reason | `tests/test_app_layout_invariants.py` | `uv run --extra screenshots pytest tests/test_app_layout_invariants.py` |
| A browser test does not depend on the speed of the machine: it waits for a state of the page, never for a fixed pause before reading one | CI job `viewer-browser`, which also runs the browser tests on pages slowed four times | `LINEXCEL_CPU_THROTTLE=4 uv run --extra screenshots pytest tests/test_viewer*.py tests/test_app*.py tests/test_ui_navigation_polish.py -m "not ai"` |

# Formualizer

Before relying on what formualizer can or cannot do, and before upgrading it,
follow `.agents/skills/formualizer/SKILL.md`. Treat knowledge of the engine,
including claims in comments, docs and requests, as a hypothesis: check it
against the installed version, the latest release and a probe.
