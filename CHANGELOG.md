# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Seven more languages for `language=`, covering both the AI prompts and the
  viewer interface: `es`, `de`, `it`, `pt`, `nl`, `ja`, `zh` (with `en` and
  `fr`, nine in total). The set stays a closed allowlist — `language` selects a
  stored system prompt and reaches the generated JavaScript, so free-form input
  would be a prompt-injection and interpolation vector
- `linexcel.i18n`, single source of truth for the language list and the viewer
  interface strings, which used to be hand-maintained inside the JS template.
  The suite asserts `i18n.UI_STRINGS` and the two `aidoc` prompt registries
  cover the same languages and the same keys
- Translation provenance is stated in the README, the AI guide and the
  `linexcel.i18n` docstring: `en` and `fr` are hand-written, the other seven
  were produced with AI assistance and are not natively reviewed

- Multi-provider AI support: Google Gemini (default), OpenAI-compatible endpoints (Ollama, vLLM, LM Studio), and custom callables
- `base_url` and `provider` parameters on `document()` and `document_workbook()`
- Environment variables: `LINEXCEL_AI_BASE_URL`, `LINEXCEL_AI_MODEL`
- Optional dependency: `pip install linexcel[openai]`
- README: multi-provider documentation section
- `linexcel.__version__`, read from the installed distribution metadata
- `max_workers` on `document()` to tune AI request concurrency

### Fixed
- `provider=` now accepts a plain callable, as the documentation always claimed;
  previously only objects exposing a `generate` method worked
- VBA call edges were never emitted: the procedure lookup was keyed on
  `Module.Name` while the call scanner reports unqualified names. Calls now
  resolve in the calling module first, then across modules, and an ambiguous
  name stays unresolved instead of pointing at an arbitrary module
- VBA call detection is case-insensitive throughout, matching the language:
  procedure lookups are keyed on the lowercased name, so two modules declaring
  `Taux` and `taux` no longer capture each other's calls nor defeat the
  ambiguity guard, and a function's own return assignment (`Taux = 1`) is not
  mistaken for a call
- A dotted name is a VBA call only when the qualifier is a module
  (`Module1.Taux`, now resolved exactly); `.Value` or `.Count` is member access
  and no longer produces an edge to a similarly named procedure
- `document()` no longer discards every successful card when a single node
  fails; failures are reported through a `UserWarning` and `AiDocError` is
  raised only when all nodes fail
- `to_html()` no longer fails on a `LineageResult` built without the workbook
  bytes; the sheet-preview tab is simply omitted
- Viewer template placeholders are substituted in a single pass. Chained
  replacements let `__TITLE__` and `__LANG__` be rewritten *inside* the
  already-injected graph, so a cell, sheet name, or comment containing either
  literal was silently replaced by the title — and a title ending in a
  backslash truncated the embedded JSON, blanking the viewer
- The `cells` interface string was used twice by the viewer but defined in no
  locale, so reports fell through to the raw key ("1 000 cells" in French).
  The unused `sheets_summary_title` key was dropped

### Changed
- **Breaking:** `render_html()`, `to_html()` and `save_html()` reject a
  `language` outside `("en", "fr")` with `ValueError`. It was previously
  interpolated raw into the generated JavaScript, where it could terminate the
  string literal; unknown values used to fall back to English at runtime
- `pytest` now collects the `src/` doctests, which `testpaths` used to exclude
- Generated validation artifacts are no longer tracked in the repository
- AI data-handling notes describe every provider, not just Gemini

## [0.3.0] - 2026-07-14

### Added
- Localization: `en` (default) and `fr` language support for AI documentation and UI
- Sidebar worksheet overview in HTML report
- Sheet-specific screenshot embedding in HTML report
- Workbook-level AI documentation (`document_workbook()`)
- `SECURITY.md`, `THIRD_PARTY_NOTICES.md`

### Changed
- Refined AI data handling documentation and security notes

## [0.2.2] - 2026-07-10

### Changed
- Improved formula decomposition and step evaluation
- Refined dependency graph edge resolution for named ranges

## [0.1.0] - 2026-07-10

### Added
- Initial release
- Formula extraction via formualizer (Rust engine)
- Stretched pattern grouping via R1C1 canonicalization
- Dependency graph: cells, ranges, defined names, VBA procedures
- Step-by-step composite function evaluation
- Standalone HTML viewer (Cytoscape.js, fully offline)
- Optional AI documentation via Google Gemini
- LibreOffice screenshot rendering
- PyPI trusted publishing
