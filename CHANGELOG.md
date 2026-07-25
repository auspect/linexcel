# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Multi-provider AI support: Google Gemini (default), OpenAI-compatible endpoints (Ollama, vLLM, LM Studio), and custom callables
- `base_url` and `provider` parameters on `document()` and `document_workbook()`
- Environment variables: `LINEXCEL_AI_BASE_URL`, `LINEXCEL_AI_MODEL`
- Optional dependency: `pip install linexcel[openai]`
- README: multi-provider documentation section

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
