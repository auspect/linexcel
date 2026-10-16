# Open-source package research

Initial findings checked on 2026-10-07; Excelize added on 2026-10-09.
HTTPX provider configuration is implemented; the other candidates remain proposals.
Package capabilities and licenses come from upstream sources.
Integration priorities and expected benefits are judgments based on this
repository. No candidate was benchmarked or security-audited in this task.

Start with reproduced interface defects, evidence-based AI review, property
testing and repeatable profiling. Trial Excelize as an optional display-value
oracle. Keep the existing Rust formula engine and cached-value reader until
measurements demonstrate a better alternative. The
[continuous improvement plan](continuous_improvement_plan.md) separates active
work, issue evidence gates and historical Markdown cleanup.

## Repository evidence

`pyproject.toml` already uses `formualizer>=0.10.1`, `python-calamine>=0.8.2`,
openpyxl and oletools. AI, Rich progress and Playwright remain optional. The
loader chooses calamine for cached values, falls back to openpyxl, and guards
dense worksheet allocations. Lazy analysis already uses SQLite and compressed
payloads. Replacing these paths with pandas or another dataframe library needs
a specific measured bottleneck.

The historical [reliability study](history/pre-1.15/reliability_study.md) records substantial
improvements from fixing XML scanning and bounded sheet context. Those are
earlier measurements, not measurements of the present tree or these packages.
`loader.py`, `lazy_store.py` and `powerquery.py` still contain explicit stdlib
XML parsing. Power Query reads DataMashup and scans M source statically.
The standalone viewer searches node labels, IDs and formulas by substring;
the application also supports paged server search for lazy graphs. CI already
checks types, tests, browser behavior and builds. Dependabot updates the lockfile.

## Prioritized candidates

| Priority | Candidate | Benefit to test | Integration cost | Decision |
| --- | --- | --- | --- | --- |
| P1 | HTTPX | Custom gateways, proxies, TLS and deterministic transport tests | Low; declared in the optional AI extra | Configurable provider implemented; retain integration tests |
| P1 | Excelize for Python | Compare source number formatting and displayed values independently | Medium; native library, locale policy, typed-value separation and fixtures | Trial as an optional display oracle; no loader or engine replacement yet |
| P1 | Hypothesis | Find parser, date and lineage edge cases | Medium; test generators and independent invariants | Add to development tooling when those tests are implemented |
| P1 | pip-audit | Detect published dependency vulnerabilities | Low; CI job and triage policy | Trial in CI against locked exported dependencies |
| P1 | defusedxml | Explicit policy for DTDs and entities | Medium; audit parsing sites and failure semantics | Trial alongside archive limits |
| P1 | pyperf | Separate genuine speed changes from timing noise | Low to medium; phase benchmark harness | Establish a baseline before performance dependencies |
| P2 | Memray | Find native and Python allocation hotspots | Medium; Linux profiling environment | Use as optional investigation tooling |
| P2 | orjson | Reduce large JSON export costs | Medium; serialization compatibility and wheel checks | Trial only if profiling identifies export cost |
| P2 | Microsoft powerquery-parser | Parse M syntax and inspect scope accurately | High; TypeScript/Node bridge or development oracle | Begin as a differential test oracle |
| P3 | Fuse.js | Rank approximate label/name matches | Medium; vendored asset and browser regressions | Prototype for bounded standalone reports |
| P3 | RESPX | Rich HTTP route assertions and retry scenarios | Low; development dependency | Use only if HTTPX MockTransport becomes cumbersome |

P1 means a useful next experiment, not proof that the package fixes a reproduced
defect. None of these candidates requires a paid service. Costs consist of
integration time, CI runtime, profiling overhead and dependency maintenance.

### Excelize for Python

The Python project is distributed as `excelize`, uses BSD-3-Clause, and lists
Windows, Linux and macOS binaries. It covers OOXML workbook families, charts,
images and styles. These are useful capabilities to study alongside linexcel's
read-only analysis; an editing API is not a reason to expand linexcel's scope.
[Python project and license](https://github.com/xuri/excelize-py),
[distribution setup](https://github.com/xuri/excelize-py/blob/main/setup.py).

`get_cell_value` returns a string with number formatting applied when supported;
otherwise it returns the unformatted value. `get_rows` has similar formatting
behavior and omits trailing empty cells. Merged-cell reads repeat the merged
value. The documented calculation API has limitations including arrays,
intersections, table formulas and iteration. This supports a formatting trial,
not a claim of complete Excel calculation parity.
[Python cell API](https://xuri.me/excelize-py/en/cell.html#get-cell-value).

`Options(raw_cell_value=True)` disables number formatting. Locale and date/time
pattern options make the display policy explicit. Opening options also include
ZIP expansion and worksheet/shared-string XML thresholds; use deliberate limits
inside the existing supervised worker. A raw string remains distinct from a
typed cache and from the original XML token.
[Python workbook options](https://xuri.me/excelize-py/en/workbook.html).

The wrapper exposes `get_cell_type`, `get_cell_style`, `get_cell_formula` and a
streaming `rows()` iterator. Its Python implementation calls a native shared
library; the inspected main-branch `go.mod` pins a Go Excelize pseudo-version.
Record the installed Python distribution, native artifact hash and Go version
actually used. Do not infer a released wheel's capabilities from the latest Go
documentation or the Python main branch.
[Python wrapper source](https://github.com/xuri/excelize-py/blob/main/excelize.py),
[Go dependency pin](https://github.com/xuri/excelize-py/blob/main/go.mod).

Decision: build a development-only comparator before adding a runtime extra.
The user's experience of accurate formatting and good speed is a useful test
hypothesis; this review has not installed or benchmarked Excelize. Compare the
same immutable workbook bytes through linexcel, Excelize and worksheet XML.
Keep separate records for raw source token/type, typed cached value, format code,
formatted text, formatter version/options and any recalculated result.
Formatted text must never become an engine input or an agreement verdict.

Use small checked fixtures for currency, percentages, scientific notation,
fractions, leading zeros, literal escapes, multi-section/conditional formats,
locale tags, Unicode, dates/times in both epochs, serial 60, durations, errors,
booleans, empty strings, missing caches and merged cells. Capture selected
reference strings from Excel under a recorded locale; LibreOffice is a second
implementation, not proof of exact Excel rendering. Separately inspect rich
text, conditional styling, column-width clipping and visual layout in images:
a formatted string alone does not establish the original screen appearance.

Gate adoption on case-level mismatches, repeated fresh-process output equality
under fixed options, worker peak memory, temporary-file cleanup, streaming versus
materialized reads, cold/warm timings, packaging on supported Python versions
and Windows/Linux. Benchmark against both the current calamine path and its
openpyxl fallback on identical dense and sparse inputs. Promote a bounded,
optional display adapter only if it improves verified formatting without losing
typed-value provenance. Consider cache-reader replacement separately, after
typed semantic and end-to-end performance evidence.

### HTTPX and RESPX

HTTPX provides custom transports, proxy routing and `MockTransport`.
The HTTPX project declares BSD-3-Clause. These facilities fit gateway routing
and offline provider tests. The locked OpenAI 3.13.0 SDK defaults to HTTPX2;
linexcel also declares HTTPX for client injection and tests. Avoid a general
retry dependency until provider and SDK retry behavior has been coordinated.
[HTTPX transports](https://www.python-httpx.org/advanced/transports/),
[HTTPX project metadata](https://github.com/encode/httpx/blob/master/pyproject.toml).

For linexcel, expose a supplied synchronous client and explicit provider
configuration. Test the final URL when a model appears in the path, percent
encoding, preserved gateway prefixes/query parameters, proxy configuration,
timeouts, extra headers/body parameters and client ownership. A transport mock
can inspect requests; proxy connection and TLS behavior need a local integration
test. Keep credentials out of exports and diagnostic text.

RESPX supplies HTTPX route mocking and pytest fixtures under BSD-3-Clause.
Its value would be simpler multi-request assertions for 429/5xx retries and
unexpected destinations. Start with the built-in mock transport to keep tooling
small. [RESPX](https://github.com/lundberg/respx),
[RESPX license](https://github.com/lundberg/respx/blob/master/LICENSE.md).

### Hypothesis

Hypothesis generates test inputs and reduces failing cases. Its repository uses
MPL-2.0 except where separately licensed. Keep it in the development group;
generated regression inputs can be synthetic workbooks.
[Hypothesis](https://github.com/HypothesisWorks/hypothesis),
[license](https://github.com/HypothesisWorks/hypothesis/blob/master/LICENSE.txt).

Prioritize `refs.py`, sparse worksheet XML, formula admission, date epochs and
Power Query lexical boundaries. Useful invariants include reference round trips,
consistent eager/lazy dependency sets, bounded processing of malformed input,
and preservation of quoted text through M splitting. Test shadowed names,
nested comments and quoted identifiers. Compare Excel semantics against an
independent oracle where possible; agreement between two linexcel paths cannot
prove that either matches Excel. Cap native-engine examples and retain process
isolation to contain crashes.

### pip-audit

pip-audit checks Python dependencies against published vulnerability databases
and can emit CycloneDX SBOMs. It uses Apache-2.0. It does not identify every
malicious package or establish that an application is secure. Its dependency
resolution can install packages, so use trusted locked input.
[pip-audit features and security model](https://github.com/pypa/pip-audit).

Add a separate CI audit of dependencies exported from `uv.lock`, including AI
and screenshots extras. Record advisory IDs and the affected resolved versions;
avoid automatic `--fix` in release CI. Give exceptions a reason and review date.
Test the audit with a known vulnerable synthetic requirements file. An audit
sends package metadata to an advisory service; workbook contents need not enter
that job. Bundled JavaScript and Rust transitive components need separate checks.

### defusedxml and archive limits

defusedxml exposes hardened XML parsing functions and explicit DTD/entity
rejection. It uses the PSF license. Its own documentation advises explicit
parsing calls and warns against globally monkey-patching stdlib XML.
[defusedxml parsing policy and license](https://github.com/tiran/defusedxml).

Introduce one parsing policy for workbook metadata, styles, relationships and
DataMashup XML, with `forbid_dtd=True`. Catch its rejection exceptions explicitly
and surface an actionable limitation. Audit openpyxl's XML path separately.
This is hardening work, not a reproduced exploit claim.

XML parser hardening does not bound ZIP expansion, large ordinary XML or base64
allocation. Check archive entry count, per-entry and total expanded bytes, then
enforce limits while reading streams. Bound outer custom XML and the decoded
DataMashup container before allocating them. Test entity/DTD rejection, truncated
archives, high expansion ratios and oversized M sections in isolated processes.
Benchmark ordinary files to assess policy overhead.

### pyperf and Memray

pyperf calibrates benchmark runs, starts worker processes, records metadata and
flags unstable timings. It uses MIT. Measure import, cached-value loading,
formula scanning, graph construction, targeted evaluation and export separately.
Record end-to-end wall time and worker peak memory too: profiling only the parent
misses isolated analysis work. Use fixed generated inputs, fresh output paths,
explicit time limits and separate cold/warm runs.
[pyperf](https://github.com/psf/pyperf),
[runner and memory options](https://pyperf.readthedocs.io/en/latest/runner.html).

Memray tracks Python and native allocations and uses Apache-2.0. Its supported
platforms are Linux and macOS, with WSL tested; it does not support native
Windows. Profile the actual analysis worker with native tracking. Treat Rust
symbol visibility and overhead as trial questions, then confirm improvements
with unprofiled runs. Preserve profiler files locally because paths and workbook
context can appear in them. Use existing Windows resource telemetry for Windows.
[Memray native mode](https://github.com/bloomberg/memray/blob/main/README.md),
[platform support](https://bloomberg.github.io/memray/supported_environments.html),
[license](https://github.com/bloomberg/memray/blob/main/LICENSE).

### orjson

orjson offers native JSON encoding and decoding, returns bytes when encoding,
and ships CPython wheels including Python 3.14. Upstream lists MPL-2.0 source
alongside MIT/Apache-2.0 source. Review the selected distribution's notices.
Upstream throughput figures do not predict linexcel's end-to-end speed.
[orjson behavior, packaging and licenses](https://github.com/ijl/orjson).

Trial an optional serializer in `viewer.py`, `web_export.py` or graph export
after profiling. Preserve browser-safe embedding, Unicode, dates, error values,
stable schema and integer precision. The existing lazy serializer rejects
non-finite numbers with `allow_nan=False`; orjson encodes them as `null`, so a
direct substitution changes that contract. Gate adoption on semantic round
trips, hostile `</script>` strings, wheel coverage, wall time and peak memory.

### Microsoft powerquery-parser

Microsoft maintains a TypeScript parser for the Power Query M language under
MIT. It can supply syntax trees for deeper source analysis. A syntax tree alone
does not evaluate queries, refresh connectors, reproduce query folding or
establish that saved results are current.
[parser](https://github.com/microsoft/powerquery-parser),
[license](https://github.com/microsoft/powerquery-parser/blob/master/LICENSE).

Start with a development-only Node helper that accepts synthetic M source and
returns JSON syntax information. Compare linexcel's splits and source references
with this oracle. Cover nested `let`, function parameters, shadowing, records,
quoted identifiers, comments and semicolons inside strings. If a runtime adapter
later proves useful, keep it optional, isolated and bounded. Node installation,
IPC, packaging and scope-aware AST traversal are real maintenance costs. Dynamic
URLs and conditional sources should remain unresolved or explicitly uncertain.

### Fuse.js

Fuse.js supports fuzzy search and weighted fields in the browser with no runtime
dependencies. Its currently inspected main branch declares Apache-2.0; verify
the license of the exact pinned release before vendoring.
[Fuse.js capabilities](https://github.com/krisk/Fuse),
[current license](https://github.com/krisk/Fuse/blob/main/LICENSE).

Use it to suggest near matches for labels, sheet names and query names while
preserving exact cell-ID lookup. Vendor assets locally to retain offline HTML.
Test accented/French labels, keyboard navigation, relevant ranking and adversarial
names. Measure index size, first interaction and repeated search on fixed small
and medium graphs. Keep the lazy application's paged server search; downloading
a whole large workbook index would undermine its memory design.

## Adoption sequence and evidence gates

1. Fix reproduced interface defects and retain the AI transport assertions.
2. Version the specialized AI review method and its source-evidence checks.
3. Trial Excelize display formatting with an independent reference corpus.
4. Add targeted generated tests and a locked dependency audit.
5. Establish phase timing and worker-memory baselines on fixed synthetic files.
6. Trial hardened XML parsing with streamed archive limits.
7. Evaluate JSON acceleration and an M parser oracle independently.
8. Prototype approximate viewer search only against a defined user task.

For each runtime change, require equal or clearer provenance and failure
messages, semantic regression coverage, and measured benefit where performance
motivates the dependency. Keep the trial optional until Windows/Linux and the
supported Python versions pass. Recheck licenses for the actual pinned artifacts.

This report does not certify Power Query correctness, benchmark candidate
packages, run a vulnerability scan, or complete the repository's manual AI/vision
acceptance procedure. See [manual validation](manual_validation.md) and
[acceptance requirements](guide/acceptance.md) for the integration acceptance run.
