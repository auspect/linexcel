# Open-source package research

Findings checked on 2026-10-07. HTTPX provider configuration is implemented in
the accompanying code change; the other candidates remain proposals.
Package capabilities and licenses come from upstream sources.
Integration priorities and expected benefits are judgments based on this
repository. No candidate was benchmarked or security-audited in this task.

Start with transport configurability, property testing, dependency auditing and
repeatable profiling. Keep the existing Rust formula engine and cached-value
reader until measurements demonstrate a better alternative.

## Repository evidence

`pyproject.toml` already uses `formualizer>=0.10.1`, `python-calamine>=0.8.2`,
openpyxl and oletools. AI, Rich progress and Playwright remain optional. The
loader chooses calamine for cached values, falls back to openpyxl, and guards
dense worksheet allocations. Lazy analysis already uses SQLite and compressed
payloads. Replacing these paths with pandas or another dataframe library needs
a specific measured bottleneck.

The historical [reliability study](reliability_study.md) records substantial
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

1. Retain the implemented AI transport and its offline request assertions.
2. Add targeted generated tests and a locked dependency audit.
3. Establish phase timing and worker-memory baselines on fixed synthetic files.
4. Trial hardened XML parsing with streamed archive limits.
5. Evaluate JSON acceleration and an M parser oracle independently.
6. Prototype approximate viewer search only against a defined user task.

For each runtime change, require equal or clearer provenance and failure
messages, semantic regression coverage, and measured benefit where performance
motivates the dependency. Keep the trial optional until Windows/Linux and the
supported Python versions pass. Recheck licenses for the actual pinned artifacts.

This report does not certify Power Query correctness, benchmark candidate
packages, run a vulnerability scan, or complete the repository's manual AI/vision
acceptance procedure. See [manual validation](manual_validation.md) and
[acceptance requirements](guide/acceptance.md) for the integration acceptance run.
