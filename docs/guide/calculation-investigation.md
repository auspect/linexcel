# Calculation investigation: scoped constant names

Status on 19 September 2026: **tested private prototype, not installed or
published upstream**. Linexcel still uses the published formualizer 0.9.3.
Its existing runtime capability observations and warnings remain unchanged.
The findings below distinguish that dependency from a separately compiled
experimental engine; they do not establish general Excel compatibility.

## Cause and experimental correction

The published Calamine importer accepts defined names referring to cells or
ranges but drops numeric and text constant definitions. The native engine
already supports scoped literal names. The prototype passes those scalar
definitions to the native name API without replacing formulas, introducing
helper cells or changing source caches.

The correction is limited to finite numeric literals and quoted text. It
preserves escaped quotes and spaces around XML entities. Invalid sheet scope
is rejected, and the scope-losing fallback is removed so a rejected local
definition cannot reappear as a workbook name. Boolean/error constants,
arrays, computed name definitions and Excel's extreme numeric limits remain
outside the demonstrated scope.

The prototype changes only the Calamine importer. Umya removes quotation
information before this conversion point, making text constant `"2"`
ambiguous with numeric constant `2`. Extending that importer requires
authoritative source metadata rather than guessing the value's type.

## Measured results

The same 40 runtime probes were executed with both engines, without changing
their inputs or expectations. The published engine matched **21/40**; the
prototype matched **27/40**. Exactly six constant-name observations changed.
The other observed values and types remained identical, including the twelve
comparison differences and one date-output type difference.

A separate synthetic matrix contains **168 observations**: 14 cases across
three sheets, two date epochs and two importers. It overlaps the runtime
probes and must not be added to them as disjoint coverage.

| Engine | Calamine | Umya | Total |
| --- | ---: | ---: | ---: |
| Published 0.9.3 | 36/84 | 36/84 | 72/168 |
| Private prototype | 72/84 | 36/84 | 108/168 |

The matrix checks global/local constants, sequential `LET` bindings, local
shadowing, scalar/reference distinctions with `ISREF`, negative numbers,
quoted and numeric text, XML entities and reference-defined names. There
were 36 improvements and no regression among previously passing cases.
Original XLSX bytes, including deliberately populated formula caches,
retained their hashes. Cache preservation does not mean those caches agree
with recalculation.

Calamine's twelve remaining failures are qualified names such as
`'O''Brien'!Rate` and invoked `LAMBDA` expressions, each repeated six times.
They also fail with the published engine. In separate invalid-scope tests,
Calamine rejected four malformed/out-of-range scope forms without promoting
them globally; a valid local definition remained local. Umya raised existing
native panic exceptions on the four invalid forms. Those exceptions are
retained failures, not successful robustness checks.

The modified Python binding was compiled and loaded in isolation. Independent
review checked the final patch, binary identity, scope handling and recorded
results, accepting the bounded private prototype. The broader upstream
regression suite has not been run. No performance improvement is claimed.

## Separate performance baseline

A synthetic baseline on main revision `4ebcf4e` used approximately 100,000 cells and
20,000 formulas, represented by two grouped graph nodes and one edge with
no reported omission. Three analysis runs took **2.6951107, 2.6274262 and
2.7150742 seconds**: median 2.6951 seconds and sample standard deviation
0.04594 seconds. Median peak resident memory was **125,239,296 bytes**,
measured through Windows `GetProcessMemoryInfo` and including native
allocations. This is RSS/working-set evidence, not committed memory.

Timing excludes module import, API startup, JSON export (measured separately)
and HTML generation. An outer isolated process imposed a 120-second
limit. The host was an Intel Core i7-13700K running Windows 11 with 31.69 GiB
of physical memory and Python 3.13.12. These three local runs are a baseline,
not a before/after comparison or a general speed guarantee.

A separate profile counted three `evaluate_all` calls and two `get_formulas`
calls. The latter correspond to two 20,000-row chunks, rather than duplicate
extraction. The 20,000 `tokenize` calls consumed approximately 9 ms in that
profile. These observations do not justify optimizing tokenization or
removing extraction calls; no performance change was made on that basis.

## Separate number/text comparison prototype

A second, separate engine build changes only comparison operators between
numbers and text. The published fallback coerces numeric-looking text to a
number, or compares converted strings when that coercion fails. The candidate
instead preserves number/text type ordering. It does not include the name
importer patch above, and neither experimental binary is installed in linexcel.

On the unchanged 40 diagnostic observations, this candidate changes the result
from 21 to 29 matches: eight comparison observations are repaired. Four other
comparison observations also involve an input encoding issue: a whitespace-only
cell written without `xml:space="preserve"` is imported as empty. A separately
hashed variant explicitly preserving the text gives 33/40 matches; the remaining
six name observations and one date-output observation are unchanged. These
different inputs must not be presented as the same benchmark or silently
substituted for the original diagnostic.

A comparison matrix contains 288 policy observations: six operators, six text
values, both operand orders, literal/reference forms and both date epochs. The
original encoding gives 160 to 272 matches; the explicitly preserved variant
gives 160 to 288. Forty separate control observations retain identical values
and types across the engines, including booleans, empty cells, text, errors,
criteria functions and arithmetic. This establishes the tested scope, not
universal Excel conformance. Broader upstream regression testing is still needed.

LibreOffice 26.2.4.2 independently recalculated copies of both encodings and
matched all 288 policy expectations for each encoding. These are the same
cases under two encodings, not 576 independent cases. This supports the tested
ordering and records a difference in whitespace import between the engines;
LibreOffice agreement is not a universal Microsoft Excel oracle.

## Date output and rejected adaptations

For `DATE(2026,2,1)-46054`, the tested engine's internal `ISNUMBER`, `TYPE` and
numeric comparison results indicate numeric arithmetic. Native output in
the 1900 epoch materializes zero as a date; its serial-output policy returns
`0.0`. The 1904 result remains `-1462.0`. This evidence distinguishes a
restitution/type issue from an established arithmetic error in that example.

A global serial-output policy was rejected because it also converts actual
source and formatted dates into numbers. Replacing constant names with
references to hidden helper cells was also rejected: it changes `ISREF`
results and introduces artificial dependencies. Neither adaptation provides
a general correction while preserving the original semantics.

## Decisions and next checks

Keep the installed dependency and its diagnostic evidence unchanged. Before
adoption, the Calamine proposal needs upstream review and broader importer
and engine regression tests, including unsupported definitions and numeric
boundaries. Qualified-name resolution, invoked `LAMBDA` and Umya's source
metadata loss require separate investigations. A date-output correction
requires explicit source-format handling and distinct raw/output evidence.
No dependency update or upstream publication is implied by this investigation.
