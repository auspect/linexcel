# Optional local Phoenix export

The application's own question trace is the source of truth. Phoenix is an
optional, best-effort copy for debugging timing, tool use and token consumption.
Neither the extra nor a collector is required for search or questions.

Start a **user-run** Phoenix in its own environment (the app does not start it):

```powershell
uvx --python 3.13 arize-phoenix serve
```

Then start linexcel with the optional extra and an explicit OTLP/HTTP endpoint:

```powershell
$env:LINEXCEL_PHOENIX_ENDPOINT = 'http://127.0.0.1:6006/v1/traces'
uv run --extra ai --extra tracing linexcel serve
```

The exporter owns its tracer provider; it never replaces the process-global
provider or instruments an SDK automatically. Each run has a trace ID in its
local log. A root `AGENT` span parents `LLM`, `TOOL` and `RETRIEVER` steps using
OpenInference conventions. The fixed allowlist contains tool names, statuses,
durations, numeric counts and token counts. There is **no content capture mode**:
no questions, cell values, formulas, labels, filenames, handles, arguments,
provider credentials or headers are sent. Identifiers, timing and counts remain
metadata and can reveal usage patterns.

The queue holds at most 256 spans, sends batches of 32, and drops excess work.
Exports use a one-second HTTP timeout, no retries, no redirects and no environment
proxy/header configuration. Missing dependencies, invalid configuration or a down
collector cannot change a question's result. `exportState` describes configuration,
not proof of delivery. Inspect Phoenix to verify delivery.

Only literal loopback IPs and `localhost` are accepted by default. Export to a
non-loopback collector requires `LINEXCEL_PHOENIX_ALLOW_REMOTE=1`. This is separate
from consent to send workbook data to a remote model. URL credentials, queries and
fragments are refused. Keep the collector user-controlled and secure its retention.

To add a link in the app's trace panel, set `LINEXCEL_PHOENIX_TRACE_URL` to the
trace URL for your Phoenix project with `{trace_id}` as its only placeholder, for
example `http://127.0.0.1:6006/projects/UHJvamVjdDoy/traces/{trace_id}`. Use the
project ID returned by your collector, rather than copying this example ID.

See [Phoenix OTLP setup](https://arize.com/docs/phoenix/tracing/how-to-tracing/setup-tracing/setup-using-phoenix-otel).
