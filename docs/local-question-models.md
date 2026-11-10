# Local tool models and failure diagnostics

The question loop requires structured tool calls. A plain-text answer produces
`tool_calls_required`; linexcel does not convert it into a factual answer. A
`length` finish reason produces `truncated_or_unsupported_calls`, even if a partial
tool call exists. The model trace records the finish reason and content/reasoning
byte counts, without retaining that text. Failed HTTP calls record a bounded
HTTP status and timeout flag when available through the exception chain.

Cell handles are opaque source references and must be copied exactly from tool
results. An unissued handle in `get_cell`, `trace_lineage` or `compute` fails with
`invalid_source_handle`, without displaying a proposal. This is an invalid model
reference; it does not establish that the workbook or question scope changed.
An invalid terminal proposal still returns `not_found`, with its rejection in
the trace and no displayed answer.

The cumulative token allowance in `AskLimits` is **not the model context window**.
Each successive request includes history, tool schemas and source evidence. A
local runtime with a smaller context may truncate history, exhaust generation
space, or reject the request before the application's cumulative budget is used.
Reasoning also consumes generation space and request time. Do not treat raising
the application budget as a correctness fix.

For Ollama, inspect `ollama ps` (or `/api/ps`) while the model is loaded. Configure
the runtime's context explicitly for your hardware; the OpenAI-compatible API's
`max_tokens` controls generation, not context allocation. See
[Ollama context length](https://docs.ollama.com/context-length) and
[OpenAI compatibility](https://docs.ollama.com/api/openai-compatibility).
Check that the installed model really supports tools and that its template is
appropriate. Provider configuration and thinking controls are model-specific;
linexcel exposes `--ai-reasoning-effort` and the existing template-thinking
options, but does not silently turn thinking off or repair malformed calls.

The November 2026 local Gemma diagnostic used `gemma4:12b`. A replay verified an
8,192-token runtime context, with one request reporting 8,023 prompt tokens plus
169 completion tokens. Earlier failures reported `length`, including 8,024 plus
167 tokens. This is evidence of the context boundary, not a test of an adequately
configured long-context model. Other length failures can consume reasoning space.
The historical plain-text reply had `stop` and no calls. The historical generic
`AiDocError` retained no underlying provider payload; its precise cause cannot
be established retroactively. A replay succeeded, so it is not labelled fixed.
New diagnostics preserve content-free failure distinctions for future runs.

These diagnostics do not establish model reliability. Compare a frozen question
set against a deterministic catalog baseline, keep failures in denominators,
and distinguish correct source facts from selection of the right cell.

For the official OpenAI endpoint, GPT-6 tool models configured with reasoning
use the Responses API. Medium reasoning with Luna is not supported for function
calling in Chat Completions; the adapter preserves the selected effort rather
than reducing it. It uses `store=false`, disabled truncation, and encrypted
reasoning history retained only in the provider's run memory. Compatible local
endpoints continue to use Chat Completions. See the
[official Luna model documentation](https://developers.openai.com/api/docs/models/gpt-6-luna)
and [function calling guide](https://developers.openai.com/api/docs/guides/function-calling).
