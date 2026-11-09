"""Optional, application-owned, content-free OpenInference OTLP export.

No global provider, instrumentation, environment headers or content capture.
The local question log remains authoritative; export is best effort.
"""

from __future__ import annotations

import ipaddress
import os
import time
import uuid
from urllib.parse import urlsplit

TOOLS = frozenset(
    {
        "list_workbooks",
        "list_revisions",
        "search_catalog",
        "describe_sheet",
        "read_range",
        "get_cell",
        "find_value",
        "trace_lineage",
        "list_external_links",
        "compute",
        "submit_answer",
        "model",
        "failure",
        "finish",
    }
)


def checked_url(value, allow_remote=False):
    parsed = urlsplit(value)
    if (
        parsed.scheme not in {"http", "https"}
        or not parsed.hostname
        or parsed.username
        or parsed.password
        or parsed.query
        or parsed.fragment
        or any(ord(c) < 33 for c in value)
    ):
        raise ValueError("Invalid tracing URL")
    _ = parsed.port
    try:
        local = ipaddress.ip_address(parsed.hostname).is_loopback
    except ValueError:
        local = parsed.hostname.lower() == "localhost"
    if not local and not allow_remote:
        raise ValueError("Remote tracing requires LINEXCEL_PHOENIX_ALLOW_REMOTE=1")
    return value


class QuestionTracing:
    def __init__(self):
        self.provider = self.tracer = None
        self.link = None
        self.state = "disabled"
        endpoint = os.getenv("LINEXCEL_PHOENIX_ENDPOINT")
        if not endpoint:
            return
        try:
            allow = os.getenv("LINEXCEL_PHOENIX_ALLOW_REMOTE") == "1"
            checked_url(endpoint, allow)
            link = os.getenv("LINEXCEL_PHOENIX_TRACE_URL")
            if link:
                if link.count("{trace_id}") != 1 or "{" in link.replace(
                    "{trace_id}", ""
                ):
                    raise ValueError("Trace URL must contain only {trace_id}")
                checked_url(link.replace("{trace_id}", "0" * 32), allow)
            import requests
            from opentelemetry.exporter.otlp.proto.common.trace_encoder import (
                encode_spans,
            )
            from opentelemetry.sdk.resources import Resource
            from opentelemetry.sdk.trace import TracerProvider
            from opentelemetry.sdk.trace.export import (
                BatchSpanProcessor,
                SpanExporter,
                SpanExportResult,
            )

            class Exporter(SpanExporter):
                def __init__(self):
                    self.session = requests.Session()
                    self.session.trust_env = False

                def export(self, spans):
                    try:
                        response = self.session.post(
                            endpoint,
                            data=encode_spans(spans).SerializeToString(),
                            headers={"Content-Type": "application/x-protobuf"},
                            timeout=1,
                            allow_redirects=False,
                        )
                        response.close()
                        return (
                            SpanExportResult.SUCCESS
                            if 200 <= response.status_code < 300
                            else SpanExportResult.FAILURE
                        )
                    except Exception:
                        return SpanExportResult.FAILURE

                def shutdown(self):
                    self.session.close()

            self.provider = TracerProvider(
                resource=Resource(
                    {
                        "service.name": "linexcel",
                        "openinference.project.name": "linexcel",
                    }
                ),
                shutdown_on_exit=False,
            )
            self.provider.add_span_processor(
                BatchSpanProcessor(
                    Exporter(),
                    max_queue_size=256,
                    max_export_batch_size=32,
                    schedule_delay_millis=500,
                    export_timeout_millis=1500,
                )
            )
            self.tracer = self.provider.get_tracer("linexcel.questions")
            self.link, self.state = link, "enabled"
        except ImportError:
            self.state = "extra_unavailable"
        except Exception:
            self.state = "configuration_refused"

    def start(self):
        return RunTrace(self)

    def close(self):
        if self.provider:
            try:
                self.provider.shutdown()
            except Exception:
                pass


class RunTrace:
    def __init__(self, owner):
        self.owner, self.root = owner, None
        self.id = uuid.uuid4().hex
        if owner.tracer:
            try:
                from opentelemetry.context import Context

                self.root = owner.tracer.start_span(
                    "question",
                    context=Context(),
                    attributes={"openinference.span.kind": "AGENT"},
                )
                self.id = format(self.root.get_span_context().trace_id, "032x")
            except Exception:
                pass

    def step(self, name, duration_ms, verdict, tokens=None, counts=None):
        if not self.root or name not in TOOLS:
            return
        try:
            from opentelemetry.trace import set_span_in_context

            attrs = {
                "openinference.span.kind": "LLM"
                if name == "model"
                else "RETRIEVER"
                if name in {"search_catalog", "find_value"}
                else "TOOL",
                "linexcel.status": verdict
                if verdict
                in {
                    "completed",
                    "failed",
                    "answer_proposed",
                    "not_found",
                    "result_bytes",
                }
                else "other",
                "tool.name": name,
            }
            for key in ("input_tokens", "output_tokens"):
                value = (tokens or {}).get(key)
                if type(value) is int and value >= 0:
                    attrs[
                        "llm.token_count."
                        + {"input_tokens": "prompt", "output_tokens": "completion"}[key]
                    ] = value
            for key in ("cells", "results", "nodes", "edges", "revisions", "links"):
                value = (counts or {}).get(key)
                if type(value) is int and value >= 0:
                    attrs["linexcel.count." + key] = value
            end = time.time_ns()
            span = self.owner.tracer.start_span(
                name,
                context=set_span_in_context(self.root),
                attributes=attrs,
                start_time=end - int(max(0, duration_ms) * 1_000_000),
            )
            span.end(end_time=end)
        except Exception:
            pass

    def end(self, status):
        if self.root:
            try:
                self.root.set_attribute(
                    "linexcel.status",
                    status
                    if status in {"failed", "not_found", "answer_proposed"}
                    else "other",
                )
                self.root.end()
            except Exception:
                pass
