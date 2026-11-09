"""Optional export cannot change question outcomes or disclose content."""

import json

import pytest

from linexcel.observability import QuestionTracing, checked_url


@pytest.mark.parametrize(
    "url",
    [
        "https://remote.test/v1/traces",
        "http://user:secret@localhost/v1/traces",
        "http://localhost/v1/traces?key=secret",
    ],
)
def test_export_url_refuses_remote_and_credentials(url):
    with pytest.raises(ValueError):
        checked_url(url)


def test_disabled_and_invalid_export_are_inert(monkeypatch):
    monkeypatch.delenv("LINEXCEL_PHOENIX_ENDPOINT", raising=False)
    tracer = QuestionTracing()
    assert tracer.state == "disabled"
    run = tracer.start()
    run.step("model", 1, "completed")
    run.end("failed")
    assert len(run.id) == 32
    monkeypatch.setenv("LINEXCEL_PHOENIX_ENDPOINT", "https://remote.test/v1/traces")
    assert QuestionTracing().state == "configuration_refused"


def test_content_free_hierarchy_and_failure_isolation(monkeypatch):
    pytest.importorskip("opentelemetry.sdk")
    from opentelemetry import trace
    from opentelemetry.sdk.trace.export import SimpleSpanProcessor
    from opentelemetry.sdk.trace.export.in_memory_span_exporter import (
        InMemorySpanExporter,
    )

    global_provider = trace.get_tracer_provider()
    monkeypatch.setenv("LINEXCEL_PHOENIX_ENDPOINT", "http://127.0.0.1:1/v1/traces")
    tracer = QuestionTracing()
    assert tracer.state == "enabled"
    memory = InMemorySpanExporter()
    tracer.provider.add_span_processor(SimpleSpanProcessor(memory))
    run = tracer.start()
    run.step(
        "search_catalog",
        2,
        "completed",
        counts={"results": 3, "SECRET_CELL": "SECRET_VALUE"},
    )
    run.step(
        "model",
        3,
        "completed",
        tokens={"input_tokens": 10, "output_tokens": 4, "key": "SECRET_KEY"},
    )
    run.step("SECRET_QUESTION", 1, "completed")
    run.end("answer_proposed")
    tracer.provider.force_flush(timeout_millis=2000)
    spans = memory.get_finished_spans()
    assert len(spans) == 3
    root = next(s for s in spans if s.name == "question")
    assert root.attributes["openinference.span.kind"] == "AGENT"
    assert all(s.parent.span_id == root.context.span_id for s in spans if s is not root)
    assert all(s.context.trace_id == int(run.id, 16) for s in spans)
    assert "SECRET" not in json.dumps([dict(s.attributes) for s in spans])
    assert trace.get_tracer_provider() is global_provider
    tracer.close()
