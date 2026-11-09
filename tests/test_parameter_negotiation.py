"""Real loopback HTTP server and SDK; no paid service or model required."""

import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pytest

from linexcel.aidoc import AiDocError, OpenAICompatProvider, estimate_tokens

pytest.importorskip("openai")


@pytest.fixture
def endpoint():
    bodies = []
    rejected = []

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def do_POST(self):
            body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
            bodies.append(body)
            parameter = next((p for p in rejected if p in body), None)
            if parameter:
                self.send_response(400)
                result = {
                    "error": {
                        "message": f"Unsupported parameter: '{parameter}'.",
                        "param": parameter,
                        "type": "invalid_request_error",
                    }
                }
                if "__string_error__" in rejected:
                    result = {"error": result["error"]["message"]}
            else:
                self.send_response(200)
                result = dict(
                    id="local-response",
                    object="chat.completion",
                    created=0,
                    model=body["model"],
                    choices=[
                        dict(
                            index=0,
                            finish_reason="stop",
                            message=dict(
                                role="assistant", content="Workbook label: Impôt."
                            ),
                        )
                    ],
                    usage=dict(prompt_tokens=7, completion_tokens=3, total_tokens=10),
                )
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(result).encode())

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_port}/v1", bodies, rejected
    finally:
        server.shutdown()
        server.server_close()
        thread.join()


@pytest.mark.parametrize(
    "parameters", [("temperature",), ("max_tokens",), ("max_tokens", "temperature")]
)
@pytest.mark.parametrize("vision", [False, True])
def test_explicit_rejections_retry_and_record_actual_parameters(
    endpoint, parameters, vision
):
    base, bodies, rejected = endpoint
    rejected.extend(parameters)
    with OpenAICompatProvider(model="arbitrary-model", base_url=base) as provider:
        if vision:
            _, usage = provider.generate_with_image(
                "system", "user", b"image", temperature=0, max_tokens=42
            )
        else:
            _, usage = provider.generate_with_usage("system", "user", max_tokens=42)
        receipt = provider.vision_receipt if vision else provider.receipt
    assert len(bodies) == len(parameters) + 1
    assert usage.requests == len(bodies)
    assert usage.input_tokens == 7 + len(parameters) * estimate_tokens("system\n\nuser")
    assert usage.estimated
    assert receipt["status"] == "complete"
    assert receipt["parameters"] == {
        k: bodies[-1][k]
        for k in ("temperature", "max_tokens", "max_completion_tokens")
        if k in bodies[-1]
    }
    assert receipt["temperature_omitted"] == ("temperature" in parameters)
    assert receipt["token_parameter"] == (
        "max_completion_tokens" if "max_tokens" in parameters else "max_tokens"
    )
    assert len(receipt["request_attempts"]) == len(bodies)
    assert receipt["request_attempts"][0]["status"] == "failed"
    # Reuse across instances and documentation/vision, but not another model.
    with OpenAICompatProvider(model="arbitrary-model", base_url=base) as provider:
        provider.generate_with_image("system", "user", b"image", max_tokens=42)
        assert len(provider.vision_receipt["request_attempts"]) == 1
    assert ("temperature" not in bodies[-1]) == ("temperature" in parameters)
    assert bodies[-1].get("max_completion_tokens", bodies[-1].get("max_tokens")) == 42
    with OpenAICompatProvider(model="another-model", base_url=base) as provider:
        _, usage = provider.generate_with_usage("system", "user", max_tokens=42)
        assert usage.requests == len(parameters) + 1


def test_unrecognized_400_does_not_retry_and_preserves_usage(endpoint):
    base, bodies, rejected = endpoint
    rejected.append("messages")
    with OpenAICompatProvider(model="any", base_url=base) as provider:
        with pytest.raises(AiDocError) as caught:
            provider.generate_with_usage("system", "user", max_tokens=42)
        assert len(bodies) == 1
        assert caught.value.usage.requests == 1


def test_validation_receipt_records_actual_parameters(endpoint, manual):
    base, bodies, rejected = endpoint
    rejected.extend(["temperature", "max_tokens"])
    receipts = []
    with manual.AcceptanceProvider(
        model="receipt-model", base_url=base, response_receipts=receipts
    ) as provider:
        provider.generate("system", "user", max_tokens=42)
    assert receipts[0]["status"] == "complete"
    assert receipts[0]["parameters"] == {"max_completion_tokens": 42}
    assert receipts[0]["temperature_omitted"] is True
    assert receipts[0]["usage"]["requests"] == len(bodies) == 3


def test_compatible_endpoint_can_return_string_error_body(endpoint):
    base, bodies, rejected = endpoint
    rejected.extend(["__string_error__", "temperature", "max_tokens"])
    with OpenAICompatProvider(model="string-error", base_url=base) as provider:
        _, usage = provider.generate_with_usage("system", "user", max_tokens=42)
        assert provider.receipt["parameters"] == {"max_completion_tokens": 42}
    assert usage.requests == len(bodies) == 3


pytest_plugins = ["test_manual_validation"]
