"""Tool calling is separate from the documentation prompt-to-text protocol."""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Protocol

from linexcel.aidoc import (
    AiDocError,
    OpenAICompatProvider,
    TokenUsage,
    _usage_from,
    estimate_tokens,
)


@dataclass(frozen=True)
class ToolResponse:
    calls: list[dict]
    finish_reason: str
    usage: TokenUsage
    diagnostics: dict | None = None


class ToolProvider(Protocol):
    def complete(
        self, messages: list[dict], tools: list[dict], *, tokens: int, seconds: float
    ) -> ToolResponse: ...

    def close(self) -> None: ...


class ToolCapabilityError(RuntimeError):
    """Configured endpoint explicitly rejected structured tool calling."""

    usage: TokenUsage | None = None


class QuestionConfigurationChangedError(RuntimeError):
    """The live configuration no longer matches this question's approval."""


class EndpointRedirectRefusedError(RuntimeError):
    usage: TokenUsage | None = None


class OpenAIToolProvider:
    """Non-streaming structured calls, using M0's client/parameter negotiation."""

    def __init__(self, config):
        # The official GPT-6 reasoning/tool protocol is Responses. Compatible
        # local runtimes keep their explicitly configured Chat Completions API.
        self.responses = (
            config.ai_base_url.rstrip("/") == "https://api.openai.com/v1"
            and config.ai_model.startswith("gpt-6")
            and config.ai_reasoning_effort not in {None, "none"}
        )
        self.reasoning_effort = config.ai_reasoning_effort
        self._response_history = []
        self._message_count = 0
        extra = {}
        if config.ai_thinking is not None:
            extra["thinking"] = {"type": config.ai_thinking}
        if config.ai_chat_template_thinking is not None:
            extra["chat_template_kwargs"] = {
                "enable_thinking": config.ai_chat_template_thinking == "enabled"
            }
        self.client = OpenAICompatProvider(
            base_url=config.ai_base_url,
            model=config.ai_model,
            api_key=config.ai_api_key or "not-needed",
            http_client_kwargs={"follow_redirects": False},
            request_kwargs={
                **({"extra_body": extra} if extra else {}),
                **(
                    {"reasoning_effort": config.ai_reasoning_effort}
                    if config.ai_reasoning_effort is not None
                    else {}
                ),
            },
        )

    def complete(self, messages, tools, *, tokens, seconds):
        try:
            return self._request(messages, tools, tokens=tokens, seconds=seconds)
        except Exception as exc:
            cause = exc.__cause__
            status = getattr(cause, "status_code", None)
            if isinstance(status, int) and 300 <= status < 400:
                error = EndpointRedirectRefusedError(
                    "Question endpoint redirect refused"
                )
                error.usage = getattr(exc, "usage", None)
                raise error from exc
            if getattr(cause, "status_code", None) in {400, 422} and any(
                word in str(cause).lower()
                for word in ("tools", "tool_choice", "function calling")
            ):
                error = ToolCapabilityError("Tool calling is not supported")
                error.usage = getattr(exc, "usage", None)
                raise error from exc
            raise

    def _request(self, messages, tools, *, tokens, seconds):
        if self.responses:
            return self._request_responses(
                messages, tools, tokens=tokens, seconds=seconds
            )
        response, usage = self.client._complete(
            messages,
            json.dumps({"messages": messages, "tools": tools}, ensure_ascii=False),
            {
                **self.client._request_options(),
                "tools": tools,
                "tool_choice": "required",
                "temperature": 0,
                "max_tokens": tokens,
                "timeout": seconds,
            },
            structured=True,
        )
        if len(response.choices) != 1:
            raise ValueError("Expected one tool-calling choice")
        choice = response.choices[0]
        return ToolResponse(
            [
                {
                    "id": call.id,
                    "type": call.type,
                    "function": {
                        "name": getattr(getattr(call, "function", None), "name", None),
                        "arguments": getattr(
                            getattr(call, "function", None), "arguments", None
                        ),
                    },
                }
                for call in (choice.message.tool_calls or [])
            ],
            choice.finish_reason,
            usage,
            {
                "contentBytes": len((choice.message.content or "").encode()),
                "reasoningBytes": len(
                    str(
                        getattr(choice.message, "reasoning", "")
                        or getattr(choice.message, "reasoning_content", "")
                        or ""
                    ).encode()
                ),
                "httpStatus": 200,
            },
        )

    def _complete_responses(self, request):
        try:
            response = self.client._client.responses.create(**request)
        except Exception as exc:
            raise AiDocError(
                "Responses tool request failed",
                usage=TokenUsage(
                    input_tokens=estimate_tokens(json.dumps(request)),
                    requests=1,
                    estimated=True,
                ),
            ) from exc
        usage = _usage_from(
            response.usage,
            ("input_tokens", "output_tokens"),
            json.dumps(request, ensure_ascii=False),
            response.model_dump_json(),
            model=self.client._model,
            provider="openai-responses",
        )
        return response, usage

    def _request_responses(self, messages, tools, *, tokens, seconds):
        # Stateless transport with encrypted reasoning retained only in memory.
        # Never use a remote conversation ID or implicit server-side history.
        for message in messages[self._message_count :]:
            if message["role"] in {"system", "user"}:
                self._response_history.append(
                    {"role": message["role"], "content": message["content"]}
                )
            elif message["role"] == "tool":
                self._response_history.append(
                    {
                        "type": "function_call_output",
                        "call_id": message["tool_call_id"],
                        "output": message["content"],
                    }
                )
        self._message_count = len(messages)
        response, usage = self._complete_responses(
            {
                "model": self.client._model,
                "input": self._response_history,
                "tools": [
                    {"type": "function", **item["function"], "strict": False}
                    for item in tools
                ],
                "tool_choice": "required",
                "store": False,
                "include": ["reasoning.encrypted_content"],
                "reasoning": {"effort": self.reasoning_effort},
                "max_output_tokens": tokens,
                "timeout": seconds,
                "truncation": "disabled",
            }
        )
        self._response_history.extend(
            item.model_dump(exclude_none=True) for item in response.output
        )
        calls = [
            {
                "id": item.call_id,
                "type": "function",
                "function": {"name": item.name, "arguments": item.arguments},
            }
            for item in response.output
            if item.type == "function_call"
        ]
        finish = "tool_calls" if calls else "stop"
        if response.status != "completed":
            finish = (
                "length"
                if getattr(response.incomplete_details, "reason", None)
                == "max_output_tokens"
                else "unsupported"
            )
        return ToolResponse(
            calls,
            finish,
            usage,
            {
                "contentBytes": len(response.output_text.encode()),
                "reasoningBytes": 0,
                "httpStatus": 200,
            },
        )

    def close(self):
        self.client.close()
