"""Tool calling is separate from the documentation prompt-to-text protocol."""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Protocol

from linexcel.aidoc import OpenAICompatProvider, TokenUsage


@dataclass(frozen=True)
class ToolResponse:
    calls: list[dict]
    finish_reason: str
    usage: TokenUsage


class ToolProvider(Protocol):
    def complete(
        self, messages: list[dict], tools: list[dict], *, tokens: int, seconds: float
    ) -> ToolResponse: ...

    def close(self) -> None: ...


class ToolCapabilityError(RuntimeError):
    """Configured endpoint explicitly rejected structured tool calling."""

    usage: TokenUsage | None = None


class OpenAIToolProvider:
    """Non-streaming structured calls, using M0's client/parameter negotiation."""

    def __init__(self, config):
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
            if getattr(cause, "status_code", None) in {400, 422} and any(
                word in str(cause).lower()
                for word in ("tools", "tool_choice", "function calling")
            ):
                error = ToolCapabilityError("Tool calling is not supported")
                error.usage = getattr(exc, "usage", None)
                raise error from exc
            raise

    def _request(self, messages, tools, *, tokens, seconds):
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
            [call.model_dump() for call in (choice.message.tool_calls or [])],
            choice.finish_reason,
            usage,
        )

    def close(self):
        self.client.close()
