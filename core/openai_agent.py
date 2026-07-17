"""OpenAI chat-completions agent: the bot's LLM brain.

Covers the surface the bot uses — ``run`` (plain completion OR a tool-calling
loop), ``structured_output`` (JSON-schema constrained), and ``tool`` registration.
Tool-loop semantics: schema inferred from each callable's signature, a
max-iterations cap, tool results fed back as ``role: tool`` messages, and tool
exceptions surfaced to the model as an ``Error: ...`` string rather than crashing.

Model note: the GPT-5 family are reasoning models. Consequences baked in here:
requests use ``max_completion_tokens`` (they reject the older ``max_tokens``
param), reasoning tokens draw from that same budget (so pass a generous cap or the
visible answer/tool call can come back empty), and ``temperature`` is never sent
(reasoning models reject non-default values).
"""

from __future__ import annotations

import inspect
import json
import logging
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from .errors import AgentError

if TYPE_CHECKING:
    from openai import OpenAI

logger = logging.getLogger(__name__)

# Current strong chat model, verified against the live /v1/models list
# (snapshot gpt-5.5-2026-04-23). Consumers pass their own per-bot model, so this
# default only guards direct construction.
DEFAULT_MODEL = "gpt-5.5"
DEFAULT_MAX_TOKENS = 4096
MAX_TOOL_ITERATIONS = 10

_PY_TO_JSON = {
    str: "string",
    int: "integer",
    float: "number",
    bool: "boolean",
    dict: "object",
    list: "array",
}


@dataclass
class Tool:
    """A registered callable exposed to the model."""

    name: str
    description: str
    parameters: dict[str, Any]
    fn: Callable[..., Any]


def _schema_from_signature(fn: Callable[..., Any]) -> dict[str, Any]:
    """Build a minimal JSON Schema from a callable's annotations.

    Required = params without a default. Unannotated params default to string.
    (Identical signature inference so a callable registers the same way
    on the agent.)
    """
    sig = inspect.signature(fn)
    properties: dict[str, Any] = {}
    required: list[str] = []
    for pname, param in sig.parameters.items():
        if pname == "self" or param.kind in (
            inspect.Parameter.VAR_POSITIONAL,
            inspect.Parameter.VAR_KEYWORD,
        ):
            continue
        json_type = _PY_TO_JSON.get(param.annotation, "string")
        properties[pname] = {"type": json_type}
        if param.default is inspect.Parameter.empty:
            required.append(pname)
    return {"type": "object", "properties": properties, "required": required}


@dataclass
class OpenAIAgent:
    """A minimal OpenAI chat agent: plain, tool-calling, and structured completions.

    Constructor shape is stable so consumers build it the
    same way (``OpenAIAgent(client, system=..., model=..., max_tokens=...)``).

    Args:
        client: An ``openai.OpenAI`` instance (sync).
        system: System prompt (sent as the leading ``system`` message).
        model: Chat model id (default :data:`DEFAULT_MODEL`).
        max_tokens: Per-response cap, sent as ``max_completion_tokens`` — must
            leave headroom for reasoning tokens on GPT-5 models.
        keep_history: When True, ``run`` persists the full turn sequence (incl.
            tool calls/results) so successive calls share context. When False
            (default), each ``run`` is fresh.
    """

    client: OpenAI
    system: str = ""
    model: str = DEFAULT_MODEL
    max_tokens: int = DEFAULT_MAX_TOKENS
    keep_history: bool = False
    _tools: dict[str, Tool] = field(default_factory=dict, init=False)
    _history: list[dict[str, Any]] = field(default_factory=list, init=False)
    last_usage: dict[str, int] = field(default_factory=dict, init=False)

    def tool(
        self,
        fn: Callable[..., Any] | None = None,
        *,
        name: str | None = None,
        description: str | None = None,
    ) -> Callable[..., Any]:
        """Register a callable as a tool. Usable as ``@agent.tool`` or
        ``agent.tool(fn)``. Schema is inferred from the signature; the docstring
        becomes the description unless one is given.
        """

        def register(target: Callable[..., Any]) -> Callable[..., Any]:
            tool_name = name or target.__name__
            tool_desc = description or (inspect.getdoc(target) or tool_name)
            self._tools[tool_name] = Tool(
                name=tool_name,
                description=tool_desc,
                parameters=_schema_from_signature(target),
                fn=target,
            )
            return target

        return register(fn) if fn is not None else register

    def reset(self) -> None:
        """Clear the conversation buffer (start a fresh session)."""
        self._history.clear()

    # ---- request assembly ----------------------------------------------

    def _conversation_start(self) -> list[dict[str, Any]]:
        messages: list[dict[str, Any]] = []
        if self.system:
            messages.append({"role": "system", "content": self.system})
        if self.keep_history:
            messages.extend(self._history)
        return messages

    def _tool_params(self) -> list[dict[str, Any]] | None:
        if not self._tools:
            return None
        return [
            {
                "type": "function",
                "function": {
                    "name": t.name,
                    "description": t.description,
                    "parameters": t.parameters,
                },
            }
            for t in self._tools.values()
        ]

    def _record_usage(self, response: Any) -> None:
        usage = getattr(response, "usage", None)
        if usage is None:
            return
        self.last_usage = {
            "input_tokens": getattr(usage, "prompt_tokens", 0) or 0,
            "output_tokens": getattr(usage, "completion_tokens", 0) or 0,
        }
        logger.info(
            "openai agent usage: input=%s output=%s",
            self.last_usage["input_tokens"],
            self.last_usage["output_tokens"],
        )

    def _create(self, messages: list[dict[str, Any]], **extra: Any) -> Any:
        try:
            return self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                max_completion_tokens=self.max_tokens,
                **extra,
            )
        except Exception as exc:  # network/API/schema error -> domain error
            raise AgentError(f"openai request failed: {exc}") from exc

    def _execute_tool(self, name: str, arguments_json: str | None) -> str:
        tool = self._tools.get(name)
        if tool is None:
            return f"Error: unknown tool {name}"
        try:
            args = json.loads(arguments_json) if arguments_json else {}
        except json.JSONDecodeError:
            args = {}
        try:
            result = tool.fn(**args)
        except Exception as exc:  # surfaced to the model as a tool error, not a crash
            logger.exception("tool %s failed", name)
            return f"Error: {exc}"
        return result if isinstance(result, str) else json.dumps(result, default=str)

    # ---- public surface -------------------------------------------------

    def run(self, prompt: str, *, max_iterations: int = MAX_TOOL_ITERATIONS) -> str:
        """Run the model→tool→model loop until it stops calling tools.

        Returns the final assistant text. With no tools registered this is a plain
        single completion. With ``keep_history=True`` the turn sequence (incl. tool
        calls/results) persists for the next call.

        Raises:
            AgentError: on an API failure, or if the loop exceeds ``max_iterations``.
        """
        messages = self._conversation_start()
        messages.append({"role": "user", "content": prompt})
        tools = self._tool_params()
        extra: dict[str, Any] = {"tools": tools} if tools is not None else {}

        for _ in range(max_iterations):
            response = self._create(messages, **extra)
            self._record_usage(response)
            message = response.choices[0].message
            tool_calls = list(getattr(message, "tool_calls", None) or [])

            assistant_msg: dict[str, Any] = {"role": "assistant", "content": message.content}
            if tool_calls:
                assistant_msg["tool_calls"] = [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {"name": tc.function.name, "arguments": tc.function.arguments},
                    }
                    for tc in tool_calls
                ]
            messages.append(assistant_msg)

            if not tool_calls:
                text = (message.content or "").strip()
                if self.keep_history:
                    self._history = [m for m in messages if m.get("role") != "system"]
                return text

            for tc in tool_calls:
                output = self._execute_tool(tc.function.name, tc.function.arguments)
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": output})

        raise AgentError(f"tool loop exceeded {max_iterations} iterations")

    def structured_output(self, prompt: str, schema: dict[str, Any]) -> dict[str, Any]:
        """One-shot call constrained to a JSON Schema; returns the parsed dict.

        Uses ``response_format`` json_schema in strict mode (no tool loop, no
        history). The schema must be strict-compatible (``additionalProperties:
        false`` and every property listed in ``required``).

        Raises:
            AgentError: if the API call fails or the response is not valid JSON.
        """
        messages: list[dict[str, Any]] = []
        if self.system:
            messages.append({"role": "system", "content": self.system})
        messages.append({"role": "user", "content": prompt})
        response = self._create(
            messages,
            response_format={
                "type": "json_schema",
                "json_schema": {"name": "structured_output", "strict": True, "schema": schema},
            },
        )
        self._record_usage(response)
        text = response.choices[0].message.content or ""
        try:
            return json.loads(text)
        except json.JSONDecodeError as exc:
            raise AgentError(f"structured_output returned non-JSON: {exc}") from exc
