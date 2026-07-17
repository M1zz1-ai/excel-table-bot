"""Shared exception types + resilience.

Resilience is the design goal of this bot: a failed work
cycle is logged, optionally alerted to Telegram, and the process KEEPS RUNNING
so the next cycle retries. A single transient failure (network blip, rate
limit, expired token) never kills a bot.
"""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Awaitable, Callable
from functools import wraps
from typing import Any, Protocol, TypeVar

logger = logging.getLogger(__name__)

T = TypeVar("T")


# ---- shared exception types --------------------------------------------


class CoreError(RuntimeError):
    """Base for all core errors."""


class ConfigError(CoreError):
    """Raised when a required configuration key is missing or empty."""


class FalError(CoreError):
    """Raised when a fal.ai request fails."""


class NotionError(CoreError):
    """Raised when a notion-cli subprocess call fails."""


class SheetError(CoreError):
    """Raised when a spreadsheet operation fails."""


class SttError(CoreError):
    """Raised when a speech-to-text transcription request fails."""


class AgentError(CoreError):
    """Raised when the LLM agent loop fails irrecoverably."""


# ---- alert sink --------------------------------------------------------


class Alerter(Protocol):
    """Anything that can deliver a one-line failure alert (e.g. TelegramClient)."""

    async def send_text(self, text: str, chat_id: int | None = ...) -> Any: ...


async def _send_alert(alerter: Alerter | None, message: str) -> None:
    """Best-effort alert; an alert failure must never mask the original error."""
    if alerter is None:
        return
    try:
        await alerter.send_text(message)
    except Exception:
        logger.exception("failed to send failure alert")


# ---- resilient execution ----------------------------------------------


async def run_resilient[T](
    work: Callable[[], Awaitable[T]],
    *,
    alerter: Alerter | None = None,
    label: str = "cycle",
    alert_prefix: str = "⚠️",
) -> T | None:
    """Run one work cycle, swallowing any exception so the caller stays alive.

    Returns the work result on success, or None if the cycle raised. The
    exception is logged with a traceback and, if an alerter is given, surfaced
    to Telegram. ``asyncio.CancelledError`` is re-raised (cancellation is not a
    failure to swallow).

    Args:
        work: A zero-arg coroutine factory (call it fresh each cycle).
        alerter: Optional Telegram client to ping on failure.
        label: Human label for logs/alerts (e.g. "poll").
        alert_prefix: Emoji/prefix for the alert line.
    """
    try:
        return await work()
    except asyncio.CancelledError:
        raise
    except Exception as exc:
        logger.exception("%s failed; continuing", label)
        await _send_alert(alerter, f"{alert_prefix} {label} failed: {exc}")
        return None


def resilient(
    *, alerter: Alerter | None = None, label: str | None = None, alert_prefix: str = "⚠️"
) -> Callable[[Callable[..., Awaitable[T]]], Callable[..., Awaitable[T | None]]]:
    """Decorator form of :func:`run_resilient` for an async function.

    The wrapped call swallows exceptions (logs + optional alert) and returns
    None on failure. ``CancelledError`` still propagates.
    """

    def decorate(fn: Callable[..., Awaitable[T]]) -> Callable[..., Awaitable[T | None]]:
        fn_label = label or fn.__name__

        @wraps(fn)
        async def wrapper(*args: Any, **kwargs: Any) -> T | None:
            return await run_resilient(
                lambda: fn(*args, **kwargs),
                alerter=alerter,
                label=fn_label,
                alert_prefix=alert_prefix,
            )

        return wrapper

    return decorate
