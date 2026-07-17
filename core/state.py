"""Redis-backed state: dedup ledger + session store, with graceful fallback.

Redis is live locally (``redis-cli ping`` -> PONG). When it's down we log once
and degrade to no-op rather than crashing — a bot that can't dedup is degraded,
not dead (resilience applied to state).

- Dedup ledger: a per-namespace seen-set with TTL, so the same item isn't
  processed twice (a lightweight processed-id ledger).
- Session store: get/set arbitrary JSON-able state under a key with TTL
  (conversation/session memory for chat bots).
"""

from __future__ import annotations

import json
import logging
from typing import Any

import redis

logger = logging.getLogger(__name__)

DEFAULT_URL = "redis://localhost:6379"
DEFAULT_SEEN_TTL = 48 * 3600  # seconds (48h dedup window)
DEFAULT_SESSION_TTL = 3600


class RedisState:
    """Dedup + session store. All methods are safe when redis is unreachable.

    Args:
        url: redis connection URL (e.g. ``redis://localhost:6379``).
        namespace: key prefix isolating one bot's state from another's.
        client: inject a client/fake for tests; otherwise built from ``url``.
    """

    def __init__(
        self,
        url: str = DEFAULT_URL,
        *,
        namespace: str = "excel",
        client: Any | None = None,
    ) -> None:
        self.namespace = namespace
        self._client = (
            client if client is not None else redis.Redis.from_url(url, decode_responses=True)
        )
        self._warned = False

    def _key(self, *parts: str) -> str:
        return ":".join((self.namespace, *parts))

    def _degraded(self, exc: Exception) -> None:
        """Log the first redis failure; stay quiet after to avoid log spam."""
        if not self._warned:
            logger.warning("redis unavailable, degrading to no-op: %s", exc)
            self._warned = True

    def ping(self) -> bool:
        """True if redis answers, False if unreachable (never raises)."""
        try:
            return bool(self._client.ping())
        except redis.RedisError as exc:
            self._degraded(exc)
            return False

    # ---- dedup ----------------------------------------------------------

    def seen(self, item_id: str, *, ledger: str = "seen") -> bool:
        """True if ``item_id`` was already marked in this ledger.

        On redis failure returns False (treat as unseen) so the bot still
        processes the item rather than silently dropping it.
        """
        try:
            return bool(self._client.exists(self._key(ledger, item_id)))
        except redis.RedisError as exc:
            self._degraded(exc)
            return False

    def mark_seen(self, item_id: str, *, ledger: str = "seen", ttl: int = DEFAULT_SEEN_TTL) -> None:
        """Mark ``item_id`` seen with a TTL. No-op on redis failure."""
        try:
            self._client.set(self._key(ledger, item_id), "1", ex=ttl)
        except redis.RedisError as exc:
            self._degraded(exc)

    # ---- session store --------------------------------------------------

    def set_session(self, key: str, value: Any, *, ttl: int = DEFAULT_SESSION_TTL) -> None:
        """Store a JSON-able value under a session key with TTL. No-op on failure."""
        try:
            self._client.set(self._key("session", key), json.dumps(value), ex=ttl)
        except redis.RedisError as exc:
            self._degraded(exc)

    def get_session(self, key: str, default: Any = None) -> Any:
        """Read a session value (parsed from JSON). Returns ``default`` on miss
        or redis failure.
        """
        try:
            raw = self._client.get(self._key("session", key))
        except redis.RedisError as exc:
            self._degraded(exc)
            return default
        if raw is None:
            return default
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return default

    def clear_session(self, key: str) -> None:
        """Delete a session key. No-op on failure."""
        try:
            self._client.delete(self._key("session", key))
        except redis.RedisError as exc:
            self._degraded(exc)
