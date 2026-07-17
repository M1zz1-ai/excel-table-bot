"""Configuration loading from a single ``.env`` file.

Config is declared per-consumer via ``load(required=[...])`` — the core never
hard-requires a fixed set, so a new consumer adds keys without touching this
module. Required keys fail loud (ConfigError naming the missing key) so a
misconfigured deploy never silently runs with empty credentials.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from dotenv import dotenv_values

from .errors import ConfigError

# Default env file: `.env` in the current working directory (see .env.example).
MASTER_ENV_PATH = Path(".env")

# Defaults for optional keys. A bot may still declare REDIS_URL as required;
# if absent from the env file it falls back to this rather than raising.
DEFAULTS: dict[str, str] = {
    "REDIS_URL": "redis://localhost:6379",
}


@dataclass(frozen=True)
class Config:
    """Loaded config values. Access via attribute (``cfg.openai_api_key``)
    or by raw env key (``cfg.get("TELEGRAM_BOT_TOKEN")``).

    Only the keys a bot requested (plus any present defaults) are populated.
    """

    values: dict[str, str] = field(default_factory=dict)

    def get(self, key: str, default: str | None = None) -> str | None:
        """Return the value for an env key, or ``default`` if absent."""
        return self.values.get(key, default)

    def require(self, key: str) -> str:
        """Return the value for an env key or raise ConfigError if absent."""
        value = self.values.get(key)
        if value is None:
            raise ConfigError(f"Missing required config key: {key}")
        return value

    def __getattr__(self, name: str) -> str:
        """Attribute access by lower-snake env key (e.g. ``cfg.fal_key``)."""
        key = name.upper()
        value = self.values.get(key)
        if value is None:
            raise AttributeError(name)
        return value


def load(required: list[str], *, env_path: Path = MASTER_ENV_PATH) -> Config:
    """Load and validate the env file for one bot's declared keys.

    Does not mutate ``os.environ`` (uses ``dotenv_values``), so loading config
    for a ``--check`` never leaks secrets into the process environment.

    A required key is satisfied by a non-empty value in the env file, or by an
    entry in :data:`DEFAULTS` (e.g. ``REDIS_URL``).

    Args:
        required: Env keys this bot needs (e.g. ["TELEGRAM_BOT_TOKEN", "FAL_KEY"]).
        env_path: Override for the master env file (tests pass a tmp path).

    Raises:
        ConfigError: if any required key is absent/empty and has no default.
    """
    raw = dict(dotenv_values(env_path))
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for key in required:
        value = raw.get(key)
        if value is not None and value.strip() != "":
            resolved[key] = value.strip()
        elif key in DEFAULTS:
            resolved[key] = DEFAULTS[key]
        else:
            missing.append(key)

    if missing:
        raise ConfigError(f"Missing required config key(s): {', '.join(missing)}")

    return Config(values=resolved)
