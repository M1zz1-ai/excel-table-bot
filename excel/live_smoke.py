"""Live end-to-end smoke test for the excel bot, gated on real credentials.

Drives the bot's OWN modules (session + tools + a real OpenAI agent) against
a tiny in-process .xlsx and the real Telegram API. Guard-railed and clearly
labelled: it uploads a synthetic workbook, asks the agent to read + edit it, and
sends the result to the owner chat. It never starts the long-poll loop.

Gating: when TELEGRAM_BOT_TOKEN_TABLE or
OPENAI_API_KEY are absent it SKIPS with a clear message and exits 0 — it never
invents creds.

Run:
  uv run python -m excel.live_smoke
"""

from __future__ import annotations

import asyncio
import logging
import tempfile
from pathlib import Path

import openai
from openpyxl import Workbook

from core import config, tg
from core.errors import ConfigError

from . import tools
from .tools import EXCEL_MODEL

logger = logging.getLogger("excel_bot.live_smoke")

# Keys whose absence means "skip, don't fail" — the gating credentials.
GATING_KEYS = ["TELEGRAM_BOT_TOKEN_TABLE", "OPENAI_API_KEY"]
# Keys needed to actually run the smoke once gated through.
SMOKE_KEYS = ["TELEGRAM_BOT_TOKEN_TABLE", "OPENAI_API_KEY", "TELEGRAM_CHAT_ID"]


def _present(key: str) -> bool:
    try:
        cfg = config.load([key], env_path=config.MASTER_ENV_PATH)
        return bool(cfg.get(key))
    except ConfigError:
        return False


def _gate() -> config.Config | None:
    """Return a loaded Config if gating creds are present, else None (skip)."""
    try:
        return config.load(SMOKE_KEYS, env_path=config.MASTER_ENV_PATH)
    except ConfigError as exc:
        missing = [k for k in GATING_KEYS if not _present(k)]
        if missing:
            print(
                f"SKIP — live smoke needs {', '.join(GATING_KEYS)} in "
                f".env (missing: {', '.join(missing)}). "
                "No real creds present; nothing to do."
            )
            return None
        print(f"Config error: {exc}")
        return None


def _make_sample(path: Path) -> None:
    """Write a tiny 'sales' workbook to drive the agent against."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["name", "qty", "status"])
    ws.append(["Apple", 10, "active"])
    ws.append(["Banana", 5, "active"])
    ws.append(["Cherry", 0, "inactive"])
    wb.save(path)


async def _run(cfg: config.Config) -> int:
    """Real E2E: agent reads + edits a sample workbook, file goes to the owner."""
    chat_id = tg.gather_chat_ids(cfg.require("TELEGRAM_CHAT_ID"))[0]
    telegram = tg.TelegramClient.from_token(cfg.require("TELEGRAM_BOT_TOKEN_TABLE"), chat_id)
    client = openai.OpenAI(api_key=cfg.require("OPENAI_API_KEY"))

    workdir = Path(tempfile.mkdtemp(prefix="excel_smoke_"))
    sample = workdir / "sample.xlsx"
    _make_sample(sample)

    from core import openai_agent as core_agent
    from core import sheets

    wb = sheets.load(sample)
    brain = core_agent.OpenAIAgent(client, system=tools.AGENT_SYSTEM, model=EXCEL_MODEL)
    tools.register_tools(brain, lambda: wb, on_change=lambda: sheets.save(wb, sample))

    failures = 0
    try:
        await telegram.send_text("🧪 <b>excel-bot LIVE E2E</b> — testing read + edit…")

        print("[..] step 1 — agent counts active rows", flush=True)
        answer = brain.run("Сколько строк со статусом active? Ответь числом и пояснением.")
        ok1 = bool(answer.strip())
        print(f"[{'PASS' if ok1 else 'FAIL'}] agent answered ({len(answer)} chars)", flush=True)
        failures += 0 if ok1 else 1

        print("[..] step 2 — agent appends a row", flush=True)
        brain2 = core_agent.OpenAIAgent(client, system=tools.AGENT_SYSTEM, model=EXCEL_MODEL)
        tools.register_tools(brain2, lambda: wb, on_change=lambda: sheets.save(wb, sample))
        brain2.run("Добавь строку: Date, 7, active")
        reloaded = sheets.load(sample)
        rows = sheets.read_range(reloaded, "A1:C5")
        ok2 = any(r and r[0] == "Date" for r in rows)
        print(f"[{'PASS' if ok2 else 'FAIL'}] row appended ({len(rows)} rows now)", flush=True)
        failures += 0 if ok2 else 1

        print("[..] step 3 — send edited file to owner chat", flush=True)
        await telegram.send_document(
            sample.read_bytes(), caption="✅ excel-bot live E2E result", filename="sample.xlsx"
        )
        print("[PASS] file sent to owner", flush=True)
    finally:
        await telegram.close()
    print(f"\n{failures} failure(s).", flush=True)
    return failures


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
    cfg = _gate()
    if cfg is None:
        return 0  # skipped or config-printed; not a hard failure for CI
    return asyncio.run(_run(cfg))


if __name__ == "__main__":
    raise SystemExit(main())
