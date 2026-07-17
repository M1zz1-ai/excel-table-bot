"""Entrypoint: one asyncio process running aiogram long-polling for the excel bot.

Run modes:
  python -m excel           # run the bot (long-polling)
  python -m excel --check   # validate config loading, then exit

Config keys (declared required; REDIS_URL has a default):
  TELEGRAM_BOT_TOKEN_TABLE, OPENAI_API_KEY, TELEGRAM_CHAT_ID, REDIS_URL

A missing key fails loud naming the key (core.config.ConfigError) — the process
never silently runs without credentials.
"""

from __future__ import annotations

import argparse
import asyncio
import logging
import sys

import openai
from aiogram import Dispatcher
from aiogram.filters import Command
from aiogram.types import Message

from core import config, state, tg
from core.errors import ConfigError

from .bot import ExcelBot
from .session import ExcelSession

logger = logging.getLogger("excel_bot")

REQUIRED_KEYS = [
    "TELEGRAM_BOT_TOKEN_TABLE",
    "OPENAI_API_KEY",
    "TELEGRAM_CHAT_ID",
    "REDIS_URL",
]


def build_bot(cfg: config.Config) -> tuple[ExcelBot, tg.TelegramClient]:
    """Wire the shared core into an ExcelBot from a loaded config."""
    chat_id = tg.gather_chat_ids(cfg.require("TELEGRAM_CHAT_ID"))[0]
    telegram = tg.TelegramClient.from_token(cfg.require("TELEGRAM_BOT_TOKEN_TABLE"), chat_id)
    client = openai.OpenAI(api_key=cfg.require("OPENAI_API_KEY"))
    store = state.RedisState(cfg.require("REDIS_URL"), namespace="excel_bot")
    session = ExcelSession(store)
    bot = ExcelBot(telegram, client, session, owner_chat_id=chat_id)
    return bot, telegram


def build_dispatcher(bot: ExcelBot) -> Dispatcher:
    """Wire aiogram message handlers onto the ExcelBot."""
    dp = Dispatcher()

    @dp.message(Command("start", "help"))
    async def _on_start(message: Message) -> None:
        await bot.on_start(message.chat.id)

    @dp.message(Command("send"))
    async def _on_send(message: Message) -> None:
        await bot.on_send(message.chat.id)

    @dp.message(Command("cancel", "reset"))
    async def _on_cancel(message: Message) -> None:
        await bot.on_cancel(message.chat.id)

    @dp.message(lambda m: m.document is not None)
    async def _on_document(message: Message) -> None:
        doc = message.document
        file = await message.bot.get_file(doc.file_id)
        buf = await message.bot.download_file(file.file_path)
        data = buf.read() if buf is not None else b""
        await bot.on_document(message.chat.id, data, doc.file_name or "table.xlsx")

    @dp.message(lambda m: bool(m.text) and not m.text.startswith("/"))
    async def _on_text(message: Message) -> None:
        await bot.on_text(message.chat.id, message.text or "")

    return dp


async def run(cfg: config.Config) -> None:
    """Build everything and run aiogram long-polling until interrupted."""
    bot, telegram = build_bot(cfg)
    dp = build_dispatcher(bot)
    logger.info("excel-bot started; long-polling")
    try:
        await dp.start_polling(telegram.bot, handle_signals=False)
    finally:
        await telegram.close()


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    parser = argparse.ArgumentParser(prog="excel")
    parser.add_argument("--check", action="store_true", help="Validate config and exit.")
    args = parser.parse_args()

    try:
        cfg = config.load(REQUIRED_KEYS)
    except ConfigError as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        return 1

    if args.check:
        print(f"Config OK — all {len(REQUIRED_KEYS)} required keys present.")
        return 0

    try:
        asyncio.run(run(cfg))
    except KeyboardInterrupt:
        logger.info("shutting down")
    return 0


if __name__ == "__main__":
    sys.exit(main())
