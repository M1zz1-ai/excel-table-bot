"""Shared core for the excel-table-bot.

Generic, bot-agnostic building blocks: Telegram I/O (:mod:`core.tg`), an OpenAI
tool-calling agent (:mod:`core.openai_agent`), redis-backed state
(:mod:`core.state`), spreadsheet load/edit (:mod:`core.sheets`), config loading
(:mod:`core.config`), and resilient error handling (:mod:`core.errors`).

The bot is a thin module on this core, and each bot capability is exposed as an
agent-compatible callable so the tools register uniformly onto the agent.
"""

__version__ = "0.1.0"
