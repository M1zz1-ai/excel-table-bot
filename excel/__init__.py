"""Excel bot: upload a spreadsheet -> conversational CRUD via an LLM -> get it back.

Three modes over Telegram: a conversational CRUD agent, a template reformat wizard,
and a smart compare. The agent's tools are ``core.sheets`` operations bound to the
chat's working file, so one tool-calling agent replaces a hand-written state machine.
"""

from .session import ExcelSession
from .tools import (
    AGENT_SYSTEM,
    build_sheet_tools,
    register_tools,
)

__all__ = [
    "AGENT_SYSTEM",
    "ExcelSession",
    "build_sheet_tools",
    "register_tools",
]
