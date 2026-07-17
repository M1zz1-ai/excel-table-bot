"""Per-chat workbook lifecycle: the current working .xlsx for each chat.

Replaces the n8n Redis state machine. We keep it minimal: each chat has at most
one working file. An uploaded file of any supported format (see
``core.sheets.read_bytes_to_workbook``) is normalized to an openpyxl workbook and
persisted as a per-chat temp .xlsx on disk; Redis (``core.state``) holds only the
path + the reply filename (extension swapped to .xlsx) so the bot survives a
restart's worth of session memory (Redis-backed, with the core's graceful no-op
fallback when Redis is down).

The loaded ``openpyxl.Workbook`` itself is cached in-process per chat so the
agent's tool calls within one turn mutate one workbook; mutations are flushed
back to the temp file by the bot after each tool change.
"""

from __future__ import annotations

import logging
import tempfile
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from core import sheets

logger = logging.getLogger(__name__)

# Per-chat temp files live under one dir so they're easy to find / clean.
WORKDIR = Path(tempfile.gettempdir()) / "excel_table_bot"


class ExcelSession:
    """Tracks the current working spreadsheet per chat.

    Args:
        state: a ``core.state.RedisState`` (or compatible) for path persistence.
        workdir: directory for per-chat temp .xlsx files (overridable in tests).
    """

    def __init__(self, state: Any, *, workdir: Path = WORKDIR) -> None:
        self._state = state
        self._workdir = workdir
        self._workbooks: dict[int, Workbook] = {}

    # ---- redis-backed file pointer -------------------------------------

    def _meta_key(self, chat_id: int) -> str:
        return f"file:{chat_id}"

    def _lang_key(self, chat_id: int) -> str:
        return f"lang:{chat_id}"

    # ---- interface language (per chat, independent of the working file) ---

    def get_lang(self, chat_id: int) -> str | None:
        """Return the chat's stored UI language code ('ru'/'en'), or None if unset."""
        return self._state.get_session(self._lang_key(chat_id))

    def set_lang(self, chat_id: int, lang: str) -> None:
        """Persist the chat's UI language choice."""
        self._state.set_session(self._lang_key(chat_id), lang)

    # ---- wizard state (reformat / compare multi-step flows) --------------

    WIZARD_TTL = 15 * 60  # 15 min mid-flow (mirrors the image bot's guided-chain TTL)

    def _wizard_key(self, chat_id: int) -> str:
        return f"wizard:{chat_id}"

    def wizard_get(self, chat_id: int) -> dict[str, Any] | None:
        """Return the chat's active wizard state, or None if no flow is running."""
        return self._state.get_session(self._wizard_key(chat_id))

    def wizard_set(self, chat_id: int, state: dict[str, Any]) -> None:
        """Persist wizard state with the mid-flow TTL."""
        self._state.set_session(self._wizard_key(chat_id), state, ttl=self.WIZARD_TTL)

    def wizard_clear(self, chat_id: int) -> None:
        """Drop the wizard state and any temp files it stored on disk."""
        state = self.wizard_get(chat_id) or {}
        for path in (state.get("files") or {}).values():
            Path(path).unlink(missing_ok=True)
        self._state.clear_session(self._wizard_key(chat_id))

    def wizard_save_file(self, chat_id: int, role: str, data: bytes, filename: str) -> str:
        """Store an uploaded wizard file's bytes on disk (not in redis). Returns the path."""
        self._workdir.mkdir(parents=True, exist_ok=True)
        suffix = Path(filename).suffix or ".xlsx"
        path = self._workdir / f"{chat_id}_{role}{suffix}"
        path.write_bytes(data)
        return str(path)

    def _path_for(self, chat_id: int) -> Path:
        return self._workdir / f"{chat_id}.xlsx"

    def has_file(self, chat_id: int) -> bool:
        """True if this chat has a working file loaded."""
        return self.get_path(chat_id) is not None

    def get_path(self, chat_id: int) -> Path | None:
        """Return the path of the chat's working file, or None if none/missing."""
        meta = self._state.get_session(self._meta_key(chat_id))
        if not meta:
            return None
        path = Path(meta["path"])
        return path if path.exists() else None

    def get_filename(self, chat_id: int) -> str:
        """Return the original uploaded filename (or a default)."""
        meta = self._state.get_session(self._meta_key(chat_id)) or {}
        return meta.get("filename", "table.xlsx")

    # ---- load / mutate / read back -------------------------------------

    def load_from_bytes(self, chat_id: int, data: bytes, filename: str) -> None:
        """Normalize uploaded bytes (any supported format) to the chat's .xlsx file.

        The input may be .xlsx/.xlsm/.xls/.csv/.tsv; it is parsed into an openpyxl
        workbook and saved as the per-chat .xlsx so all downstream code stays
        xlsx-shaped. The stored reply filename swaps the extension to .xlsx.

        Raises:
            SheetError: if the bytes are an unsupported or corrupt spreadsheet.
        """
        self._workdir.mkdir(parents=True, exist_ok=True)
        # Parse + normalize BEFORE touching disk, so a bad upload leaves no file.
        wb = sheets.read_bytes_to_workbook(data, filename=filename)
        path = self._path_for(chat_id)
        sheets.save(wb, path)  # always written as .xlsx
        self._workbooks[chat_id] = wb
        reply_filename = Path(filename or "table.xlsx").with_suffix(".xlsx").name
        self._state.set_session(
            self._meta_key(chat_id), {"path": str(path), "filename": reply_filename}
        )

    def workbook(self, chat_id: int) -> Workbook:
        """Return the chat's loaded Workbook (re-loading from disk if not cached).

        Raises:
            SheetError: if no working file exists for this chat.
        """
        cached = self._workbooks.get(chat_id)
        if cached is not None:
            return cached
        path = self.get_path(chat_id)
        if path is None:
            raise sheets.SheetError("no working file for this chat")
        wb = sheets.load(path)
        self._workbooks[chat_id] = wb
        return wb

    def flush(self, chat_id: int) -> None:
        """Save the chat's in-memory workbook back to its temp file."""
        wb = self._workbooks.get(chat_id)
        path = self.get_path(chat_id)
        if wb is not None and path is not None:
            sheets.save(wb, path)

    def read_bytes(self, chat_id: int) -> bytes:
        """Return the current working file's bytes (after flushing pending edits).

        Raises:
            SheetError: if no working file exists for this chat.
        """
        self.flush(chat_id)
        path = self.get_path(chat_id)
        if path is None:
            raise sheets.SheetError("no working file for this chat")
        return path.read_bytes()

    def clear(self, chat_id: int) -> None:
        """Drop the chat's working file (in-memory cache + redis pointer + disk)."""
        self._workbooks.pop(chat_id, None)
        path = self._path_for(chat_id)
        if path.exists():
            path.unlink()
        self._state.clear_session(self._meta_key(chat_id))
