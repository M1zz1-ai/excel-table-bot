"""Unit tests for excel.session — per-chat workbook lifecycle.

Uses a real temp .xlsx on disk and an in-memory state stand-in (no redis).
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook

from core import sheets
from core.errors import SheetError
from excel.session import ExcelSession


class _FakeState:
    """In-memory stand-in for core.state.RedisState session store."""

    def __init__(self) -> None:
        self.store: dict[str, Any] = {}

    def get_session(self, key: str, default: Any = None) -> Any:
        return self.store.get(key, default)

    def set_session(self, key: str, value: Any, **kw: Any) -> None:
        self.store[key] = value

    def clear_session(self, key: str) -> None:
        self.store.pop(key, None)


def _xlsx_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(["name", "qty"])
    ws.append(["Apple", 10])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _session(tmp_path: Path) -> ExcelSession:
    return ExcelSession(_FakeState(), workdir=tmp_path)


def test_no_file_initially(tmp_path: Path) -> None:
    sess = _session(tmp_path)
    assert not sess.has_file(99)
    assert sess.get_path(99) is None


def test_load_from_bytes_makes_working_file(tmp_path: Path) -> None:
    sess = _session(tmp_path)
    sess.load_from_bytes(99, _xlsx_bytes(), "sales.xlsx")
    assert sess.has_file(99)
    assert sess.get_filename(99) == "sales.xlsx"
    wb = sess.workbook(99)
    assert sheets.read_range(wb, "A1:B2") == [["name", "qty"], ["Apple", 10]]


def test_load_rejects_garbage_bytes(tmp_path: Path) -> None:
    sess = _session(tmp_path)
    with pytest.raises(SheetError):
        sess.load_from_bytes(99, b"this is not an xlsx", "bad.xlsx")


def test_load_csv_normalizes_filename_and_disk_to_xlsx(tmp_path: Path) -> None:
    """A non-xlsx upload is normalized: reply filename + on-disk file become .xlsx."""
    sess = _session(tmp_path)
    sess.load_from_bytes(99, b"name,qty\nApple,10\n", "sales.csv")
    # reply filename swaps the extension; the working file on disk is .xlsx
    assert sess.get_filename(99) == "sales.xlsx"
    assert sess.get_path(99) == tmp_path / "99.xlsx"
    wb = sess.workbook(99)
    assert sheets.read_range(wb, "A1:B2") == [["name", "qty"], ["Apple", "10"]]


def test_edit_flush_read_bytes_round_trip(tmp_path: Path) -> None:
    sess = _session(tmp_path)
    sess.load_from_bytes(99, _xlsx_bytes(), "sales.xlsx")
    wb = sess.workbook(99)
    sheets.write_cell(wb, "B2", 42)
    # read_bytes flushes pending edits to disk and returns the new file.
    data = sess.read_bytes(99)
    # Persist the returned bytes and confirm the edit survived a real save/load.
    verify = tmp_path / "verify.xlsx"
    verify.write_bytes(data)
    reloaded = sheets.load(verify)
    assert sheets.read_range(reloaded, "B2") == [[42]]


def test_workbook_without_file_raises(tmp_path: Path) -> None:
    sess = _session(tmp_path)
    with pytest.raises(SheetError):
        sess.workbook(99)


def test_clear_drops_file(tmp_path: Path) -> None:
    sess = _session(tmp_path)
    sess.load_from_bytes(99, _xlsx_bytes(), "sales.xlsx")
    sess.clear(99)
    assert not sess.has_file(99)
    assert not (tmp_path / "99.xlsx").exists()
