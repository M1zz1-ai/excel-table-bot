"""Excel-bot capabilities: sheet-op tools over ``core.sheets``, driven by Claude.

The n8n "Excel Tables · TG Bot" (9O36SPWrJ3hpOzAQ) hard-codes three modes
(Find / Reformat-to-template / Compare) across 13 code nodes and a Redis state
machine. The win here is to collapse that into ONE Claude tool-calling agent:
each spreadsheet operation is a small callable registered on ``core.openai_agent.OpenAIAgent``,
and Claude decides which to call. The agent IS the router — no state machine.

Each capability is a plain callable (a unification constraint), so a
future unified agent can register every bot's tools onto one Agent. The tools are
bound to a single workbook via :func:`build_sheet_tools` closures, because a tool
needs to mutate the chat's currently loaded spreadsheet.

The system prompt is lifted from the n8n "AI Find" node and generalized from
read-only analysis to full CRUD.
"""

from __future__ import annotations

import json
import logging
import os
import re
from collections.abc import Callable
from typing import Any

from openpyxl import Workbook

from core import sheets

logger = logging.getLogger(__name__)

# LLM brain: OpenAI. A fast/cheap reasoning model handles the tool-calling CRUD
# loop; bump to gpt-5.5 via ``EXCEL_MODEL`` if CRUD accuracy regresses.
EXCEL_MODEL = os.getenv("EXCEL_MODEL", "gpt-5.4-mini")

# Lifted from the n8n "AI Find" node and generalized to conversational CRUD.
AGENT_SYSTEM = (
    "You are a precise spreadsheet assistant operating over Telegram. The user "
    "has uploaded an .xlsx file and gives you free-form instructions to inspect "
    "and edit it. You drive the spreadsheet ONLY through the provided tools — "
    "never invent data.\n\n"
    "DATA FORMAT — IMPORTANT:\n"
    "Cells are addressed in A1 notation (A1, B2, ...). HEADERS ARE NOT "
    "AUTO-DETECTED. The first few rows often contain metadata (report title, "
    "date range, store name, blank rows, group headings). The ACTUAL "
    "column-header row appears somewhere in the top 5-6 rows — recognise it by "
    'content (e.g. "Payment Type", "Net sales", "price", "quantity", "name", '
    '"category"). After you identify the header row:\n'
    "  - map cell letters to real column names (A -> the first header, etc.);\n"
    "  - all rows BELOW the header are data; rows ABOVE are metadata (use for "
    "context, not as data);\n"
    '  - a row labelled "Total"/"Totals" near the bottom is a precomputed '
    "summary — exclude it from per-item math unless the user asks for the total.\n\n"
    "Available tools:\n"
    "  - list_sheets(): names of the worksheets.\n"
    "  - read_range(cell_range, sheet): read a range like 'A1:D20' or a single "
    "cell 'B2'. Use this to FIND, FILTER and inspect labels — NOT for arithmetic.\n"
    "  - column_stats(cell_range, sheet): deterministically sum/count/min/max/mean "
    "the numbers in a range. ALWAYS use this for any total/sum/average — it parses "
    "text-formatted numbers ('12,50', '1 234,56', '828,62 EUR') that hand-math "
    "would get wrong. Pass the DATA rows only (exclude any total row).\n"
    "  - write_cell(cell, value, sheet): EDIT a single cell.\n"
    "  - append_row(values_json, sheet): ADD a row at the bottom; values_json is "
    'a JSON array, e.g. \'["Alice", 99, "active"]\'.\n'
    "  - write_row(row_index, values_json, sheet): overwrite an existing row "
    "(RESHAPE / RECOMPUTE a row in place).\n\n"
    "Rules:\n"
    "1. Reply in Russian, concise and to the point. Telegram legacy Markdown: "
    "*bold*, _italic_, `code`.\n"
    "2. Before reading/aggregating, ALWAYS read the top rows first to locate the "
    "header row; do not assume row 1 is the header.\n"
    "3. For counting / filtering / aggregating — NEVER add numbers in your head. "
    "Call column_stats over the data range and report its result; do the math in "
    "code, not mentally.\n"
    "3a. TOTAL/SUMMARY ROWS: if the sheet has an explicit total row (Итого / Итог / "
    "Всего / Total / Sum / Grand total), report THAT stated value as the sheet's own "
    "total, AND verify it with column_stats over the data rows only (excluding the "
    "total row). If the stated total differs from the recomputed sum, present BOTH "
    "and flag the discrepancy explicitly — e.g. «в файле указано 828.62, пересчёт по "
    "колонке даёт 530.72 — расходится; вероятная причина: …». Never silently pick one.\n"
    "4. For an edit (write_cell / append_row / write_row) confirm what you "
    "changed (which cell/row, old -> new) after the tool succeeds.\n"
    "5. If a request can't be satisfied from this table — say so honestly. DO "
    "NOT invent values.\n"
    "6. After editing, remind the user they can send /send to get the updated "
    "file back.\n"
    "7. If the user gave no clear instruction — describe the table in 3-5 lines "
    "and suggest 2-3 example actions."
)


def _coerce_values(values_json: Any) -> list[Any]:
    """Normalize tool-supplied row values to a flat ``list``.

    The model is supposed to pass a JSON-array string (its tool schema declares a
    scalar param), e.g. ``'["Alice", 99]'`` — but in practice it often sends a
    real JSON array or object instead. Accept any of:

      - a ``list`` (already a row);
      - a ``dict`` (use its values, in insertion order);
      - a JSON-array string ``'["a", 1]'``;
      - a JSON-object string ``'{"a": 1}'`` (use its values).

    Genuinely bad input — a non-JSON string, or a scalar (number/bool/null) — is
    still rejected with ``SheetError``.
    """
    if isinstance(values_json, list):
        return values_json
    if isinstance(values_json, dict):
        return list(values_json.values())
    if not isinstance(values_json, str):
        raise sheets.SheetError(
            f"values must be a list, dict, or JSON string, got {type(values_json).__name__}"
        )
    try:
        parsed = json.loads(values_json)
    except json.JSONDecodeError as exc:
        raise sheets.SheetError(f"values_json is not valid JSON: {exc}") from exc
    if isinstance(parsed, dict):
        return list(parsed.values())
    if not isinstance(parsed, list):
        raise sheets.SheetError("values_json must be a JSON array or object")
    return parsed


# Unicode spaces used as thousands separators in exported spreadsheets.
_SPACE_CHARS = "\xa0   "


def _coerce_number(value: Any) -> float | None:
    """Best-effort parse a cell value into a float, or ``None`` if not numeric.

    Handles the text-formatted numbers common in Russian ``.xls`` exports that
    otherwise trip up naive parsing (the root cause of inconsistent column sums):

      - comma decimal separator: ``"12,50"`` -> 12.5
      - space / no-break-space thousands: ``"1 234,56"`` / ``"1\\xa0234,56"``
      - currency suffix/prefix and stray symbols: ``"828,62 EUR"``, ``"€1 000"``
      - mixed EU/US grouping: ``"1.234,56"`` and ``"1,234.56"`` -> 1234.56

    ``bool`` is treated as non-numeric (a checkbox cell is not a quantity).
    """
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        return None
    s = value.strip()
    for ch in _SPACE_CHARS:
        s = s.replace(ch, " ")
    # Drop everything that isn't a digit, separator, or sign (currency, letters, %).
    s = re.sub(r"[^\d,.\-+ ]", "", s).replace(" ", "")
    if not s or s in "+-":
        return None
    has_comma, has_dot = "," in s, "." in s
    if has_comma and has_dot:
        # The right-most separator is the decimal point; the other groups thousands.
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma:
        s = s.replace(",", "") if s.count(",") > 1 else s.replace(",", ".")
    elif has_dot and s.count(".") > 1:
        s = s.replace(".", "")  # e.g. "1.234.567" (EU grouping)
    try:
        return float(s)
    except ValueError:
        return None


def build_sheet_tools(
    wb_provider: Callable[[], Workbook],
    on_change: Callable[[], None] | None = None,
) -> list[Callable[..., Any]]:
    """Build sheet-op tool callables bound to a chat's live workbook.

    The tools close over ``wb_provider`` (returns the chat's currently loaded
    Workbook) so the same callables operate on whichever file is loaded. Edits
    call ``on_change`` (if given) so the bot can persist the workbook to its
    per-chat temp file after each mutation.

    Returns plain callables; register them on a ``core.openai_agent.OpenAIAgent`` with
    :func:`register_tools` (schema is inferred from each signature, the docstring
    becomes the description — matching a unification constraint).

    The optional ``sheet`` param is an empty string by default; an empty string
    means "the active sheet" (``core.sheets`` treats ``None`` that way, and the
    model can only send strings through its tool schema).
    """

    def _sheet_arg(sheet: str) -> str | None:
        return sheet or None

    def _persist() -> None:
        if on_change is not None:
            on_change()

    def list_sheets() -> list[str]:
        """Return the worksheet names in order."""
        return sheets.sheet_names(wb_provider())

    def read_range(cell_range: str, sheet: str = "") -> list[list[Any]]:
        """Read a range in A1 notation (e.g. 'A1:D20' or a single cell 'B2').

        Returns rows of cell values. Use to find / filter / inspect labels.
        For arithmetic over numbers use column_stats — do NOT sum in your head.
        """
        return sheets.read_range(wb_provider(), cell_range, sheet=_sheet_arg(sheet))

    def column_stats(cell_range: str, sheet: str = "") -> dict[str, Any]:
        """Deterministically aggregate the NUMBERS in a range (sum/count/min/max/mean).

        Use this for every sum / average / count / total instead of adding values
        yourself. Text-formatted numbers ('12,50', '1 234,56', '828,62 EUR') are
        parsed robustly. ``skipped`` counts non-empty cells that were NOT numeric
        (e.g. a label like 'Итого:'), so pass the DATA rows only (exclude the
        total row) when recomputing a total. Blank cells are ignored.
        """
        rows = sheets.read_range(wb_provider(), cell_range, sheet=_sheet_arg(sheet))
        nums: list[float] = []
        skipped = 0
        for row in rows:
            for cell in row:
                n = _coerce_number(cell)
                if n is None:
                    if cell is not None and str(cell).strip() != "":
                        skipped += 1
                else:
                    nums.append(n)
        if not nums:
            return {"count": 0, "sum": 0.0, "skipped": skipped,
                    "min": None, "max": None, "mean": None}
        total = round(sum(nums), 10)
        return {
            "count": len(nums),
            "sum": total,
            "skipped": skipped,
            "min": min(nums),
            "max": max(nums),
            "mean": round(total / len(nums), 10),
        }

    def write_cell(cell: str, value: str, sheet: str = "") -> str:
        """Write a single cell in A1 notation (e.g. write_cell('B2', '42'))."""
        sheets.write_cell(wb_provider(), cell, value, sheet=_sheet_arg(sheet))
        _persist()
        return f"wrote {value!r} to {cell}"

    def append_row(values_json: Any, sheet: str = "") -> str:
        """Append a row at the bottom. values_json is a JSON array of cell values."""
        values = _coerce_values(values_json)
        sheets.append_row(wb_provider(), values, sheet=_sheet_arg(sheet))
        _persist()
        return f"appended row with {len(values)} value(s)"

    def write_row(row_index: int, values_json: str, sheet: str = "") -> str:
        """Overwrite an existing row (1-based). values_json is a JSON array."""
        values = _coerce_values(values_json)
        sheets.write_row(wb_provider(), row_index, values, sheet=_sheet_arg(sheet))
        _persist()
        return f"overwrote row {row_index} with {len(values)} value(s)"

    return [list_sheets, read_range, column_stats, write_cell, append_row, write_row]


def register_tools(
    agent: Any,
    wb_provider: Callable[[], Workbook],
    on_change: Callable[[], None] | None = None,
) -> None:
    """Register the bot's sheet-op tools on a ``core.openai_agent.OpenAIAgent``.

    Keeps the unification path open: a later unified agent registers every bot's
    tools onto one Agent and routes across them.
    """
    for fn in build_sheet_tools(wb_provider, on_change):
        agent.tool(fn)
