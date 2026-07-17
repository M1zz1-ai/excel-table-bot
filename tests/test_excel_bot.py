"""Unit tests for excel.bot — menu-button routing, i18n language switch,
and the duplicate-reply collapse. Telegram, the LLM agent, and the session store
are all faked; no network, no real bot.

Regression coverage for two live-smoke bugs:
  1. Help / RU / ENG buttons fell through to the LLM CRUD agent.
  2. A reply arrived as one message with the same paragraph twice.
"""

from __future__ import annotations

from typing import Any

import pytest

from excel import i18n
from excel.bot import ExcelBot

# ---- fakes --------------------------------------------------------------


class _FakeTg:
    def __init__(self) -> None:
        self.texts: list[dict[str, Any]] = []
        self.docs: list[dict[str, Any]] = []

    async def send_text(self, text: str, chat_id: int | None = None, **kw: Any) -> Any:
        self.texts.append({"text": text, "chat_id": chat_id, "reply_markup": kw.get("reply_markup")})

    async def send_document(self, document: Any, chat_id: int | None = None, **kw: Any) -> Any:
        self.docs.append({"chat_id": chat_id, "caption": kw.get("caption")})


class _FakeActive:
    max_row = 3
    max_column = 2


class _FakeWb:
    sheetnames = ["Sheet"]
    active = _FakeActive()


class _FakeSession:
    def __init__(self, *, has: bool = False) -> None:
        self._has = has
        self._lang: str | None = None
        self._wb = _FakeWb()

    def has_file(self, chat_id: int) -> bool:
        return self._has

    def wizard_get(self, chat_id: int) -> Any:
        return None  # no wizard active in these tests

    def wizard_set(self, chat_id: int, state: Any) -> None:
        pass

    def wizard_clear(self, chat_id: int) -> None:
        pass

    def get_lang(self, chat_id: int) -> str | None:
        return self._lang

    def set_lang(self, chat_id: int, lang: str) -> None:
        self._lang = lang

    def workbook(self, chat_id: int) -> _FakeWb:
        return self._wb

    def flush(self, chat_id: int) -> None:
        pass

    def get_filename(self, chat_id: int) -> str:
        return "t.xlsx"

    def read_bytes(self, chat_id: int) -> bytes:
        return b"XLSX"

    def load_from_bytes(self, chat_id: int, data: bytes, filename: str) -> None:
        self._has = True


class _FakeBrain:
    def __init__(self, reply: str) -> None:
        self._reply = reply

    def run(self, prompt: str) -> str:
        return self._reply


def _bot(*, has: bool = False):
    tg = _FakeTg()
    sess = _FakeSession(has=has)
    bot = ExcelBot(tg, object(), sess, owner_chat_id=1)  # type: ignore[arg-type]
    return bot, tg, sess


# ---- bug 1: buttons must never reach the LLM agent ----------------------


@pytest.mark.asyncio
async def test_help_button_does_not_reach_agent():
    bot, tg, sess = _bot(has=True)
    called: list[str] = []
    bot._run_agent = lambda chat_id, text: called.append(text) or "X"  # type: ignore[method-assign]

    await bot.on_text(7, i18n.BTN_HELP)

    assert called == []  # the LLM was NOT invoked
    assert tg.texts[-1]["text"] == i18n.t("ru", "help")
    assert tg.texts[-1]["reply_markup"] is not None  # keyboard re-attached


@pytest.mark.asyncio
async def test_eng_button_switches_and_persists_language():
    bot, tg, sess = _bot(has=True)
    called: list[str] = []
    bot._run_agent = lambda chat_id, text: called.append(text) or "X"  # type: ignore[method-assign]

    await bot.on_text(7, i18n.BTN_ENG)

    assert called == []
    assert sess.get_lang(7) == "en"
    assert tg.texts[-1]["text"] == i18n.t("en", "lang_set")


@pytest.mark.asyncio
async def test_ru_button_switches_language():
    bot, tg, sess = _bot(has=True)
    sess.set_lang(7, "en")
    await bot.on_text(7, i18n.BTN_RU)
    assert sess.get_lang(7) == "ru"
    assert tg.texts[-1]["text"] == i18n.t("ru", "lang_set")


@pytest.mark.asyncio
async def test_find_button_sends_hint_not_agent():
    bot, tg, sess = _bot(has=True)
    called: list[str] = []
    bot._run_agent = lambda chat_id, text: called.append(text) or "X"  # type: ignore[method-assign]

    await bot.on_text(7, i18n.BTN_FIND)

    assert called == []  # never reaches the LLM
    assert tg.texts[-1]["text"] == i18n.t("ru", "find_hint")
    assert tg.texts[-1]["reply_markup"] is not None


@pytest.mark.asyncio
@pytest.mark.parametrize("label,step1", [(i18n.BTN_REFORMAT, "1/3"), (i18n.BTN_COMPARE, "1/3")])
async def test_reformat_compare_buttons_start_wizard_not_agent(label, step1):
    bot, tg, sess = _bot(has=True)
    called: list[str] = []
    bot._run_agent = lambda chat_id, text: called.append(text) or "X"  # type: ignore[method-assign]

    await bot.on_text(7, label)

    assert called == []  # starts a wizard, never the LLM CRUD agent
    assert step1 in tg.texts[-1]["text"]


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "label,action",
    [
        ("🌐 𝐄𝐍𝐆", i18n.ACTION_LANG_EN),
        ("🔍 𝐅𝐢𝐧𝐝", i18n.ACTION_FIND),
        ("⚖️ 𝐂𝐨𝐦𝐩𝐚𝐫𝐞", i18n.ACTION_COMPARE),
        ("✏️ 𝐑𝐞𝐟𝐨𝐫𝐦𝐚𝐭", i18n.ACTION_REFORMAT),
        ("ℹ️ 𝐇𝐞𝐥𝐩", i18n.ACTION_HELP),
    ],
)
async def test_legacy_fancy_button_labels_still_intercepted(label, action):
    """Users whose Telegram still shows the OLD n8n keyboard press fancy labels."""
    bot, tg, sess = _bot(has=True)
    called: list[str] = []
    bot._run_agent = lambda chat_id, text: called.append(text) or "X"  # type: ignore[method-assign]

    await bot.on_text(7, label)

    assert called == []  # no fancy legacy label reaches the LLM
    assert i18n.classify_button(label) == action


def test_menu_layout_matches_original_n8n_rows():
    """Snapshot: 3 rows, mirroring Excel_Tables_·_TG_Bot.json."""
    assert i18n.MENU_ROWS == [
        [i18n.BTN_FIND, i18n.BTN_COMPARE, i18n.BTN_REFORMAT],
        [i18n.BTN_HELP],
        [i18n.BTN_RU, i18n.BTN_ENG],
    ]


@pytest.mark.asyncio
async def test_plain_text_reaches_agent():
    bot, tg, sess = _bot(has=True)
    called: list[str] = []
    bot._run_agent = lambda chat_id, text: called.append(text) or "42, объяснение"  # type: ignore[method-assign]

    await bot.on_text(7, "сколько строк со статусом active")

    assert called == ["сколько строк со статусом active"]
    assert tg.texts[-1]["text"] == "42, объяснение"


# ---- i18n on the other surfaces ----------------------------------------


@pytest.mark.asyncio
async def test_no_file_message_localized_to_english():
    bot, tg, sess = _bot(has=False)
    sess.set_lang(7, "en")
    await bot.on_text(7, "count the rows")
    assert tg.texts[-1]["text"] == i18n.t("en", "no_file")


@pytest.mark.asyncio
async def test_start_attaches_menu_keyboard():
    bot, tg, sess = _bot()
    await bot.on_start(7)
    assert tg.texts[-1]["text"] == i18n.t("ru", "welcome")
    assert tg.texts[-1]["reply_markup"] is not None


@pytest.mark.asyncio
async def test_upload_ack_localized_english():
    bot, tg, sess = _bot()
    sess.set_lang(7, "en")
    await bot.on_document(7, b"data", "sales.xlsx")
    ack = tg.texts[-1]["text"]
    assert "Loaded" in ack and "sales.xlsx" in ack and "3 rows × 2 columns" in ack


# ---- bug 2: duplicate-reply collapse -----------------------------------


def test_collapse_repeat_exact_duplicate_with_separator():
    dup = "Файл открыт, 3 строки, 2 заголовка.\n\nФайл открыт, 3 строки, 2 заголовка."
    assert i18n.collapse_repeat(dup) == "Файл открыт, 3 строки, 2 заголовка."


def test_collapse_repeat_exact_duplicate_no_separator():
    assert i18n.collapse_repeat("abcabc") == "abc"


def test_collapse_repeat_leaves_non_duplicate_untouched():
    text = "42 строки.\n\nПояснение: посчитал по колонке статус."
    assert i18n.collapse_repeat(text) == text


def test_run_agent_collapses_doubled_reply():
    bot, tg, sess = _bot(has=True)
    doubled = "Готово, добавил строку.\n\nГотово, добавил строку."
    bot._build_agent = lambda chat_id: _FakeBrain(doubled)  # type: ignore[method-assign]
    assert bot._run_agent(7, "добавь строку") == "Готово, добавил строку."
