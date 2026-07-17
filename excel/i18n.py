"""Interface localization for the excel bot (RU/ENG) + the persistent menu.

The original n8n "Excel Tables · TG Bot" had a persistent reply keyboard with a
Help page and an RU/ENG language switcher (state key ``lang``). The Python rewrite
dropped it, so those buttons fell through to the LLM CRUD agent. This module
restores the feature: the button labels, the per-language UI strings, and a
``classify_button`` that keeps a button press from ever reaching the agent.

Language is persisted per chat in redis (see ``ExcelSession.get_lang`` /
``set_lang``). Default is Russian (the bot's primary audience); ENG is opt-in.
"""

from __future__ import annotations

import re

LANG_RU = "ru"
LANG_EN = "en"
DEFAULT_LANG = LANG_RU

SUPPORTED_FORMATS = ".xlsx, .xlsm, .xls, .csv, .tsv"

# Current button labels (sent on the fresh keyboard). The layout mirrors the
# original n8n "Excel Tables · TG Bot" keyboard (3 rows): Find/Compare/Reformat,
# then Help, then the RU/ENG language toggle.
BTN_FIND = "🔍 Find"
BTN_COMPARE = "⚖️ Compare"
BTN_REFORMAT = "✏️ Reformat"
BTN_HELP = "ℹ️ Help"
BTN_RU = "🌐 RU"
BTN_ENG = "🌐 ENG"
MENU_ROWS = [
    [BTN_FIND, BTN_COMPARE, BTN_REFORMAT],
    [BTN_HELP],
    [BTN_RU, BTN_ENG],
]

# Button actions.
ACTION_FIND = "find"
ACTION_COMPARE = "compare"
ACTION_REFORMAT = "reformat"
ACTION_HELP = "help"
ACTION_LANG_RU = "lang_ru"
ACTION_LANG_EN = "lang_en"

# Simple-reply buttons -> the i18n string key to send back (localized). The two
# language buttons are handled separately (they mutate persisted state).
SIMPLE_ACTION_STRING = {
    ACTION_FIND: "find_hint",
    ACTION_COMPARE: "compare_unavailable",
    ACTION_REFORMAT: "reformat_hint",
    ACTION_HELP: "help",
}

# Label -> action. Includes the OLD n8n keyboard's fancy-unicode labels so the
# fix works immediately for users whose Telegram still shows the legacy keyboard
# (Telegram keeps a reply keyboard until it's replaced by a new /start), plus a
# bare-word alias for each.
_BUTTON_ACTIONS = {
    BTN_FIND: ACTION_FIND,
    "Find": ACTION_FIND,
    "🔍 𝐅𝐢𝐧𝐝": ACTION_FIND,  # legacy n8n label
    BTN_COMPARE: ACTION_COMPARE,
    "Compare": ACTION_COMPARE,
    "⚖️ 𝐂𝐨𝐦𝐩𝐚𝐫𝐞": ACTION_COMPARE,  # legacy n8n label
    BTN_REFORMAT: ACTION_REFORMAT,
    "Reformat": ACTION_REFORMAT,
    "✏️ 𝐑𝐞𝐟𝐨𝐫𝐦𝐚𝐭": ACTION_REFORMAT,  # legacy n8n label
    BTN_HELP: ACTION_HELP,
    "Help": ACTION_HELP,
    "ℹ️ 𝐇𝐞𝐥𝐩": ACTION_HELP,  # legacy n8n label
    BTN_RU: ACTION_LANG_RU,
    "RU": ACTION_LANG_RU,
    "🌐 𝐑𝐔": ACTION_LANG_RU,  # legacy n8n label
    BTN_ENG: ACTION_LANG_EN,
    "ENG": ACTION_LANG_EN,
    "🌐 𝐄𝐍𝐆": ACTION_LANG_EN,  # legacy n8n label
}


def classify_button(text: str) -> str | None:
    """Return the menu action for a button label, or ``None`` if it's not a button.

    A non-``None`` result means the message is a UI button press and must be
    handled by the bot, never forwarded to the LLM CRUD agent.
    """
    return _BUTTON_ACTIONS.get((text or "").strip())


def normalize_lang(value: str | None) -> str:
    """Coerce a stored/raw language value to a supported code (default RU)."""
    return value if value in (LANG_RU, LANG_EN) else DEFAULT_LANG


_STRINGS: dict[str, dict[str, str]] = {
    LANG_RU: {
        "welcome": (
            "<b>📊 Excel bot</b>\n\n"
            f"Пришли таблицу (<b>{SUPPORTED_FORMATS}</b>) — потом пиши, что с ней "
            "сделать обычными словами:\n"
            "• «сколько строк со статусом active»\n"
            "• «добавь строку: Богдан, 99, active»\n"
            "• «поставь в B2 значение 42»\n"
            "• «пересчитай сумму в последней строке»\n\n"
            "Когда закончишь — отправь /send, и я верну изменённый файл в <b>.xlsx</b>.\n"
            "Кнопки: ℹ️ Help — справка, 🌐 RU / 🌐 ENG — язык интерфейса."
        ),
        "help": (
            "ℹ️ <b>Как пользоваться Excel-ботом</b>\n\n"
            f"1. Пришли таблицу файлом: <b>{SUPPORTED_FORMATS}</b>.\n"
            "2. Пиши, что сделать, обычными словами:\n"
            "   • «сколько строк со статусом active»\n"
            "   • «добавь строку: Богдан, 99, active»\n"
            "   • «поставь в B2 значение 42»\n"
            "   • «посчитай сумму по колонке цена»\n"
            "3. Отправь /send — верну изменённый файл в <b>.xlsx</b>.\n\n"
            "Кнопки 🌐 RU / 🌐 ENG переключают язык интерфейса."
        ),
        "no_file": (
            f"Сначала пришли таблицу документом (<b>{SUPPORTED_FORMATS}</b>), "
            "потом дай инструкцию по ней."
        ),
        "unsupported": (
            f"Не смог прочитать файл. Поддерживаю таблицы: <b>{SUPPORTED_FORMATS}</b>. "
            "Пришли файл в одном из этих форматов."
        ),
        "upload_ack": (
            "✅ Загрузил <b>{filename}</b>\n"
            "Листы: {sheets}\n"
            "Активный лист: {rows} строк × {cols} столбцов\n\n"
            "Пиши, что сделать с таблицей."
        ),
        "done": "Готово.",
        "send_caption": "📊 Обновлённый файл",
        "lang_set": "🌐 Язык интерфейса: русский. Пиши команды или пришли таблицу.",
        "find_hint": (
            "🔍 <b>Поиск / вопрос по таблице</b>\n\n"
            "Просто напиши вопрос обычными словами — например «сколько строк со "
            "статусом active» или «найди строку с именем Богдан». Сначала пришли "
            f"таблицу документом (<b>{SUPPORTED_FORMATS}</b>)."
        ),
        "reformat_hint": (
            "✏️ <b>Переформатирование</b>\n\n"
            "Опиши словами, как пересобрать таблицу — например «перенеси колонку "
            "цена в конец», «заполни SKU по совпадению имени», «поставь в B2 значение "
            "42». Пришли файл, дай инструкцию, потом /send — верну результат в .xlsx."
        ),
        "compare_unavailable": (
            "⚖️ <b>Сравнение файлов</b> пока не поддерживается в этой версии. "
            "Загрузи одну таблицу и работай с ней командами."
        ),
    },
    LANG_EN: {
        "welcome": (
            "<b>📊 Excel bot</b>\n\n"
            f"Send a spreadsheet (<b>{SUPPORTED_FORMATS}</b>), then tell me what to do "
            "in plain words:\n"
            "• \"how many rows have status active\"\n"
            "• \"add a row: Alex, 99, active\"\n"
            "• \"set B2 to 42\"\n"
            "• \"recompute the sum in the last row\"\n\n"
            "When you're done, send /send and I'll return the edited file as <b>.xlsx</b>.\n"
            "Buttons: ℹ️ Help — help, 🌐 RU / 🌐 ENG — interface language."
        ),
        "help": (
            "ℹ️ <b>How to use the Excel bot</b>\n\n"
            f"1. Send a spreadsheet as a file: <b>{SUPPORTED_FORMATS}</b>.\n"
            "2. Tell me what to do in plain words:\n"
            "   • \"how many rows have status active\"\n"
            "   • \"add a row: Alex, 99, active\"\n"
            "   • \"set B2 to 42\"\n"
            "   • \"sum the price column\"\n"
            "3. Send /send — I'll return the edited file as <b>.xlsx</b>.\n\n"
            "The 🌐 RU / 🌐 ENG buttons switch the interface language."
        ),
        "no_file": (
            f"Send a spreadsheet as a document first (<b>{SUPPORTED_FORMATS}</b>), "
            "then give an instruction for it."
        ),
        "unsupported": (
            f"Couldn't read that file. Supported spreadsheets: <b>{SUPPORTED_FORMATS}</b>. "
            "Please send one of these formats."
        ),
        "upload_ack": (
            "✅ Loaded <b>{filename}</b>\n"
            "Sheets: {sheets}\n"
            "Active sheet: {rows} rows × {cols} columns\n\n"
            "Tell me what to do with the table."
        ),
        "done": "Done.",
        "send_caption": "📊 Updated file",
        "lang_set": "🌐 Interface language: English. Send commands or upload a table.",
        "find_hint": (
            "🔍 <b>Find / ask about the table</b>\n\n"
            "Just type your question in plain words — e.g. \"how many rows have status "
            "active\" or \"find the row named Alex\". Send a spreadsheet as a document "
            f"first (<b>{SUPPORTED_FORMATS}</b>)."
        ),
        "reformat_hint": (
            "✏️ <b>Reformat</b>\n\n"
            "Describe in words how to reshape the table — e.g. \"move the price column "
            "to the end\", \"fill SKU by matching name\", \"set B2 to 42\". Send the "
            "file, give the instruction, then /send to get the result as .xlsx."
        ),
        "compare_unavailable": (
            "⚖️ <b>Comparing files</b> is not supported in this version yet. "
            "Upload a single table and work with it via commands."
        ),
    },
}

# Appended to the CRUD agent's system prompt so the LLM answers in the chat's
# selected language (the model follows this naturally).
_LANG_INSTRUCTION = {
    LANG_RU: "\n\nAlways write your replies to the user in Russian.",
    LANG_EN: "\n\nAlways write your replies to the user in English.",
}


# Wizard mode identifiers (persisted in redis state).
MODE_REFORMAT = "reformat"
MODE_COMPARE = "compare"

_WIZARD_STRINGS: dict[str, dict[str, str]] = {
    LANG_RU: {
        "reformat_step1": (
            "✏️ <b>Переформатирование — шаг 1/3</b>\n\n"
            "Пришли <b>файл-шаблон</b> (.xlsx) — таблицу с нужными колонками. "
            "Колонку можно пометить обязательной, обернув имя в *звёздочки*."
        ),
        "reformat_step2": (
            "✅ Шаблон принят: {cols} колонок{mandatory}.\n\n"
            "<b>Шаг 2/3</b> — пришли <b>файл с данными</b> "
            f"(любой формат: {SUPPORTED_FORMATS})."
        ),
        "reformat_step3": (
            "✅ Данные приняты ({rows} строк).\n\n"
            "<b>Шаг 3/3</b> — опиши словами, как разложить данные по колонкам шаблона "
            "(например «CID из шаблона, quantity = quantity из данных»)."
        ),
        "reformat_done": "✅ Готово: {rows} строк по шаблону.{notes}",
        "compare_step1": "⚖️ <b>Сравнение — шаг 1/3</b>\n\nПришли <b>первый файл</b>.",
        "compare_step2": "✅ Первый файл принят.\n\n<b>Шаг 2/3</b> — пришли <b>второй файл</b>.",
        "compare_step3": (
            "✅ Второй файл принят.\n\n<b>Шаг 3/3</b> — напиши, что сравнить или "
            "какой у тебя вопрос (например «почему итоговая сумма разная»)."
        ),
        "need_file": "На этом шаге нужен файл документом, не текст. Или /cancel — отменить.",
        "need_text": "На этом шаге нужен текст-инструкция. Или /cancel — отменить.",
        "bad_file": f"Не смог прочитать файл. Форматы: {SUPPORTED_FORMATS}. Или /cancel.",
        "cancelled": "❌ Отменил. Нажми кнопку, чтобы начать заново.",
        "compare_caption": "⚖️ Полная таблица расхождений",
    },
    LANG_EN: {
        "reformat_step1": (
            "✏️ <b>Reformat — step 1/3</b>\n\n"
            "Send the <b>template file</b> (.xlsx) — a table with the target columns. "
            "Mark a column mandatory by wrapping its name in *asterisks*."
        ),
        "reformat_step2": (
            "✅ Template accepted: {cols} columns{mandatory}.\n\n"
            f"<b>Step 2/3</b> — send the <b>data file</b> (any format: {SUPPORTED_FORMATS})."
        ),
        "reformat_step3": (
            "✅ Data accepted ({rows} rows).\n\n"
            "<b>Step 3/3</b> — describe in words how to map the data into the template "
            "columns (e.g. \"CID from template, quantity = quantity from data\")."
        ),
        "reformat_done": "✅ Done: {rows} rows into the template.{notes}",
        "compare_step1": "⚖️ <b>Compare — step 1/3</b>\n\nSend the <b>first file</b>.",
        "compare_step2": "✅ First file accepted.\n\n<b>Step 2/3</b> — send the <b>second file</b>.",
        "compare_step3": (
            "✅ Second file accepted.\n\n<b>Step 3/3</b> — tell me what to compare or "
            "your question (e.g. \"why are the totals different\")."
        ),
        "need_file": "This step needs a file as a document, not text. Or /cancel.",
        "need_text": "This step needs a text instruction. Or /cancel to abort.",
        "bad_file": f"Couldn't read that file. Formats: {SUPPORTED_FORMATS}. Or /cancel.",
        "cancelled": "❌ Cancelled. Tap a button to start again.",
        "compare_caption": "⚖️ Full discrepancy table",
    },
}


def t(lang: str | None, key: str, **fmt: object) -> str:
    """Return the localized string for ``key`` in ``lang`` (default RU)."""
    table = _STRINGS[normalize_lang(lang)]
    text = table[key] if key in table else _WIZARD_STRINGS[normalize_lang(lang)][key]
    return text.format(**fmt) if fmt else text


def lang_instruction(lang: str | None) -> str:
    """System-prompt suffix instructing the agent to answer in the chat's language."""
    return _LANG_INSTRUCTION[normalize_lang(lang)]


def collapse_repeat(text: str) -> str:
    """Collapse an exact back-to-back duplicated block into one copy.

    The LLM occasionally restates its pre-tool narration verbatim in its final
    answer, so the reply arrives as ``"X\\n\\nX"`` — one Telegram message with the
    same paragraph twice. This collapses only an EXACT full duplication (the whole
    string is ``block + whitespace + block``); anything else is returned unchanged.
    """
    s = (text or "").strip()
    if not s:
        return text
    m = re.fullmatch(r"(.+?)\s+\1", s, re.DOTALL)
    if m:
        return m.group(1).strip()
    n = len(s)
    if n % 2 == 0 and s[: n // 2] == s[n // 2 :]:
        return s[: n // 2].strip()
    return text
