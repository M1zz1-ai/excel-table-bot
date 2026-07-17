"""Excel-bot handlers: upload a file, talk to it (CRUD), get it back.

Replicates the essence of the n8n "Excel Tables · TG Bot" (9O36SPWrJ3hpOzAQ) —
upload a table (.xlsx/.xlsm/.xls/.csv/.tsv, normalized to .xlsx on load) ->
conversational find / edit / add / reshape over an OpenAI agent -> send the edited
file back — but collapses the 13-code-node state machine into
ONE tool-calling agent (the big win vs n8n). The agent's tools are bound to the
chat's working workbook (see :mod:`excel.tools` / :mod:`excel.session`).

Every turn runs inside ``core.errors.run_resilient`` so an OpenAI/openpyxl
failure pings the user and is logged, but never kills the long-poll process.
"""

from __future__ import annotations

import logging
from typing import Any, Protocol

import openai

from core import openai_agent as core_agent
from core import sheets
from core.errors import SheetError, run_resilient

from . import engine, i18n, tools
from .session import ExcelSession
from .tools import EXCEL_MODEL

logger = logging.getLogger(__name__)


class _TgLike(Protocol):
    async def send_text(self, text: str, chat_id: int | None = ..., **kw: Any) -> Any: ...
    async def send_document(self, document: Any, chat_id: int | None = ..., **kw: Any) -> Any: ...


class ExcelBot:
    """Wires the shared core into upload / chat / send handlers.

    Args:
        telegram: a ``core.tg.TelegramClient`` (or compatible).
        client: an ``openai.OpenAI`` instance (the agent is built per chat
            so its tools bind to that chat's workbook).
        session: an :class:`ExcelSession` for per-chat file lifecycle.
        owner_chat_id: chat id used for failure alerts.
    """

    def __init__(
        self,
        telegram: _TgLike,
        client: openai.OpenAI,
        session: ExcelSession,
        owner_chat_id: int,
    ) -> None:
        self._tg = telegram
        self._client = client
        self._session = session
        self._owner = owner_chat_id

    # ---- handlers -------------------------------------------------------

    def _lang(self, chat_id: int) -> str:
        """The chat's UI language code (normalized; default RU)."""
        return i18n.normalize_lang(self._session.get_lang(chat_id))

    def _menu(self) -> Any:
        """Persistent reply keyboard mirroring the original n8n layout (3 rows:
        Find/Compare/Reformat, Help, RU/ENG).

        Imported lazily so the pure handler logic stays importable without aiogram.
        """
        from core.tg import reply_keyboard

        return reply_keyboard([list(row) for row in i18n.MENU_ROWS])

    async def on_start(self, chat_id: int) -> None:
        """Render the welcome message and (re)attach the Help/RU/ENG keyboard."""
        self._session.wizard_clear(chat_id)  # a new command resets any wizard
        lang = self._lang(chat_id)
        await self._tg.send_text(
            i18n.t(lang, "welcome"), chat_id=chat_id, reply_markup=self._menu()
        )

    async def _handle_button(self, chat_id: int, action: str) -> None:
        """Handle any menu button press. NEVER reaches the LLM CRUD agent.

        A button press always resets any in-progress wizard first. Reformat /
        Compare START their multi-step wizards; RU/ENG mutate the language;
        Find/Help send a localized simple reply.
        """
        self._session.wizard_clear(chat_id)
        if action in (i18n.ACTION_LANG_RU, i18n.ACTION_LANG_EN):
            new_lang = i18n.LANG_RU if action == i18n.ACTION_LANG_RU else i18n.LANG_EN
            self._session.set_lang(chat_id, new_lang)
            await self._tg.send_text(
                i18n.t(new_lang, "lang_set"), chat_id=chat_id, reply_markup=self._menu()
            )
            return
        if action == i18n.ACTION_REFORMAT:
            await self._start_wizard(chat_id, i18n.MODE_REFORMAT, "reformat_step1")
            return
        if action == i18n.ACTION_COMPARE:
            await self._start_wizard(chat_id, i18n.MODE_COMPARE, "compare_step1")
            return
        string_key = i18n.SIMPLE_ACTION_STRING[action]
        await self._tg.send_text(
            i18n.t(self._lang(chat_id), string_key), chat_id=chat_id, reply_markup=self._menu()
        )

    async def _start_wizard(self, chat_id: int, mode: str, step1_key: str) -> None:
        """Initialise a wizard flow at step 1 and prompt for the first file."""
        self._session.wizard_set(chat_id, {"mode": mode, "step": 1, "files": {}})
        await self._tg.send_text(i18n.t(self._lang(chat_id), step1_key), chat_id=chat_id)

    async def on_cancel(self, chat_id: int) -> None:
        """Cancel an in-progress wizard (/cancel)."""
        self._session.wizard_clear(chat_id)
        await self._tg.send_text(
            i18n.t(self._lang(chat_id), "cancelled"), chat_id=chat_id, reply_markup=self._menu()
        )

    async def on_document(self, chat_id: int, data: bytes, filename: str) -> None:
        """Store an uploaded table (any supported format) and ACK its shape.

        Unsupported/corrupt files get a clear localized reply listing the formats,
        not a generic failure alert. When a wizard is running, the file feeds the
        current wizard step instead of the CRUD workbook.
        """
        if self._session.wizard_get(chat_id) is not None:
            await run_resilient(
                lambda: self._wizard_document(chat_id, data, filename),
                alerter=self._alerter(chat_id),
                label="wizard file",
            )
            return
        lang = self._lang(chat_id)

        async def work() -> None:
            try:
                self._session.load_from_bytes(chat_id, data, filename)
            except SheetError:
                logger.info("excel: rejected unsupported upload %r", filename)
                await self._tg.send_text(i18n.t(lang, "unsupported"), chat_id=chat_id)
                return
            wb = self._session.workbook(chat_id)
            ws = wb.active
            await self._tg.send_text(
                i18n.t(
                    lang,
                    "upload_ack",
                    filename=filename,
                    sheets=", ".join(wb.sheetnames),
                    rows=ws.max_row,
                    cols=ws.max_column,
                ),
                chat_id=chat_id,
            )

        await run_resilient(work, alerter=self._alerter(chat_id), label="file upload")

    async def on_text(self, chat_id: int, text: str) -> None:
        """Route a text message: menu button, else the agent CRUD loop.

        A Help/RU/ENG button press is handled here and MUST NOT reach the LLM
        agent (that was the bug: buttons falling through to the CRUD agent).
        """
        action = i18n.classify_button(text)
        if action is not None:
            await self._handle_button(chat_id, action)
            return

        if self._session.wizard_get(chat_id) is not None:
            await run_resilient(
                lambda: self._wizard_text(chat_id, text),
                alerter=self._alerter(chat_id),
                label="wizard step",
            )
            return

        if not self._session.has_file(chat_id):
            await self._tg.send_text(i18n.t(self._lang(chat_id), "no_file"), chat_id=chat_id)
            return

        lang = self._lang(chat_id)

        async def work() -> None:
            reply = self._run_agent(chat_id, text)
            await self._tg.send_text(reply or i18n.t(lang, "done"), chat_id=chat_id)

        await run_resilient(work, alerter=self._alerter(chat_id), label="agent turn")

    async def on_send(self, chat_id: int) -> None:
        """Send the chat's current working file back as a document."""
        self._session.wizard_clear(chat_id)  # a new command resets any wizard
        lang = self._lang(chat_id)
        if not self._session.has_file(chat_id):
            await self._tg.send_text(i18n.t(lang, "no_file"), chat_id=chat_id)
            return

        async def work() -> None:
            data = self._session.read_bytes(chat_id)
            filename = self._session.get_filename(chat_id)
            await self._tg.send_document(
                data, chat_id=chat_id, caption=i18n.t(lang, "send_caption"), filename=filename
            )

        await run_resilient(work, alerter=self._alerter(chat_id), label="send file")

    # ---- wizards (reformat / compare) -----------------------------------

    def _plan_agent(self, system: str) -> core_agent.OpenAIAgent:
        """A tool-less agent for wizard planning / short answers."""
        return core_agent.OpenAIAgent(self._client, system=system, model=EXCEL_MODEL)

    def _load_wb(self, data: bytes, filename: str):
        """Parse uploaded bytes into a workbook (any supported format)."""
        return sheets.read_bytes_to_workbook(data, filename=filename)

    async def _wizard_document(self, chat_id: int, data: bytes, filename: str) -> None:
        """Accept a file for the current wizard step; advance or nudge."""
        state = self._session.wizard_get(chat_id)
        if state is None:
            return
        lang = self._lang(chat_id)
        step = state["step"]
        if step == 3:  # step 3 wants text, not a file
            await self._tg.send_text(i18n.t(lang, "need_text"), chat_id=chat_id)
            return
        try:
            wb = self._load_wb(data, filename)
        except SheetError:
            await self._tg.send_text(i18n.t(lang, "bad_file"), chat_id=chat_id)
            return

        files = state.get("files", {})
        if state["mode"] == i18n.MODE_REFORMAT and step == 1:
            template = engine.parse_template_planned(wb, plan_fn=self._table_plan)
            files["tpl"] = self._session.wizard_save_file(chat_id, "tpl", data, filename)
            n_mand = sum(1 for h in template["headers"] if h["mandatory"])
            mand = f", {n_mand} обязательных/mandatory" if n_mand else ""
            state.update(step=2, template=template, files=files)
            self._session.wizard_set(chat_id, state)
            await self._tg.send_text(
                i18n.t(lang, "reformat_step2", cols=len(template["headers"]), mandatory=mand),
                chat_id=chat_id,
            )
        elif state["mode"] == i18n.MODE_REFORMAT and step == 2:
            # Plan the data grid ONCE here (planned path, same as the actual run) and
            # stash the plan so step 3 reuses it — no second table_plan LLM call.
            grid = engine.sheet_rows(wb)
            data_plan: dict[str, Any] | None = None
            rows: list[dict[str, Any]] = []
            try:
                data_plan = self._table_plan(grid)
                _, rows = engine.extract_by_plan(grid, data_plan)
            except Exception:
                logger.warning("reformat step2 table_plan failed", exc_info=True)
            if not rows:
                data_plan = None
                _, rows = engine.load_table(wb)
            files["data"] = self._session.wizard_save_file(chat_id, "data", data, filename)
            state.update(step=3, data_filename=filename, files=files, data_plan=data_plan)
            self._session.wizard_set(chat_id, state)
            await self._tg.send_text(i18n.t(lang, "reformat_step3", rows=len(rows)), chat_id=chat_id)
        elif state["mode"] == i18n.MODE_COMPARE and step in (1, 2):
            role = "a" if step == 1 else "b"
            files[role] = self._session.wizard_save_file(chat_id, role, data, filename)
            files[f"{role}_name"] = filename
            key = "compare_step2" if step == 1 else "compare_step3"
            state.update(step=step + 1, files=files)
            self._session.wizard_set(chat_id, state)
            await self._tg.send_text(i18n.t(lang, key), chat_id=chat_id)

    async def _wizard_text(self, chat_id: int, text: str) -> None:
        """Step-3 instruction/question: run the deterministic engine + LLM plan."""
        state = self._session.wizard_get(chat_id)
        if state is None:
            return
        lang = self._lang(chat_id)
        if state["step"] != 3:  # earlier steps expect a file
            await self._tg.send_text(i18n.t(lang, "need_file"), chat_id=chat_id)
            return
        if state["mode"] == i18n.MODE_REFORMAT:
            await self._run_reformat(chat_id, state, text, lang)
        else:
            await self._run_compare(chat_id, state, text, lang)

    async def _run_reformat(self, chat_id, state, instruction, lang) -> None:
        from datetime import date
        from pathlib import Path

        template = state["template"]
        wb = self._load_wb(Path(state["files"]["data"]).read_bytes(), state["data_filename"])
        grid = engine.sheet_rows(wb)
        data_plan = state.get("data_plan")
        if data_plan:  # reuse the plan from step 2 (no extra LLM call)
            headers, rows = engine.extract_by_plan(grid, data_plan)
            if not rows:
                headers, rows = engine.load_table(wb)
        else:
            headers, rows = engine.load_table(wb)

        plan = self._plan_agent(engine.REFORMAT_SYSTEM).structured_output(
            engine.reformat_prompt(template, headers, rows[:15], instruction), engine.REFORMAT_SCHEMA
        )
        out_rows, notes = engine.apply_mapping(template, rows, plan)
        quality = engine.reformat_quality(notes)
        logger.info(
            "reformat chat=%s rows=%s quality=%s overall_fill=%.0f%% mapping=%s unmapped=%s",
            chat_id, notes["row_count"], quality, notes["overall_fill_pct"] * 100,
            notes["mapping_pairs"], notes["unmapped_columns"],
        )

        # No silent empty output: refuse a useless file, explain what to do.
        if quality == "empty":
            self._session.wizard_clear(chat_id)  # nothing usable; user restarts if wanted
            await self._tg.send_text(
                self._reformat_explanation(template, headers, notes, sent_file=False),
                chat_id=chat_id, reply_markup=self._menu(),
            )
            return

        xlsx = engine.build_reformat_xlsx(template, out_rows)
        fname = f"reformatted_data_{date.today().isoformat()}.xlsx"
        await self._tg.send_document(xlsx, chat_id=chat_id, caption=fname, filename=fname)

        if quality == "weak":
            # File went out, but keep the wizard open so a re-worded instruction re-runs.
            await self._tg.send_text(
                self._reformat_explanation(template, headers, notes, sent_file=True),
                chat_id=chat_id, reply_markup=self._menu(),
            )
            return

        note_bits = []
        if notes["unmapped_mandatory"]:
            note_bits.append("нет маппинга (укажи значение): " + ", ".join(notes["unmapped_mandatory"]))
        if notes["fuzzy_fields"]:
            note_bits.append("приблизительно: " + ", ".join(notes["fuzzy_fields"]))
        if notes["duplicate_keys"]:
            note_bits.append(f"дубликатов ключа: {notes['duplicate_keys']}")
        notes_str = ("\n— " + "\n— ".join(note_bits)) if note_bits else ""
        self._session.wizard_clear(chat_id)
        await self._tg.send_text(
            i18n.t(lang, "reformat_done", rows=notes["row_count"], notes=notes_str),
            chat_id=chat_id, reply_markup=self._menu(),
        )

    @staticmethod
    def _reformat_explanation(
        template: dict[str, Any], source_headers: list[str], notes: dict[str, Any], *, sent_file: bool
    ) -> str:
        """Russian diagnostic when the reformat mapping is empty/weak (no silent dud)."""
        tpl_cols = [h["name"] for h in template["headers"]]
        if sent_file:
            head = (
                f"⚠️ Файл собрал, но заполнено мало — {int(notes['overall_fill_pct'] * 100)}% ячеек. "
                "Проверьте и при желании уточните сопоставление."
            )
        else:
            head = "⚠️ Не удалось сопоставить ни одной колонки — пустой файл отправлять не стал."
        lines = [head]
        if notes["unmapped_columns"]:
            lines.append("Без сопоставления: " + ", ".join(notes["unmapped_columns"]))
        lines.append("Колонки шаблона: " + ", ".join(tpl_cols))
        lines.append("Колонки в вашем файле: " + ", ".join(source_headers or ["(не распознаны)"]))
        lines.append(
            "Опишите словами, какая колонка файла идёт в какую колонку шаблона, "
            "и пришлите инструкцию ещё раз."
        )
        return "\n".join(lines)

    async def _run_compare(self, chat_id, state, question, lang) -> None:
        from pathlib import Path

        fa, fb = state["files"]["a"], state["files"]["b"]
        wb_a = self._load_wb(Path(fa).read_bytes(), state["files"]["a_name"])
        wb_b = self._load_wb(Path(fb).read_bytes(), state["files"]["b_name"])
        headers_a, rows_a = engine.load_table_planned(wb_a, plan_fn=self._table_plan)
        headers_b, rows_b = engine.load_table_planned(wb_b, plan_fn=self._table_plan)
        plan = self._plan_agent(engine.COMPARE_SYSTEM).structured_output(
            engine.compare_prompt(headers_a, headers_b, rows_a[:10], rows_b[:10], question),
            engine.COMPARE_SCHEMA,
        )
        diff = engine.diff_tables(rows_a, rows_b, plan, llm_pair=self._llm_pair)
        xlsx = engine.build_compare_xlsx(diff)
        answer = self._plan_agent(engine.COMPARE_ANSWER_SYSTEM + i18n.lang_instruction(lang)).run(
            f"Question: {question}\nDiff stats: {engine.compare_stats_line(diff)}\n"
            f"Reconciliation of the total difference:\n{engine.reconcile_line(diff)}"
        )
        mc = diff["match_counts"]
        # Deterministic top-discrepancy list FIRST (the thing the user actually wants),
        # then the aggregate stats, then the LLM narration.
        block = engine.compare_discrepancy_block(diff)
        stats = (
            f"⚖️ Σ A = {diff['sum_a']} | Σ B = {diff['sum_b']} | Δ = {diff['delta']}\n"
            f"Строк A={diff['count_a']} B={diff['count_b']}; "
            f"совпало {len(diff['matched'])} "
            f"(точно {mc['exact']}, похоже {mc['fuzzy']}, по смыслу {mc['llm']}); "
            f"только в A={len(diff['only_in_a'])}, только в B={len(diff['only_in_b'])}; "
            f"расхождений полей={len(diff['mismatches'])}"
        )
        summary = "\n\n".join(p for p in (block, stats, answer) if p).strip()
        self._session.wizard_clear(chat_id)
        await self._tg.send_document(
            xlsx, chat_id=chat_id, caption=i18n.t(lang, "compare_caption"),
            filename="compare_report.xlsx",
        )
        await self._tg.send_text(summary, chat_id=chat_id, reply_markup=self._menu())

    def _table_plan(self, rows: list[list[Any]]) -> dict[str, Any]:
        """LLM structure plan for a raw grid (engine extracts by column index).

        Injected into ``engine.load_table_planned`` for compare sides and the
        reformat source, so messy print/accounting exports (sparse columns,
        horizontal 2-up duplication, metadata rows) are parsed correctly instead
        of by positional heuristics. The engine falls back to heuristics on failure.
        """
        return self._plan_agent(engine.TABLE_PLAN_SYSTEM).structured_output(
            engine.table_plan_prompt(rows), engine.TABLE_PLAN_SCHEMA
        )

    def _llm_pair(self, keys_a: list[str], keys_b: list[str]) -> list[dict[str, str]]:
        """Tier-3 residual key pairing for Compare (RU<->UA / renames).

        The engine calls this only for keys still unmatched after exact + fuzzy;
        it stays here (not in the pure engine) so the real model call lives with
        the bot. Returns ``[{"a": .., "b": ..}]``; the engine validates the keys.
        """
        result = self._plan_agent(engine.LLM_PAIR_SYSTEM).structured_output(
            engine.llm_pair_prompt(keys_a, keys_b), engine.LLM_PAIR_SCHEMA
        )
        pairs = result.get("pairs", [])
        return [p for p in pairs if isinstance(p, dict) and "a" in p and "b" in p]

    # ---- helpers --------------------------------------------------------

    def _build_agent(self, chat_id: int) -> core_agent.OpenAIAgent:
        """Build a fresh agent whose sheet-op tools are bound to this chat's file.

        A new agent per turn keeps tool closures pointing at the current
        workbook and avoids cross-chat state leaking through one shared agent.
        The system prompt is suffixed with the chat's language instruction.
        """
        system = tools.AGENT_SYSTEM + i18n.lang_instruction(self._lang(chat_id))
        brain = core_agent.OpenAIAgent(self._client, system=system, model=EXCEL_MODEL)
        tools.register_tools(
            brain,
            lambda: self._session.workbook(chat_id),
            on_change=lambda: self._session.flush(chat_id),
        )
        return brain

    def _run_agent(self, chat_id: int, text: str) -> str:
        """Run the tool-calling loop; return the assistant's final text.

        The reply is passed through :func:`i18n.collapse_repeat` because the model
        occasionally restates its answer verbatim (its pre-tool narration echoed
        in the final turn), which otherwise reaches the user as one doubled message.
        """
        brain = self._build_agent(chat_id)
        return i18n.collapse_repeat(brain.run(text))

    def _alerter(self, chat_id: int) -> Any:
        """Adapt the tg client to the Alerter protocol, pinning chat_id."""
        tg_client = self._tg
        chat_id_outer = chat_id

        class _Alerter:
            async def send_text(self, text: str, chat_id: int | None = None) -> Any:
                return await tg_client.send_text(text, chat_id=chat_id_outer)

        return _Alerter()
