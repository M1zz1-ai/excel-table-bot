"""Tiered key-matcher tests (normalization / fuzzy / mocked LLM tier).

No network: tier 3 (LLM residual pairing) is injected as a callable and mocked,
so the RU<->UA translation case is covered deterministically.
"""

from __future__ import annotations

from excel import engine

# ---- tier 1: normalization ---------------------------------------------


def test_normalize_lowercases_and_collapses_whitespace():
    assert engine.normalize_key("  Apple   Juice  ") == "apple juice"


def test_normalize_yo_to_e():
    assert engine.normalize_key("Гвоздёв") == engine.normalize_key("Гвоздев")


def test_normalize_strips_punctuation():
    assert engine.normalize_key("Болт, М6 (оцинк.)") == engine.normalize_key("Болт М6 оцинк")


def test_normalize_unifies_latin_cyrillic_lookalikes():
    # 'Аpple' typed with a Cyrillic 'А' must normalize to the same key as latin.
    cyr = "Аpple"  # Cyrillic A + latin pple
    assert engine.normalize_key(cyr) == engine.normalize_key("Apple")


def test_normalize_strips_trailing_unit_noise():
    assert engine.normalize_key("Сахар 1 кг") == engine.normalize_key("Сахар 1")


# ---- tier 1 exact-on-normalized ----------------------------------------


def test_match_keys_exact_on_normalized():
    res = engine.match_keys(["Apple Juice"], ["apple  juice"])
    assert len(res["matched"]) == 1
    pair = res["matched"][0]
    assert pair["method"] == "exact"
    assert pair["confidence"] == 1.0
    assert pair["key_a"] == "Apple Juice" and pair["key_b"] == "apple  juice"
    assert res["unmatched_a"] == [] and res["unmatched_b"] == []


# ---- tier 2: fuzzy ------------------------------------------------------


def test_match_keys_fuzzy_typo():
    # A typo below exact but above threshold should fuzzy-match.
    res = engine.match_keys(["Гвоздь строительный 100мм"], ["Гвозди строительные 100 мм"])
    assert len(res["matched"]) == 1
    pair = res["matched"][0]
    assert pair["method"] == "fuzzy"
    assert 0.85 <= pair["confidence"] <= 1.0


def test_match_keys_fuzzy_is_one_to_one():
    # Two A keys close to one B key: only the best pair is taken, the other is unmatched.
    res = engine.match_keys(
        ["Болт М6", "Болт М6 оцинкованный"],
        ["Болт М6 оцинкованный"],
    )
    assert len(res["matched"]) == 1
    assert res["matched"][0]["key_a"] == "Болт М6 оцинкованный"
    assert res["unmatched_a"] == ["Болт М6"]


def test_match_keys_below_threshold_stays_unmatched():
    res = engine.match_keys(["Apple"], ["Bicycle"])
    assert res["matched"] == []
    assert res["unmatched_a"] == ["Apple"] and res["unmatched_b"] == ["Bicycle"]


# ---- tier 3: LLM residual pairing (mocked) -----------------------------


def test_match_keys_llm_tier_pairs_ru_ua():
    calls: list[tuple[list[str], list[str]]] = []

    def fake_llm(keys_a: list[str], keys_b: list[str]) -> list[dict[str, str]]:
        calls.append((keys_a, keys_b))
        return [{"a": "гвоздь", "b": "цвях"}]

    res = engine.match_keys(["гвоздь"], ["цвях"], llm_pair=fake_llm)
    assert len(calls) == 1  # tier 3 was invoked on the residual
    assert len(res["matched"]) == 1
    pair = res["matched"][0]
    assert pair["method"] == "llm"
    assert pair["key_a"] == "гвоздь" and pair["key_b"] == "цвях"
    assert res["unmatched_a"] == [] and res["unmatched_b"] == []
    assert res["match_counts"] == {"exact": 0, "fuzzy": 0, "llm": 1}


def test_match_keys_llm_drops_hallucinated_pairs():
    def fake_llm(keys_a: list[str], keys_b: list[str]) -> list[dict[str, str]]:
        # 'phantom' is not in either input list -> must be dropped.
        return [{"a": "phantom", "b": "цвях"}, {"a": "гвоздь", "b": "ghost"}]

    res = engine.match_keys(["гвоздь"], ["цвях"], llm_pair=fake_llm)
    assert res["matched"] == []
    assert res["unmatched_a"] == ["гвоздь"] and res["unmatched_b"] == ["цвях"]


def test_match_keys_llm_enforces_one_to_one():
    def fake_llm(keys_a: list[str], keys_b: list[str]) -> list[dict[str, str]]:
        # Model tries to reuse 'y' for two A keys -> only the first wins.
        return [{"a": "x1", "b": "y"}, {"a": "x2", "b": "y"}]

    res = engine.match_keys(["x1", "x2"], ["y"], llm_pair=fake_llm)
    assert len(res["matched"]) == 1
    assert res["matched"][0]["key_a"] == "x1"
    assert res["unmatched_a"] == ["x2"]


def test_match_keys_skips_llm_when_too_many_keys():
    called = False

    def fake_llm(keys_a: list[str], keys_b: list[str]) -> list[dict[str, str]]:
        nonlocal called
        called = True
        return []

    big_a = [f"a-{i}" for i in range(engine.LLM_MAX_KEYS + 1)]
    big_b = [f"b-{i}" for i in range(engine.LLM_MAX_KEYS + 1)]
    res = engine.match_keys(big_a, big_b, llm_pair=fake_llm)
    assert called is False
    assert res["llm_skipped"] is True


def test_match_keys_no_llm_leaves_residual_unmatched():
    res = engine.match_keys(["гвоздь"], ["цвях"])  # no llm_pair injected
    assert res["matched"] == []
    assert res["llm_skipped"] is False  # not skipped-by-cap; simply no LLM available


# ---- tiered diff integration -------------------------------------------


def _plan() -> dict:
    return {
        "key_a": "name",
        "key_b": "name",
        "compare_columns": [{"a": "qty", "b": "qty", "label": "qty"}],
        "sum_column_a": "qty",
        "sum_column_b": "qty",
    }


def test_diff_tables_uses_fuzzy_and_llm_join():
    rows_a = [
        {"name": "Apple Juice", "qty": "10"},          # exact (after normalize)
        {"name": "Молоток слесарный", "qty": "5"},      # fuzzy (typo variant)
        {"name": "гвоздь", "qty": "7"},                # llm -> цвях
        {"name": "Уникальный товар А", "qty": "1"},     # only in A
    ]
    rows_b = [
        {"name": "apple juice", "qty": "10"},
        {"name": "Молоток слесарской", "qty": "6"},     # fuzzy, qty mismatch
        {"name": "цвях", "qty": "7"},
        {"name": "Совсем другое Б", "qty": "2"},        # only in B
    ]

    def fake_llm(keys_a: list[str], keys_b: list[str]) -> list[dict[str, str]]:
        return [{"a": "гвоздь", "b": "цвях"}]

    diff = engine.diff_tables(rows_a, rows_b, _plan(), llm_pair=fake_llm)
    assert diff["match_counts"] == {"exact": 1, "fuzzy": 1, "llm": 1}
    assert [r["key"] for r in diff["only_in_a"]] == ["Уникальный товар А"]
    assert [r["key"] for r in diff["only_in_b"]] == ["Совсем другое Б"]
    # only the fuzzy pair has a qty mismatch (5 vs 6)
    assert len(diff["mismatches"]) == 1
    m = diff["mismatches"][0]
    assert m["match_method"] == "fuzzy"
    assert m["key_a"] == "Молоток слесарный" and m["key_b"] == "Молоток слесарской"


def test_resolve_join_keys_prefers_code_over_language_names():
    # Plan picks the descriptive NAME column, which fails across RU/UA. The resolver
    # must switch to the article-code column (unique + high overlap).
    rows_a = [
        {"code": "SKU-001", "name": "Швидкий віск"},
        {"code": "SKU-003", "name": "Автошампунь"},
        {"code": "SKU-100", "name": "Ароматизатор"},
    ]
    rows_b = [
        {"code": "SKU-001", "name": "Быстрый воск"},
        {"code": "SKU-003", "name": "Автошампунь концентрат"},
        {"code": "SKU-100", "name": "Ароматизатор для авто"},
    ]
    plan = {"key_a": "name", "key_b": "name", "compare_columns": [], "sum_column_a": "", "sum_column_b": ""}
    assert engine.resolve_join_keys(rows_a, rows_b, plan) == ("code", "code")


def test_resolve_join_keys_rejects_row_ordinal_column():
    # A '№' row-counter overlaps perfectly (1,2,3…) but must NOT become the key;
    # the real code column (different values, high overlap) must win instead.
    rows_a = [
        {"№": 1, "code": "SKU-001", "name": "Швидкий віск"},
        {"№": 2, "code": "SKU-003", "name": "Автошампунь"},
        {"№": 3, "code": "SKU-100", "name": "Ароматизатор"},
    ]
    rows_b = [
        {"№": 1, "code": "SKU-001", "name": "Быстрый воск"},
        {"№": 2, "code": "SKU-003", "name": "Автошампунь концентрат"},
        {"№": 3, "code": "SKU-100", "name": "Освежитель воздуха"},
    ]
    plan = {"key_a": "name", "key_b": "name", "compare_columns": [], "sum_column_a": "", "sum_column_b": ""}
    assert engine.resolve_join_keys(rows_a, rows_b, plan) == ("code", "code")


def test_resolve_join_keys_keeps_good_planned_key():
    rows_a = [{"id": "1", "x": "a"}, {"id": "2", "x": "b"}]
    rows_b = [{"id": "1", "x": "a"}, {"id": "2", "x": "c"}]
    plan = {"key_a": "id", "key_b": "id", "compare_columns": [], "sum_column_a": "", "sum_column_b": ""}
    assert engine.resolve_join_keys(rows_a, rows_b, plan) == ("id", "id")


def test_resolve_join_keys_ignores_compare_value_column():
    # A value column that coincidentally overlaps must NOT be chosen as the key.
    rows_a = [{"name": "гвоздь", "amount": "1"}, {"name": "болт", "amount": "2"}]
    rows_b = [{"name": "цвях", "amount": "1"}, {"name": "гвинт", "amount": "2"}]
    plan = {
        "key_a": "name", "key_b": "name",
        "compare_columns": [{"a": "amount", "b": "amount", "label": "amount"}],
        "sum_column_a": "amount", "sum_column_b": "amount",
    }
    assert engine.resolve_join_keys(rows_a, rows_b, plan) == ("name", "name")


def _recon_plan() -> dict:
    return {
        "key_a": "code", "key_b": "code",
        "compare_columns": [
            {"a": "qty", "b": "qty", "label": "кол-во"},
            {"a": "sum", "b": "sum", "label": "сумма"},
        ],
        "sum_column_a": "sum", "sum_column_b": "sum",
    }


def test_reconcile_sum_all_three_cases_total_to_delta():
    rows_a = [
        {"code": "A1", "name": "Товар1", "qty": "24", "sum": "48"},   # matched, sum differs
        {"code": "A2", "name": "Товар2", "qty": "1", "sum": "10"},    # only in A
        {"code": "A3", "name": "Товар3", "qty": "5", "sum": "20"},    # matched, equal
    ]
    rows_b = [
        {"code": "A1", "name": "Товар1", "qty": "20", "sum": "40"},   # -> contribution +8
        {"code": "A3", "name": "Товар3", "qty": "5", "sum": "20"},    # equal -> 0
        {"code": "B9", "name": "ТоварB", "qty": "2", "sum": "15"},    # only in B -> -15
    ]
    diff = engine.diff_tables(rows_a, rows_b, _recon_plan())
    recon = diff["reconciliation"]
    total = round(sum(c["contribution"] for c in recon), 2)
    assert total == diff["delta"]  # contributions reconcile to Δ exactly
    by_code = {c["code"]: c for c in recon}
    assert by_code["A1"]["contribution"] == 8.0 and "кол-во" in by_code["A1"]["reason"]
    assert by_code["A2"]["contribution"] == 10.0 and by_code["A2"]["reason"] == "есть только в A"
    assert by_code["B9"]["contribution"] == -15.0 and by_code["B9"]["reason"] == "есть только в B"
    assert "A3" not in by_code  # equal-sum matched product is not a contributor
    # sorted by absolute impact, largest first
    assert [c["code"] for c in recon] == ["B9", "A2", "A1"]


def test_compare_discrepancy_block_lists_top_products():
    diff = engine.diff_tables(
        [{"code": "A1", "name": "Товар1", "qty": "24", "sum": "48"},
         {"code": "A2", "name": "Товар2", "qty": "1", "sum": "10"}],
        [{"code": "A1", "name": "Товар1", "qty": "20", "sum": "40"}],
        _recon_plan(),
    )
    block = engine.compare_discrepancy_block(diff, top=5)
    assert block.startswith("Товары с расхождениями")
    assert "Товар1 (A1)" in block  # product name + code, deterministic
    assert "Всего расхождений:" in block and 'лист "Расхождения"' in block
    # residual line (no code) must not appear as a bullet
    assert "прочее" not in block


def test_reconcile_sum_residual_line_when_rounding_gap():
    # Two matched products each off by 0.005 -> per-product rounds to 0, but Δ is 0.01.
    rows_a = [{"code": "A1", "name": "T1", "sum": "10.005"}, {"code": "A2", "name": "T2", "sum": "10.005"}]
    rows_b = [{"code": "A1", "name": "T1", "sum": "10.00"}, {"code": "A2", "name": "T2", "sum": "10.00"}]
    plan = {"key_a": "code", "key_b": "code", "compare_columns": [{"a": "sum", "b": "sum", "label": "сумма"}],
            "sum_column_a": "sum", "sum_column_b": "sum"}
    diff = engine.diff_tables(rows_a, rows_b, plan)
    recon = diff["reconciliation"]
    total = round(sum(c["contribution"] for c in recon), 2)
    assert total == diff["delta"]
    assert any(c["reason"] == "остаток" for c in recon)  # residual line present


def test_diff_tables_matched_carries_method_and_confidence():
    rows_a = [{"name": "Apple", "qty": "1"}]
    rows_b = [{"name": "apple", "qty": "1"}]
    diff = engine.diff_tables(rows_a, rows_b, _plan())
    assert len(diff["matched"]) == 1
    assert diff["matched"][0]["method"] == "exact"
    assert diff["matched"][0]["confidence"] == 1.0
