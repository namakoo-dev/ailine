"""C1-F2: verify_dsl_args(op, args, book_meta, task, vocab) の (ok, resolved, inferred, err)
を JSON として凍結する（★エラー文言も含めて）。

★★ 本番コード(ailine.py)は一切変更しない。ここは凍結用のテスト追加のみ。

対象: 正常系・未実在列・数字表記・factor 各種（依頼文抽出/用語集/未解決の2パターン/
非正の数/LLM 食い違い警告/税込・税抜ラベル）を全 op にわたって収載する。
CLARIFY/OUT_OF_VOCAB は verify_dsl_args 自身の分岐ではない（_normalize_plan_step 側で
弾かれ、verify_dsl_args には渡らない）ため、代わりに「未対応の op」「シート無し」を
境界ケースとして収載する。

ゴールデンの置き場所: tests/golden/f2_verify/<name>.json。更新の作法は
tests/golden/_harness.py の docstring 参照。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine  # noqa: E402

from golden._harness import GOLDEN_ROOT, assert_golden_json, sorted_list  # noqa: E402

F2_DIR = GOLDEN_ROOT / "f2_verify"

_HEADERS = {"Sheet": ["商品", "数量", "単価", "金額"], "単価表": ["商品", "単価"]}
BM = {"sheets": ["Sheet", "単価表"], "headers": _HEADERS,
      "header_rows": {"Sheet": 1, "単価表": 1}}
BM_NO_SHEETS = {"sheets": [], "headers": {}, "header_rows": {}}
# LOOKUP_FILL の digit 候補が一意に決まるケース専用（列2つだけの対象シート）。
BM_LOOKUP_2COL = {"sheets": ["Sheet", "単価表"],
                   "headers": {"Sheet": ["商品", "単価"], "単価表": ["商品", "単価"]},
                   "header_rows": {"Sheet": 1, "単価表": 1}}

# 名前 -> (op, args, book_meta, task, vocab)
CASES: dict = {}


def _add(name, op, args, book_meta=BM, task="", vocab=None):
    assert name not in CASES, f"重複した case 名: {name}"
    CASES[name] = (op, args, book_meta, task, vocab or {})


# --- SORT -----------------------------------------------------------------
_add("sort_ok_by_name", "SORT", {"col": "金額", "order": "desc"})
_add("sort_ok_digit_resolve", "SORT", {"col": "3", "order": "asc"})   # 0起点3=金額
_add("sort_unknown_column", "SORT", {"col": "存在しない", "order": "asc"})
_add("sort_bad_order", "SORT", {"col": "金額", "order": "ascending"})

# --- COMPUTE_COLUMN（2列演算） ---------------------------------------------
_add("compute_column_2op_ok", "COMPUTE_COLUMN", {"operands": ["数量", "単価"], "operator": "*"})
_add("compute_column_bad_operand_count", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価", "金額"], "operator": "*"})
_add("compute_column_2op_unknown_operand", "COMPUTE_COLUMN",
     {"operands": ["数量", "存在しない"], "operator": "*"})
_add("compute_column_2op_bad_operator", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "%"})
_add("compute_column_target_ambiguous_digit", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*", "target": "1"})   # 0起点1=数量/1起点1=商品
_add("compute_column_target_unknown_falls_back_to_new_column", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*", "target": "存在しない列"})
# --- COMPUTE_COLUMN（単列×倍率） -------------------------------------------
_add("compute_column_single_bad_operator", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "+"})
_add("compute_column_single_factor_from_task_text", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*"}, task="金額に消費税10%を掛けた列を作って")
_add("compute_column_single_factor_from_vocab", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*"}, task="消費税込みの金額にして",
     vocab={"消費税": 1.1})
_add("compute_column_single_no_rate_signal_clarify", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*"}, task="金額の列の値を変えて")
_add("compute_column_single_rate_signal_unresolved_clarify", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*"}, task="税率をどうにかして")
_add("compute_column_single_factor_le0", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*"}, task="金額を0倍にした列を作って")
_add("compute_column_single_llm_factor_mismatch_warns", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*", "factor": "1.05"},
     task="金額に消費税10%を掛けた列を作って")
_add("compute_column_single_tax_inclusive_label", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*"}, task="金額に消費税10%を掛けた税込金額列を作って")
_add("compute_column_single_tax_exclusive_label", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "/"}, task="金額を1.1で割った税抜金額列を作って")
_add("compute_column_single_rate_keyword_fallback_label", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "/"}, task="倍率1.1を使った金額列を作って")
_add("compute_column_single_target_present_no_label", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "/", "target": "金額"},
     task="金額を1.1で割った税抜金額にして")

# --- LOOKUP_FILL ------------------------------------------------------------
_add("lookup_fill_ok_exists_and_mentioned", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "数量", "source_sheet": "単価表", "key_col": "商品"},
     task="数量を単価表から転記して")
_add("lookup_fill_ok_exists_matches_value_col_only", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"},
     task="単価表から値を埋めて")
_add("lookup_fill_exists_without_grounds_errors", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "数量", "source_sheet": "単価表", "key_col": "商品"},
     task="単価表から値を埋めて")
_add("lookup_fill_missing_digit_candidate_resolves", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "2", "source_sheet": "単価表", "key_col": "商品"},
     book_meta=BM_LOOKUP_2COL, task="単価表から転記して")
_add("lookup_fill_missing_and_mentioned_creates_new", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "備考", "source_sheet": "単価表", "key_col": "商品"},
     task="備考という列を作って単価表から転記して")
_add("lookup_fill_missing_not_mentioned_errors", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "備考", "source_sheet": "単価表", "key_col": "商品"},
     task="単価表から値を転記して")
_add("lookup_fill_unknown_target_sheet", "LOOKUP_FILL",
     {"target_sheet": "存在しないシート", "target_col": "単価", "source_sheet": "単価表",
      "key_col": "商品"})
_add("lookup_fill_unknown_source_sheet", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "単価", "source_sheet": "存在しないシート",
      "key_col": "商品"})
_add("lookup_fill_target_sheet_not_first", "LOOKUP_FILL",
     {"target_sheet": "単価表", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"})
_add("lookup_fill_unknown_key_col", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "単価", "source_sheet": "単価表", "key_col": "不明列"},
     task="単価転記して")

# --- AGGREGATE --------------------------------------------------------------
_add("aggregate_ok", "AGGREGATE", {"group_col": "商品", "value_col": "金額"})
_add("aggregate_unknown_group_col", "AGGREGATE", {"group_col": "不明", "value_col": "金額"})
_add("aggregate_unknown_value_col", "AGGREGATE", {"group_col": "商品", "value_col": "不明"})

# --- BOLD / FILL_COLOR / CENTER_ALIGN ---------------------------------------
_add("bold_target_all_unsupported", "BOLD", {"target": "all"})
_add("center_align_target_all_ok", "CENTER_ALIGN", {"target": "all"})
_add("bold_target_row_ok", "BOLD", {"target": "row:2"})
_add("bold_target_row_bad_number", "BOLD", {"target": "row:abc"})
_add("fill_color_target_col_ok", "FILL_COLOR", {"target": "col:金額", "color": "red"})
_add("fill_color_target_col_unknown", "FILL_COLOR", {"target": "col:存在しない", "color": "red"})
_add("bold_target_unknown_format", "BOLD", {"target": "xyz"})
_add("fill_color_unknown_color", "FILL_COLOR", {"target": "col:金額", "color": "虹色"})

# --- NUMBER_FORMAT -----------------------------------------------------------
_add("number_format_ok", "NUMBER_FORMAT", {"col": "金額", "style": "thousands"})
_add("number_format_unknown_col", "NUMBER_FORMAT", {"col": "不明", "style": "thousands"})
_add("number_format_unsupported_style", "NUMBER_FORMAT", {"col": "金額", "style": "percent"})

# --- MERGE --------------------------------------------------------------
_add("merge_ok", "MERGE", {"range": "A1:C3"})
_add("merge_bad_format", "MERGE", {"range": "A1-C3"})

# --- CHART --------------------------------------------------------------
_add("chart_ok", "CHART", {"value_col": "金額"})
_add("chart_unknown_value_col", "CHART", {"value_col": "不明"})

# --- APPEND_TOTAL ----------------------------------------------------------
_add("append_total_ok_default_factor", "APPEND_TOTAL", {"col": "金額"})
_add("append_total_ok_factor_from_task_text", "APPEND_TOTAL", {"col": "金額"},
     task="金額に消費税10%を掛けた合計を出して")
_add("append_total_ok_factor_from_vocab", "APPEND_TOTAL", {"col": "金額"},
     task="消費税込みの合計を出して", vocab={"消費税": 1.1})
_add("append_total_factor_le0", "APPEND_TOTAL", {"col": "金額"}, task="金額を0倍にした合計を出して")
_add("append_total_tax_label_unresolved_factor_errors", "APPEND_TOTAL",
     {"col": "金額", "label": "税込合計"})
_add("append_total_llm_factor_mismatch_warns", "APPEND_TOTAL",
     {"col": "金額", "factor": "1.05"}, task="金額に消費税10%を掛けた合計を出して")
_add("append_total_unknown_col", "APPEND_TOTAL", {"col": "不明"})

# --- INSERT_ROWS -----------------------------------------------------------
_add("insert_rows_ok_explicit_count", "INSERT_ROWS", {"at": "3", "count": "2"})
_add("insert_rows_ok_default_count_inferred", "INSERT_ROWS", {"at": "3"})
_add("insert_rows_bad_at", "INSERT_ROWS", {"at": "0"})
_add("insert_rows_bad_count", "INSERT_ROWS", {"at": "3", "count": "abc"})

# --- DRAW_BORDERS / AUTOFIT（引数無し） ------------------------------------
_add("draw_borders_ok", "DRAW_BORDERS", {})
_add("autofit_ok", "AUTOFIT", {})

# --- PIVOT --------------------------------------------------------------
_add("pivot_ok", "PIVOT", {"group_col": "商品", "value_col": "金額"})
_add("pivot_unknown_group_col", "PIVOT", {"group_col": "不明", "value_col": "金額"})
_add("pivot_unknown_value_col", "PIVOT", {"group_col": "商品", "value_col": "不明"})

# --- SET_COLUMN_VALUE --------------------------------------------------------
_add("set_column_value_ok_quoted", "SET_COLUMN_VALUE", {"col": "商品"},
     task="商品列を全部『確認済み』にして")
_add("set_column_value_llm_mismatch_warns", "SET_COLUMN_VALUE",
     {"col": "商品", "value": "LLM値"}, task="商品列を全部『確認済み』にして")
_add("set_column_value_no_quoted_literal_errors", "SET_COLUMN_VALUE", {"col": "商品"},
     task="商品列を全部確認済みにして")
_add("set_column_value_unknown_col", "SET_COLUMN_VALUE", {"col": "不明"},
     task="不明列を全部『確認済み』にして")

# --- 境界: verify_dsl_args 自体の全体ガード ----------------------------------
_add("no_sheets_in_book", "SORT", {"col": "金額", "order": "asc"}, book_meta=BM_NO_SHEETS)
_add("unsupported_op", "FOOBAR", {})


def _case_ids():
    return sorted(CASES.keys())


def test_verify_dsl_args_golden_coverage_matches_declared_ops():
    """★網羅性の自己検査: OP_WRITE_TARGET に載っている全 op が最低1ケース入っているか。"""
    covered_ops = {op for op, *_ in CASES.values()}
    all_ops = set(ailine.OP_WRITE_TARGET.keys())
    missing = all_ops - covered_ops
    assert not missing, f"verify_dsl_args golden が網羅していない op: {sorted(missing)}"


@pytest.mark.parametrize("name", _case_ids())
def test_verify_dsl_args_golden(name):
    op, args, book_meta, task, vocab = CASES[name]
    ok, resolved, inferred, err = ailine.verify_dsl_args(op, dict(args), book_meta, task, vocab)
    payload = {"ok": ok, "resolved": resolved, "inferred": sorted_list(inferred), "err": err}
    assert_golden_json(F2_DIR / f"{name}.json", payload, label=name)
