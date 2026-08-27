"""C1-F1: codegen_dsl(op, resolved, book_meta, use_formula) の .bas 文字列をバイト単位で凍結する。

★★ 本番コード(ailine.py)は一切変更しない。ここは凍結用のテスト追加のみ。

対象: 全17 op（SORT/COMPUTE_COLUMN/LOOKUP_FILL/AGGREGATE/BOLD/FILL_COLOR/NUMBER_FORMAT/
MERGE/CHART/CENTER_ALIGN/APPEND_TOTAL/INSERT_ROWS/DRAW_BORDERS/AUTOFIT/PIVOT/
SET_COLUMN_VALUE/EXTRACT）。各 op について、生成に効く軸（use_formula・header_row・target有無・
target 種別 row:/col:/all）が意味を持つ場合だけ変化させる。効かない軸（例: BOLD に
use_formula を渡しても出力が変わらない）は 1 通りだけ回して「不変であること」も
ゴールデンで検証する（同じ内容が 2 通り生成されて golden も同一になる）。

ゴールデンの置き場所: tests/golden/f1_codegen/<name>.bas （.bas そのものが人間可読）。
更新の作法は tests/golden/_harness.py の docstring 参照
（AILINE_REGEN_GOLDEN=1 で再生成 → git diff を人が読んで承認 → commit）。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import ailine  # noqa: E402

from golden._harness import GOLDEN_ROOT, assert_golden_bytes  # noqa: E402

F1_DIR = GOLDEN_ROOT / "f1_codegen"

# --- book_meta 2 通り（header_row = 1 / 3）------------------------------------
_HEADERS = {"Sheet": ["商品", "数量", "単価", "金額"], "単価表": ["商品", "単価"]}

BM_HR1 = {"sheets": ["Sheet", "単価表"], "headers": _HEADERS,
          "header_rows": {"Sheet": 1, "単価表": 1}}
BM_HR3 = {"sheets": ["Sheet", "単価表"], "headers": _HEADERS,
          "header_rows": {"Sheet": 3, "単価表": 1}}
# 旧テスト値相当（header_rows キー自体が無い book_meta）＝ hr0=0 の後方互換パス。
BM_LEGACY_NO_HEADER_ROWS = {"sheets": ["Sheet", "単価表"], "headers": _HEADERS}

# 名前 -> (op, resolved_args, book_meta, use_formula)
CASES: dict = {}


def _add(name, op, resolved, book_meta=BM_HR1, use_formula=True):
    assert name not in CASES, f"重複した case 名: {name}"
    CASES[name] = (op, resolved, book_meta, use_formula)


# --- SORT ---------------------------------------------------------------
_add("sort_asc_hr1", "SORT", {"col": "金額", "order": "asc"})
_add("sort_desc_hr1", "SORT", {"col": "金額", "order": "desc"})
_add("sort_asc_hr3", "SORT", {"col": "金額", "order": "asc"}, book_meta=BM_HR3)
_add("sort_use_formula_false_is_same_axis_noop", "SORT",
     {"col": "金額", "order": "asc"}, use_formula=False)  # SORT は use_formula を無視

# --- COMPUTE_COLUMN（2列演算・新規列） -----------------------------------
_add("compute_column_2op_newcol_formula_hr1", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*"})
_add("compute_column_2op_newcol_values_hr1", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*"}, use_formula=False)
_add("compute_column_2op_newcol_formula_hr3", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*"}, book_meta=BM_HR3)
_add("compute_column_2op_newcol_values_hr3", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*"}, book_meta=BM_HR3, use_formula=False)
# --- COMPUTE_COLUMN（2列演算・既存列 target あり） -----------------------
_add("compute_column_2op_target_formula_hr1", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*", "target": "金額"})
_add("compute_column_2op_target_values_hr1", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*", "target": "金額"}, use_formula=False)
# --- COMPUTE_COLUMN（単列×倍率・新規列） ---------------------------------
_add("compute_column_1op_newcol_formula_taxinc_hr1", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*", "factor": 1.1, "_new_col_label": "税込金額"})
_add("compute_column_1op_newcol_values_taxinc_hr1", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "*", "factor": 1.1, "_new_col_label": "税込金額"},
     use_formula=False)
_add("compute_column_1op_newcol_formula_no_label_hr1", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "/", "factor": 1.1})   # _new_col_label 無し → 数式風見出し
# --- COMPUTE_COLUMN（単列×倍率・既存列 target あり） ---------------------
_add("compute_column_1op_target_formula_hr1", "COMPUTE_COLUMN",
     {"operands": ["金額"], "operator": "/", "factor": 1.1, "target": "金額"})

# --- LOOKUP_FILL ----------------------------------------------------------
_add("lookup_fill_existing_target_col_hr1", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"})
_add("lookup_fill_new_target_col_hr1", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "備考", "source_sheet": "単価表", "key_col": "商品"})
_add("lookup_fill_existing_target_col_hr3", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"},
     book_meta=BM_HR3)
_add("lookup_fill_new_target_col_hr3", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "備考", "source_sheet": "単価表", "key_col": "商品"},
     book_meta=BM_HR3)

# --- AGGREGATE --------------------------------------------------------------
_add("aggregate_hr1", "AGGREGATE", {"group_col": "商品", "value_col": "金額"})
_add("aggregate_hr3", "AGGREGATE", {"group_col": "商品", "value_col": "金額"}, book_meta=BM_HR3)

# --- BOLD ---------------------------------------------------------------
_add("bold_row_hr1", "BOLD", {"target": "row:1"})
_add("bold_col_hr1", "BOLD", {"target": "col:商品"})
_add("bold_row_hr3", "BOLD", {"target": "row:1"}, book_meta=BM_HR3)
_add("bold_col_hr3", "BOLD", {"target": "col:商品"}, book_meta=BM_HR3)

# --- FILL_COLOR ------------------------------------------------------------
_add("fill_color_row_hr1", "FILL_COLOR", {"target": "row:1", "color": "yellow"})
_add("fill_color_col_hr1", "FILL_COLOR", {"target": "col:金額", "color": "red"})
_add("fill_color_row_hr3", "FILL_COLOR", {"target": "row:1", "color": "yellow"}, book_meta=BM_HR3)
_add("fill_color_col_hr3", "FILL_COLOR", {"target": "col:金額", "color": "red"}, book_meta=BM_HR3)

# --- NUMBER_FORMAT -----------------------------------------------------------
_add("number_format_hr1", "NUMBER_FORMAT", {"col": "金額", "style": "thousands"})
_add("number_format_hr3", "NUMBER_FORMAT", {"col": "金額", "style": "thousands"}, book_meta=BM_HR3)

# --- MERGE (header_row を使わない op) ----------------------------------------
_add("merge_a1_b1", "MERGE", {"range": "A1:B1"})
_add("merge_c2_d5", "MERGE", {"range": "C2:D5"})

# --- CHART --------------------------------------------------------------
_add("chart_hr1", "CHART", {"value_col": "金額"})
_add("chart_hr3", "CHART", {"value_col": "金額"}, book_meta=BM_HR3)

# --- CENTER_ALIGN ---------------------------------------------------------
_add("center_align_all_hr1", "CENTER_ALIGN", {"target": "all"})
_add("center_align_col_hr1", "CENTER_ALIGN", {"target": "col:商品"})
_add("center_align_col_hr3", "CENTER_ALIGN", {"target": "col:商品"}, book_meta=BM_HR3)

# --- APPEND_TOTAL ----------------------------------------------------------
_add("append_total_col0_no_label_factor1_hr1", "APPEND_TOTAL",
     {"col": "商品", "label": "合計", "factor": 1.0})   # col_idx=0 → ラベル省略・factor_tail無し
_add("append_total_with_label_factor1_hr1", "APPEND_TOTAL",
     {"col": "金額", "label": "合計", "factor": 1.0})
_add("append_total_with_label_factor1_1_hr1", "APPEND_TOTAL",
     {"col": "金額", "label": "税込合計", "factor": 1.1})
_add("append_total_with_label_factor1_1_hr3", "APPEND_TOTAL",
     {"col": "金額", "label": "税込合計", "factor": 1.1}, book_meta=BM_HR3)

# --- INSERT_ROWS -----------------------------------------------------------
_add("insert_rows_at3_count1", "INSERT_ROWS", {"at": 3, "count": 1})
_add("insert_rows_at1_count2", "INSERT_ROWS", {"at": 1, "count": 2})

# --- ★ 2026-08-26: 表の基本操作 3 種 --------------------------------------
_add("add_row_middle", "ADD_ROW",
     {"at": 3, "values": {"商品": "梨", "金額": 600}, "_headers": ["商品", "金額"]})
_add("add_row_string_only", "ADD_ROW",
     {"at": 2, "values": {"商品": "梨"}, "_headers": ["商品", "金額"]})
_add("delete_rows_at3", "DELETE_ROWS", {"at": 3, "count": 1})
_add("delete_rows_at2_count3", "DELETE_ROWS", {"at": 2, "count": 3})
_add("delete_column_second", "DELETE_COLUMN", {"col": "金額", "_headers": ["商品", "金額"]})

_add("set_cell_value_number", "SET_CELL_VALUE",
     {"row": "りんご", "col": "金額", "value": "2000", "_write_numeric": True,
      "_write_numeric_value": 2000.0, "_headers": ["商品", "金額"]})
_add("set_cell_value_text", "SET_CELL_VALUE",
     {"row": "りんご", "col": "商品", "value": "洋梨", "_headers": ["商品", "金額"]})

_add("set_where_gte", "SET_WHERE",
     {"col": "在庫", "cond_col": "金額", "cmp": "gte", "cond_value": 40000.0, "value": "◎",
      "_headers": ["商品", "金額", "在庫"], "_header_row": 1})
_add("set_where_contains", "SET_WHERE",
     {"col": "在庫", "cond_col": "商品", "cmp": "contains", "cond_value": "セット",
      "value": "×", "_headers": ["商品", "金額", "在庫"], "_header_row": 1})
_add("add_column_named", "ADD_COLUMN",
     {"name": "備考", "_at_col": 3, "_header_row": 1, "_headers": ["商品", "金額"]})
_add("add_column_unnamed", "ADD_COLUMN",
     {"name": "", "_at_col": 1, "_header_row": 1, "_headers": ["商品", "金額"]})

# --- SWAP（2026-08-27）── Basic には**名前**を渡す（番号を渡さない）------------
_add("swap_rows_by_name", "SWAP",
     {"a": "みかん", "b": "ぶどう", "_axis": "row", "_header_row": 1,
      "_a_pos": 3, "_b_pos": 4, "_headers": ["商品", "金額"]})
_add("swap_columns_by_name", "SWAP",
     {"a": "商品", "b": "金額", "_axis": "column", "_header_row": 1,
      "_a_pos": 1, "_b_pos": 2, "_headers": ["商品", "金額"]})

# --- DRAW_BORDERS / AUTOFIT（引数無し） ------------------------------------
_add("draw_borders", "DRAW_BORDERS", {})
_add("autofit", "AUTOFIT", {})

# --- PIVOT --------------------------------------------------------------
_add("pivot_hr1", "PIVOT", {"group_col": "商品", "value_col": "金額"})
_add("pivot_hr3", "PIVOT", {"group_col": "商品", "value_col": "金額"}, book_meta=BM_HR3)

# --- SET_COLUMN_VALUE --------------------------------------------------------
_add("set_column_value_hr1", "SET_COLUMN_VALUE", {"col": "商品", "value": "確認済み"})
_add("set_column_value_hr3", "SET_COLUMN_VALUE", {"col": "商品", "value": "確認済み"},
     book_meta=BM_HR3)

# --- EXTRACT --------------------------------------------------------------
# ★ _new_sheet は verify_dsl_args が積む値（_extract_output_sheet_name）だが、ここは
#   codegen_dsl を直接呼ぶ golden なので、素の resolved args として明示的に渡す。
_add("extract_numeric_gte_hr1", "EXTRACT",
     {"col": "金額", "cmp": "gte", "value": 40000, "_new_sheet": "金額40000以上"})
_add("extract_string_contains_hr1", "EXTRACT",
     {"col": "商品", "cmp": "contains", "value": "セット", "_new_sheet": "商品セットを含む"})
_add("extract_numeric_lt_hr3", "EXTRACT",
     {"col": "金額", "cmp": "lt", "value": 20, "_new_sheet": "金額20未満"}, book_meta=BM_HR3)

# --- DEDUP（EXTRACT の兄弟）--------------------------------------------------
# ★ _new_sheet は verify_dsl_args が積む値（_dedup_output_sheet_name）だが、EXTRACT と
#   同じ理由で素の resolved args として明示的に渡す。
_add("dedup_single_key_hr1", "DEDUP", {"keys": ["商品"], "_new_sheet": "商品の重複除去"})
_add("dedup_multi_key_hr1", "DEDUP",
     {"keys": ["商品", "単価"], "_new_sheet": "商品・単価の重複除去"})
_add("dedup_single_key_hr3", "DEDUP", {"keys": ["商品"], "_new_sheet": "商品の重複除去"},
     book_meta=BM_HR3)

# --- REPORT_PER_ROW（帳票段）--------------------------------------------------
# ★ _target_sheet/_report_rows は verify_dsl_args が積む値（unique_sheet_name で
#   一意名を決め切ってから渡す ── Basic 側で名前を作らない・設計文書の指示どおり）。
_add("report_per_row_two_rows_hr1", "REPORT_PER_ROW",
     {"template_sheet": "単価表", "name_col": "商品", "_target_sheet": "Sheet",
      "_report_rows": [{"row": 2, "sheet": "甲社"}, {"row": 3, "sheet": "乙社"}]})
_add("report_per_row_two_rows_hr3", "REPORT_PER_ROW",
     {"template_sheet": "単価表", "name_col": "商品", "_target_sheet": "Sheet",
      "_report_rows": [{"row": 4, "sheet": "甲社"}, {"row": 5, "sheet": "乙社"}]},
     book_meta=BM_HR3)

# --- FORMAT_MAP（様式写像段。REPORT_PER_ROW の兄弟・縦の展開）------------------
# ★ _target_sheet/_output_sheet/_header_tpl_row/_placeholder_tpl_row/_data_rows は
#   verify_dsl_args が積む値（unique_sheet_name で一意名を決め切ってから渡す）。
_add("format_map_two_rows_hr1", "FORMAT_MAP",
     {"template_sheet": "単価表", "_target_sheet": "Sheet", "_output_sheet": "単価表_出力",
      "_header_tpl_row": 1, "_placeholder_tpl_row": 2, "_data_rows": [2, 3]})
_add("format_map_two_rows_hr3", "FORMAT_MAP",
     {"template_sheet": "単価表", "_target_sheet": "Sheet", "_output_sheet": "単価表_出力",
      "_header_tpl_row": 1, "_placeholder_tpl_row": 2, "_data_rows": [4, 5]},
     book_meta=BM_HR3)

# --- SPLIT_CELL（1セルの複数値を右の列へ割る）----------------------------------
# ★ _parts/_new_cols は verify_dsl_args が**実データを数えて**積む値（LLM には数えさせない）。
_add("split_cell_newline_hr1", "SPLIT_CELL",
     {"col": "商品", "sep": chr(10), "_parts": 2, "_new_cols": ["商品_1", "商品_2"]})
_add("split_cell_comma_hr3", "SPLIT_CELL",
     {"col": "商品", "sep": ",", "_parts": 3,
      "_new_cols": ["商品_1", "商品_2", "商品_3"]}, book_meta=BM_HR3)


# --- 後方互換: header_rows キーが book_meta に無い旧テスト値 ------------------
_add("sort_legacy_no_header_rows_key", "SORT", {"col": "金額", "order": "asc"},
     book_meta=BM_LEGACY_NO_HEADER_ROWS)


def _case_ids():
    return sorted(CASES.keys())


def test_codegen_dsl_golden_coverage_matches_declared_ops():
    """★網羅性の自己検査: OP_WRITE_TARGET に載っている全 op が最低1ケース入っているか。"""
    covered_ops = {op for op, *_ in CASES.values()}
    all_ops = set(ailine.OP_WRITE_TARGET.keys())
    missing = all_ops - covered_ops
    assert not missing, f"codegen_dsl golden が網羅していない op: {sorted(missing)}"


@pytest.mark.parametrize("name", _case_ids())
def test_codegen_dsl_golden(name):
    op, resolved, book_meta, use_formula = CASES[name]
    actual = ailine.codegen_dsl(op, dict(resolved), book_meta, use_formula=use_formula)
    assert_golden_bytes(F1_DIR / f"{name}.bas", actual.encode("utf-8"), label=name)
