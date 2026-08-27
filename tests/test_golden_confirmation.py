"""C1-F5: format_confirmation_line(op, resolved_args, inferred) の確認行文字列を凍結する。

対象: 全17 op（_CONFIRM_FIELDS に載っている全部）。★推定タグ(推定)・出典タグ
（_sources 由来）・M2c のフィールド省略（キー自体が resolved_args に無い任意項目）も
それぞれ最低1ケース収載する。

ゴールデンは1ファイル（name -> 確認行）にまとめる。件数が少なく短い文字列なので、
diff の読みやすさを優先して1 JSON に集約する（tests/golden/f5_confirmation/golden.json）。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import ailine  # noqa: E402

from golden._harness import GOLDEN_ROOT, assert_golden_json  # noqa: E402

F5_FILE = GOLDEN_ROOT / "f5_confirmation" / "golden.json"

# 名前 -> (op, resolved_args, inferred)
CASES: dict = {}


def _add(name, op, resolved, inferred=()):
    assert name not in CASES, f"重複した case 名: {name}"
    CASES[name] = (op, resolved, set(inferred))


_add("sort", "SORT", {"col": "金額", "order": "desc"})
_add("compute_column_with_target_inferred", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*", "target": "金額"}, inferred=("target",))
_add("compute_column_without_target_omits_field", "COMPUTE_COLUMN",
     {"operands": ["数量", "単価"], "operator": "*"})
_add("lookup_fill_key_col_inferred", "LOOKUP_FILL",
     {"target_sheet": "Sheet", "target_col": "単価", "source_sheet": "単価表", "key_col": "商品"},
     inferred=("key_col",))
_add("aggregate", "AGGREGATE", {"group_col": "商品", "value_col": "金額"})
_add("bold_col_inferred", "BOLD", {"target": "col:商品"}, inferred=("target",))
_add("fill_color_row", "FILL_COLOR", {"target": "row:1", "color": "red"})
_add("number_format", "NUMBER_FORMAT", {"col": "金額", "style": "thousands"})
_add("merge", "MERGE", {"range": "A1:B1"})
_add("chart", "CHART", {"value_col": "金額"})
_add("center_align_all", "CENTER_ALIGN", {"target": "all"})
_add("append_total_with_source", "APPEND_TOTAL",
     {"col": "金額", "label": "合計", "factor": 1.1, "_sources": {"factor": "依頼文: 10%"}})
_add("append_total_without_source", "APPEND_TOTAL",
     {"col": "金額", "label": "合計", "factor": 1.0})
_add("insert_rows_count_inferred", "INSERT_ROWS", {"at": 3, "count": 1}, inferred=("count",))

_add("add_row_values", "ADD_ROW",
     {"at": 3, "values": {"商品": "梨"}, "_values_label": "商品=梨／金額=600"})
_add("delete_rows_inferred", "DELETE_ROWS", {"at": 3, "count": 1}, inferred=("count",))
_add("delete_column_one", "DELETE_COLUMN", {"col": "金額"})
_add("set_cell_value_one", "SET_CELL_VALUE",
     {"row": "りんご", "col": "金額", "value": "2000"})
_add("set_where", "SET_WHERE",
     {"col": "チェック", "value": "◎", "_cond_label": "『原価』が 500 以上",
      "_match_label": "2 行（2、4行目）"})
_add("add_column_named", "ADD_COLUMN",
     {"name": "備考", "_name_label": "備考", "_at_basis": "『金額』（2列目）の右＝3列目"})
_add("add_column_unnamed", "ADD_COLUMN",
     {"name": "", "_name_label": "（名前なし・見出しは空のまま）", "_at_basis": "末尾＝3列目"})
_add("swap_rows", "SWAP",
     {"a": "みかん", "b": "ぶどう", "_axis": "row",
      "_axis_label": "行（『みかん』の行＝3行目／『ぶどう』の行＝4行目）"})
_add("swap_columns", "SWAP",
     {"a": "売上", "b": "原価", "_axis": "column", "_axis_label": "列（見出しで一致）"})
_add("draw_borders_no_fields", "DRAW_BORDERS", {})
_add("autofit_no_fields", "AUTOFIT", {})
_add("pivot", "PIVOT", {"group_col": "商品", "value_col": "金額"})
_add("set_column_value_with_source", "SET_COLUMN_VALUE",
     {"col": "商品", "value": "確認済み", "_sources": {"value": "依頼文: 「確認済み」"}})
_add("extract", "EXTRACT", {"col": "金額", "cmp": "gte", "value": 40000.0})
_add("dedup", "DEDUP", {"keys": ["商品"]})
_add("dedup_multi_key", "DEDUP", {"keys": ["商品", "金額"]})
_add("report_per_row", "REPORT_PER_ROW", {"template_sheet": "雛形", "name_col": "取引先"})
_add("format_map", "FORMAT_MAP", {"template_sheet": "様式"})
_add("split_cell_newline", "SPLIT_CELL", {"col": "URL", "sep": chr(10),
                                           "_new_cols": ["URL_1", "URL_2"]})
_add("split_cell_comma", "SPLIT_CELL", {"col": "備考", "sep": ",",
                                         "_new_cols": ["備考_1", "備考_2"]})


def _case_ids():
    return sorted(CASES.keys())


def test_format_confirmation_line_golden_coverage_matches_declared_ops():
    covered_ops = {op for op, *_ in CASES.values()}
    all_ops = set(ailine.OP_WRITE_TARGET.keys())
    missing = all_ops - covered_ops
    assert not missing, f"format_confirmation_line golden が網羅していない op: {sorted(missing)}"


def test_format_confirmation_line_golden():
    result = {}
    for name in _case_ids():
        op, resolved, inferred = CASES[name]
        result[name] = ailine.format_confirmation_line(op, dict(resolved), set(inferred))
    assert_golden_json(F5_FILE, result, label="f5_confirmation")


@pytest.mark.parametrize("name", _case_ids())
def test_format_confirmation_line_each_case_is_nonempty(name):
    # golden 本体は1ファイル集約なので、個々の case が壊れていないかだけ最低限見る。
    op, resolved, inferred = CASES[name]
    line = ailine.format_confirmation_line(op, dict(resolved), set(inferred))
    assert line.startswith("解釈: 操作:")
