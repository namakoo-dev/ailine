# 1 セル書換（2026-08-27）── Namakoo「梨の売上にピンポイントで入れたい」。
#
# ★ それまでは `SET_COLUMN_VALUE`（列を丸ごと同じ値にする）しか無く、
#   「梨の売上を2000にして」は「値を『』で囲め」と断られていた。
#
# ★★ architect の査読が名指しした穴（これが番人の要）:
#   既存の `check_set_column_value` は「対象列のデータ行が**全部**その値か」を見る。
#   1 セル用に流用すると **列全体を潰した方が pass する**（逆向きの検算）。
#   この機能で最も起きやすい壊れ方（列全体の codegen を流用して走査範囲を間違える）を、
#   番人が通してしまう。
#
# 契約:
#   ① 宣言したセルが宣言した値になっている
#   ② **変わったセルはちょうど 1 個**（列を潰していない）
#   ③ その 1 個の座標が宣言と一致する
#   ④ 行が見つからない・複数ある時は**決めない**（推測で別の行に書かない）
#   ⑤ 数字は数値のまま入る（文字列だと下流の SUM が静かに壊れる）
#   ⑥ 破壊の関所は**その 1 セルだけ**を見る（触らない行の値で止めない）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

ROWS = [["商品", "売上", "原価"], ["りんご", 1200, 700], ["みかん", 800, 300],
         ["梨", None, None], ["ぶどう", 1500, 900]]


def _book(tmp_path, rows=None, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "売上"
    for r in (rows or ROWS):
        ws.append(r)
    wb.save(p)
    return p


def _meta(p):
    return {"sheets": ["売上"], "headers": {"売上": ["商品", "売上", "原価"]},
             "header_rows": {"売上": 1}, "path": str(p)}


# --- ①②③ 事後条件（恒真殺しが本体）---------------------------------------------------

def test_one_cell_change_passes(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["りんご", 1200, 700],
                              ["みかん", 800, 300], ["梨", 2000, None],
                              ["ぶどう", 1500, 900]], name="after.xlsx")
    args = {"row": "梨", "col": "売上", "value": "2000",
             "_write_numeric": True, "_write_numeric_value": 2000.0}
    status, reason = ailine.check_set_cell_value(after, args, source_book=before)
    assert status == "pass", reason
    assert "1 個" in reason


def test_overwriting_the_whole_column_is_caught(tmp_path):
    """★★ architect が名指しした穴。列全体の番人ならここが **pass** してしまう。"""
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["りんご", 2000, 700],
                              ["みかん", 2000, 300], ["梨", 2000, None],
                              ["ぶどう", 2000, 900]], name="after.xlsx")
    args = {"row": "梨", "col": "売上", "value": "2000",
             "_write_numeric": True, "_write_numeric_value": 2000.0}
    status, reason = ailine.check_set_cell_value(after, args, source_book=before)
    assert status == "fail", f"列全体を潰したのに通した: {reason}"
    assert "列全体を潰した疑い" in reason, reason
    # ★ 対比: 既存の列全体の番人なら、この出力は pass する（逆向きの検算だった証拠）
    col_status, _ = ailine.check_set_column_value(
        after, {"col": "売上", "value": "2000", "_write_numeric": True,
                 "_write_numeric_value": 2000.0})
    assert col_status == "pass", "前提が崩れた（この対比が番人の存在理由）"


def test_writing_the_wrong_row_is_caught(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価"], ["りんご", 1200, 700],
                              ["みかん", 2000, 300], ["梨", None, None],
                              ["ぶどう", 1500, 900]], name="after.xlsx")
    args = {"row": "梨", "col": "売上", "value": "2000",
             "_write_numeric": True, "_write_numeric_value": 2000.0}
    status, reason = ailine.check_set_cell_value(after, args, source_book=before)
    assert status == "fail", reason


def test_no_source_book_means_warn_not_pass(tmp_path):
    """★ 変えていないセルを見ていないなら、断定しない。"""
    after = _book(tmp_path, [["商品", "売上", "原価"], ["梨", 2000, None]], name="a.xlsx")
    args = {"row": "梨", "col": "売上", "value": "2000",
             "_write_numeric": True, "_write_numeric_value": 2000.0}
    status, _ = ailine.check_set_cell_value(after, args, source_book=None)
    assert status == "warn"


# --- ④ 決められない時は決めない --------------------------------------------------------

@pytest.mark.parametrize("row,expect", [
    ("すいか", "見つかりません"),
    ("", "どの行か"),
])
def test_refuses_when_the_row_is_not_determined(tmp_path, row, expect):
    p = _book(tmp_path)
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", {"row": row, "col": "売上"}, _meta(p), task="2000にして")
    assert not ok and expect in err, err


def test_refuses_an_ambiguous_row(tmp_path):
    p = _book(tmp_path, [["商品", "売上", "原価"], ["梨", 1, 1], ["梨", 2, 2]], name="d.xlsx")
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", {"row": "梨", "col": "売上"}, _meta(p), task="2000にして")
    assert not ok and "2 行あります" in err, err


def test_refuses_an_unknown_column(tmp_path):
    p = _book(tmp_path)
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", {"row": "梨", "col": "利益"}, _meta(p), task="2000にして")
    assert not ok and "がこの表にありません" in err, err


# --- ⑤ 数字は数値のまま ----------------------------------------------------------------

def test_a_bare_number_is_taken_as_a_number(tmp_path):
    """★ 引用符を要求しない（1 セルなので）。★ 型は数値（文字列だと SUM が壊れる）。"""
    p = _book(tmp_path)
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", {"row": "梨", "col": "売上"}, _meta(p), task="梨の売上を2000にして")
    assert ok, err
    assert resolved["value"] == "2000"
    assert resolved.get("_write_numeric") is True, resolved
    assert resolved["_write_numeric_value"] == 2000.0


def test_a_quoted_string_is_taken_as_text(tmp_path):
    p = _book(tmp_path)
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", {"row": "梨", "col": "商品"}, _meta(p),
        task="梨の商品を「洋梨」にして")
    assert ok, err
    assert resolved["value"] == "洋梨"
    assert not resolved.get("_write_numeric")


# --- ⑥ 関所は 1 セルだけを見る ----------------------------------------------------------

def test_the_gate_looks_at_one_cell_only(tmp_path):
    """★ 実測で踏んだ: 空のセルに書くだけなのに「対象列に既存の値が 3 件」で止まった。

    宣言が「どの範囲を書くか」を持っていないのが根 ── 1 セル用は 1 セルだけ見る。
    """
    p = _book(tmp_path)
    meta = _meta(p)
    empty = ailine._maybe_warn_target_overwrite(
        "SET_CELL_VALUE", {"row": "梨", "col": "売上", "_target_sheet": "売上"}, meta, p)
    assert empty is None, f"空のセルに書くだけで止めた: {empty}"
    filled = ailine._maybe_warn_target_overwrite(
        "SET_CELL_VALUE", {"row": "みかん", "col": "売上", "_target_sheet": "売上"}, meta, p)
    assert filled and "みかん" in filled and "1 セルだけ" in filled, filled


# --- ⑤ 「空欄への一括書き込み」の助言が、1 セル書きで誤爆しないこと（2026-08-27）--------
#
# ★ 実測: README の手順（みかんの下に梨を追加 → 梨の売上を 2000 に）をそのまま通したら、
#   正しく 1 セルだけ書いたのに ★疑わしい が立ち、✓ が △ に落ちた。
#   梨の行は追加されたばかりで売上が空欄 ── 「空欄に同じ値を入れた」に当たってしまう。
# ★ 直しは op 名の if ではなく**宣言**（WRITE_SINGLE_CELL）。新しい op が増えても配線が要らない。
# ★ ただし緩めっぱなしにはしない: 宣言が 1 セルでも、実際に 2 セル以上変わったら鳴る。

def test_single_cell_write_does_not_raise_the_bulk_fill_advisory():
    before = {"cells": {("売上", 4, 2): (None, None)}}
    after = {"cells": {("売上", 4, 2): (2000, None)}}
    assert ailine.detect_uniform_fill(before, after) is not None, "前提: 既定では鳴る"
    assert ailine.detect_uniform_fill(before, after, single_cell=True) is None


def test_the_advisory_still_fires_when_more_than_one_cell_was_filled():
    """★ 恒真殺し: 宣言が 1 セルでも、実体が 2 セル以上なら鳴る（宣言と実体のずれ）。"""
    before = {"cells": {("売上", 3, 2): (None, None), ("売上", 4, 2): (None, None)}}
    after = {"cells": {("売上", 3, 2): (2000, None), ("売上", 4, 2): (2000, None)}}
    assert ailine.detect_uniform_fill(before, after, single_cell=True) is not None


def test_the_declaration_is_what_carries_it():
    """★ 配線の実在: SET_CELL_VALUE が WRITE_SINGLE_CELL を宣言していること。
       （助言の側だけ直して宣言を忘れると、本番では今までどおり鳴る＝片配線）"""
    assert ailine._op_writes("SET_CELL_VALUE", ailine.WRITE_SINGLE_CELL)
    assert not ailine._op_writes("SET_COLUMN_VALUE", ailine.WRITE_SINGLE_CELL)
