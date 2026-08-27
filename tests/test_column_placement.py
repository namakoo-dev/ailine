# 位置の言い回しを**全部の op へ**（横断層）── 2026-08-27。
#
# ★ Namakoo「これらの曖昧な入力『〜の右側に』『〜と〜の間に』などの操作は頻出だから
#   全ての操作でこれらを有効にする必要がある」。そのとおりで、位置は **op の性質でなく
#   依頼文の性質**。op ごとに if を書くと、op が増えるたびに配線が要る（今日 4 回踏んだ）。
#
# 契約:
#   ① 位置の解決は 1 箇所（resolve_new_column_placement）。宣言（WRITE_NEW_COLUMN）を
#      持つ op すべてに効く。op 名の if は 1 つも書かない
#   ② 動かし方も 1 箇所（wrap() が MoveColumnTo を足す）── codegen を op ごとに直さない
#   ③ 測っていないものは動かさない: 1 回で N 列作る op（cols_key）と、自分で位置を
#      決める op（col_index_key）は対象外
#   ④ 動かした回は**位置で比べる前提**（new_column）を使わない ── 右の列がずれるので
#      「同じ列を 2 回作った」と必ず誤報する
#   ⑤ 動かしたら、**どの op でも**根拠を解釈行に出す（見えない変更を作らない）

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

META = {"sheets": ["売上"], "headers": {"売上": ["商品", "売上", "原価"]},
         "header_rows": {"売上": 1}}


# --- ① 横断層そのもの -----------------------------------------------------------------

def test_the_layer_is_declaration_driven_not_op_named():
    """① 新しい列を作ると**宣言した** op なら、どれでも位置が効く。"""
    resolved = {"operands": ["売上", "原価"], "operator": "-", "_target_sheet": "売上"}
    got = ailine.resolve_new_column_placement(
        "COMPUTE_COLUMN", resolved, META, "商品の右に利益の列を作って", "売上")
    assert got and got["_move_new_col_to"] == 1 and got["_new_col_from"] == 3, got
    assert "『商品』（1列目）の右" in got["_at_basis"]


def test_an_op_that_makes_no_new_column_is_untouched():
    for op in ("SORT", "SET_CELL_VALUE", "DELETE_COLUMN", "SWAP"):
        got = ailine.resolve_new_column_placement(
            op, {"_target_sheet": "売上"}, META, "商品の右に列を作って", "売上")
        assert got is None, f"{op} に位置が効いてしまった: {got}"


def test_ops_that_place_themselves_or_make_many_are_excluded():
    """③ 測っていないものは動かさない ── ADD_COLUMN は自分で位置を決め、
       SPLIT_CELL は 1 回で N 列作る（複数本の移動はまだ測っていない）。"""
    for op in ("ADD_COLUMN", "SPLIT_CELL"):
        got = ailine.resolve_new_column_placement(
            op, {"_target_sheet": "売上"}, META, "商品の右に列を作って", "売上")
        assert got is None, f"{op} は対象外のはず: {got}"


def test_writing_into_an_existing_column_is_not_a_placement():
    """既存列への書き込みなら新しい列は生まれない ── 位置の話自体が起きない。"""
    got = ailine.resolve_new_column_placement(
        "COMPUTE_COLUMN", {"target": "原価", "_target_sheet": "売上"}, META,
        "商品の右に原価を計算して", "売上")
    assert got is None, got


def test_no_move_when_it_lands_there_anyway():
    """右端に出来るものを右端へ動かさない（無駄な移動は無駄な危険）。"""
    got = ailine.resolve_new_column_placement(
        "COMPUTE_COLUMN", {"_target_sheet": "売上"}, META,
        "原価の右に利益の列を作って", "売上")
    assert got is None, got


# --- ② 動かし方も 1 箇所 --------------------------------------------------------------

def test_codegen_appends_the_move_in_one_place():
    ok, r, _i, err = ailine.verify_dsl_args(
        "COMPUTE_COLUMN", {"operands": ["売上", "原価"], "operator": "-"}, META,
        task="商品の右に売上から原価を引いた列を作って")
    assert ok, err
    code = ailine.codegen_dsl("COMPUTE_COLUMN", r, META)
    assert "Call MoveColumnTo(oDoc, 3, 1)" in code, code


def test_no_op_specific_wiring_for_the_move():
    """② codegen に op 名で分けた移動のコードが無いこと（横断層 1 箇所で足す）。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    # ★ コメントは数えない（説明は何度書いてもよい）。**生成する Call は 1 箇所だけ**。
    calls = [l for l in src.splitlines()
              if "MoveColumnTo" in l and not l.lstrip().startswith("#")]
    assert len(calls) == 1, f"移動の配線が 2 箇所以上ある: {calls}"


def test_the_helper_exists():
    bas = (REPO / "src" / "ailine" / "helpers" / "AiLineHelpers.bas").read_text(encoding="utf-8")
    assert "Sub MoveColumnTo(" in bas
    # ★ 予約語との衝突で黙って死ぬ形を、もう一度作らない
    assert "Dim oR " not in bas and "Dim oR," not in bas


# --- ④ 位置がずれる回の前提 ------------------------------------------------------------

def test_position_based_precondition_is_skipped_when_this_run_moves_a_column():
    """④ 宣言でなく**その回の引数**から分かる事実で外す（2 度目の同じ形）。"""
    from ailine_core.write_precondition import check_write_preconditions_detail
    before = {"sheets": ["売上"], "cells": {"売上!1,3": ("原価", "General", None, False, None, None)}}
    after = {"sheets": ["売上"], "cells": {"売上!1,3": ("利益", "General", None, False, None, None),
                                            "売上!1,4": ("原価", "General", None, False, None, None)}}
    kw = dict(cell_ref=ailine._cell_ref, fmt_value=ailine._fmt_cell_value)
    assert check_write_preconditions_detail(("new_column",), before, after, **kw) is not None, \
        "前提: 動かさない回は（ずれを知らないので）鳴る"
    assert check_write_preconditions_detail(("new_column",), before, after,
                                             positions_shifted=True, **kw) is None


# --- ⑤ 見えない変更を作らない ----------------------------------------------------------

def test_the_basis_is_shown_for_any_op():
    """⑤ op ごとの表示登録に頼らない ── 足し忘れた op が黙って位置を動かす形を作らない。"""
    line = ailine.format_confirmation_line(
        "COMPUTE_COLUMN", {"operands": ["売上"], "operator": "-",
                            "_at_basis": "『商品』（1列目）の右＝2列目"}, set())
    assert "入れる位置:『商品』（1列目）の右＝2列目" in line, line
