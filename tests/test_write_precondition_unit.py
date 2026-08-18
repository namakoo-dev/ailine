"""単位F: ailine_core/write_precondition.py の単体（純ロジック）と宣言の番人。

★ 番人が2つある（この repo で唯一再発していない形＝宣言表＋網羅の検査、に倣う）:
  1. 書き込み領域の種類（ailine.WRITE_KINDS）が、必ず「前提あり(PRECONDITIONS)」か
     「前提なし(NO_PRECONDITION)」のどちらかに宣言されていること。種類を1つ足したのに
     前提を決め忘れる＝黙って素通りする新しい領域、を機械で防ぐ。
  2. 1つの op の writes が「前提あり」と「前提なし」を混ぜていないこと。混ざると
     「片方の前提が破れたが、もう片方の宣言では正常」という判定不能が生まれる
     （今日そういう op は無い ── 足す人が最初にここでぶつかる）。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine  # noqa: E402
from ailine_core.write_precondition import (  # noqa: E402
    NO_PRECONDITION, PRECONDITIONS, check_write_preconditions,
)


def _snap(cells: dict, sheets=("Sheet",)) -> dict:
    """snapshot() の必要な部分だけを持つ最小の dict。cells は {"シート!行,列": 値}。"""
    return {"sheets": list(sheets),
            "cells": {k: (v, "General", None, False, None, None) for k, v in cells.items()}}


def _check(writes, before, after):
    return check_write_preconditions(writes, before, after,
                                      cell_ref=ailine._cell_ref, fmt_value=ailine._fmt_cell_value)


# --- 宣言の番人 ---------------------------------------------------------------

def test_every_write_kind_declares_whether_it_has_a_precondition():
    declared = set(PRECONDITIONS) | set(NO_PRECONDITION)
    assert declared == set(ailine.WRITE_KINDS), (
        f"前提の宣言が無い/余分な書き込み領域の種類: {declared ^ set(ailine.WRITE_KINDS)}")


def test_no_op_mixes_kinds_with_and_without_a_precondition():
    for op, wt in ailine.OP_WRITE_TARGET.items():
        with_pre = [k for k in wt.writes if k in PRECONDITIONS]
        without = [k for k in wt.writes if k in NO_PRECONDITION]
        assert not (with_pre and without), (
            f"{op}: 前提のある領域 {with_pre} と 前提のない領域 {without} を同時に宣言している"
            f"（どちらの前提で判定すべきか決まらない）")


# --- new_row_at_end -----------------------------------------------------------

def test_new_row_at_end_silent_when_the_written_row_was_empty():
    before = _snap({"Sheet!1,1": "品名", "Sheet!1,2": "金額", "Sheet!2,1": "a", "Sheet!2,2": 100})
    after = _snap({**{"Sheet!1,1": "品名", "Sheet!1,2": "金額", "Sheet!2,1": "a", "Sheet!2,2": 100},
                   "Sheet!3,1": "合計", "Sheet!3,2": 100})
    assert _check(("new_row_at_end",), before, after) is None


def test_new_row_at_end_fires_when_an_occupied_row_was_overwritten():
    before = _snap({"Sheet!3,3": "合計", "Sheet!3,4": 116600})
    after = _snap({"Sheet!3,3": "合計", "Sheet!3,4": "=SUM(D2:INDEX(D:D,ROW()-1))"})
    msg = _check(("new_row_at_end",), before, after)
    assert msg == ("★ 末尾に新しい行を足すはずが、既存の行の値を 1 件書き換えました"
                    "（Sheet!D3: 116600 → '=SUM(D2:INDEX(D:D,ROW()-1))'）")


def test_new_row_at_end_counts_every_hit_and_shows_at_most_three():
    before = _snap({f"Sheet!2,{c}": c * 10 for c in range(1, 6)})
    after = _snap({f"Sheet!2,{c}": c * 99 for c in range(1, 6)})
    msg = _check(("new_row_at_end",), before, after)
    assert "既存の行の値を 5 件書き換えました" in msg
    assert "ほか2件" in msg


# --- new_sheet ----------------------------------------------------------------

def test_new_sheet_silent_when_only_a_brand_new_sheet_was_written():
    before = _snap({"工事台帳!1,1": "取引先"}, sheets=("工事台帳",))
    after = _snap({"工事台帳!1,1": "取引先", "集計!1,1": "取引先"}, sheets=("工事台帳", "集計"))
    assert _check(("new_sheet",), before, after) is None


def test_new_sheet_fires_when_an_existing_sheet_was_replaced():
    before = _snap({"集計!1,1": "年度", "集計!2,1": 2025}, sheets=("工事台帳", "集計"))
    after = _snap({"集計!1,1": "取引先", "集計!2,1": "a"}, sheets=("工事台帳", "集計"))
    msg = _check(("new_sheet",), before, after)
    assert msg == ("★ 新しいシートを作るはずが、既存のシート『集計』の値を 2 件書き換えました"
                    "（集計!A1: '年度' → '取引先'、集計!A2: 2025 → 'a'）")


# --- format_only ---------------------------------------------------------------

def test_format_only_silent_when_no_value_moved():
    snap = _snap({"Sheet!1,1": "品名", "Sheet!2,1": "a"})
    assert _check(("format_only",), snap, snap) is None


def test_format_only_fires_on_any_value_change():
    before = _snap({"Sheet!1,2": "金額"})
    after = _snap({})
    assert _check(("format_only",), before, after) == (
        "★ 書式だけのはずが、セルの値が 1 件変わりました（Sheet!B1: '金額' → (空)）")


# --- row_shift / reorder --------------------------------------------------------

def test_reorder_silent_when_values_are_only_permuted():
    before = _snap({"Sheet!2,1": "a", "Sheet!2,2": 200, "Sheet!3,1": "b", "Sheet!3,2": 300})
    after = _snap({"Sheet!2,1": "b", "Sheet!2,2": 300, "Sheet!3,1": "a", "Sheet!3,2": 200})
    assert _check(("reorder",), before, after) is None


def test_row_shift_silent_when_rows_move_down():
    before = _snap({"Sheet!2,1": "a", "Sheet!3,1": "b"})
    after = _snap({"Sheet!4,1": "a", "Sheet!5,1": "b"})
    assert _check(("row_shift",), before, after) is None


def test_reorder_ignores_formulas_because_references_follow_the_rows():
    # 行が動けば =D2-C2 は =D3-C3 に書き換わる ―― 文字列として比べると必ず食い違うが、
    # それは移動の正常な副作用であって破壊ではない。
    before = _snap({"Sheet!2,4": "=D2-C2", "Sheet!3,4": "=D3-C3"})
    after = _snap({"Sheet!2,4": "=D5-C5", "Sheet!3,4": "=D6-C6"})
    assert _check(("reorder",), before, after) is None


def test_reorder_fires_when_a_value_disappears():
    before = _snap({"Sheet!2,1": "a", "Sheet!3,1": "b"})
    after = _snap({"Sheet!2,1": "a"})
    assert _check(("reorder",), before, after) == "★ 行を動かすだけのはずが、値が 1 件消えました（'b'）"


def test_reorder_silent_when_values_are_only_added():
    # 増えた分は「破壊」ではない（幽霊データ検出の領分）。関所は消えた側だけを見る。
    before = _snap({"Sheet!2,1": "a"})
    after = _snap({"Sheet!2,1": "a", "Sheet!3,1": "新規"})
    assert _check(("reorder",), before, after) is None


# --- 前提を持たない領域 ----------------------------------------------------------

@pytest.mark.parametrize("kind", sorted(NO_PRECONDITION))
def test_kinds_without_a_precondition_never_fire(kind):
    before = _snap({"Sheet!2,1": "a"})
    after = _snap({"Sheet!2,1": "b"})
    assert _check((kind,), before, after) is None


def test_unknown_or_empty_declaration_is_silent():
    snap = _snap({"Sheet!2,1": "a"})
    assert _check((), snap, snap) is None
    assert _check(None, snap, snap) is None


# --- ailine.py 側の薄い配線 -------------------------------------------------------

def test_maybe_warn_write_precondition_reads_the_declaration():
    before = _snap({"Sheet!3,4": 116600})
    after = _snap({"Sheet!3,4": "=SUM(D2:INDEX(D:D,ROW()-1))"})
    # APPEND_TOTAL は writes=new_row_at_end → 既存行の上書きで鳴る
    assert ailine._maybe_warn_write_precondition("APPEND_TOTAL", before, after) is not None
    # SET_COLUMN_VALUE は writes=existing_column（前提なし＝既存の関所の担当）→ 鳴らない
    assert ailine._maybe_warn_write_precondition("SET_COLUMN_VALUE", before, after) is None
    # 未知の op（FREEFORM 等・宣言そのものが無い）は判定材料が無いので黙る
    assert ailine._maybe_warn_write_precondition("FREEFORM", before, after) is None
