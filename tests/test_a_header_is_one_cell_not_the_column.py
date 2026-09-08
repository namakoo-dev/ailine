# -*- coding: utf-8 -*-
"""「<列名>の見出しに」は列ぜんぶでなく**その 1 セル**（2026-09-08・盲検 B）。

★★ 出所（false ✓）:

    依頼   「金額の**見出し**に色を付けて太字にして」
    宣言   操作:背景色 対象:col:金額 ／ 操作:太字 対象:col:金額
    実物   E1〜E8（データ行も合計式の行も）が黄色＋太字
    出力   **✓ 機械検証済み**

  ★ 依頼の「見出し」が宣言のどこにも出ていない ── 三項の「依頼」が落ちた形（#4 と同型）。

★★ 08-27 の判断は**反転しない**（この repo の既存の決裁）:
  「見出しを太字にして」（列を言わない）は**見出し行ぜんぶ**の意味でもありうるので
  曖昧 ── 勝手に狭めない。★ 分かれ目は**列が名指しされているか**で、
  列の見出しは 1 つしかないので曖昧でない。

★ 置き場は「読み直しの門」ではなく**引数の解決**（col: を実在の列に解いた直後）:
  門は len(plan)==1 の時しか動かず、この依頼は 2 段（背景色＋太字）なので届かなかった。
  解決に置けば**複合計画の各段**に自動で効く。
"""
from __future__ import annotations

import pytest

import ailine
from ailine_core.header_cell import header_cell_target, names_a_column_header

COLS = ["取引先", "項目", "件数", "単価", "金額"]


# --- ① 判定（純ロジック）-----------------------------------------------------

@pytest.mark.parametrize("task, want", [
    # ★ 実測した事故そのもの
    ("金額の見出しに色を付けて太字にして", "cell:1,5"),
    ("金額のヘッダーだけ塗って", "cell:1,5"),
    ("金額のヘッダを太字に", "cell:1,5"),
    # ★ 対の試験: 列ぜんぶを頼んだ回は狭めない（狭めるのも事故）
    ("金額の列を太字にして", None),
    ("金額を太字にして", None),
    # ★ 08-27 の判断: 列を言わない「見出し」は曖昧なので狭めない
    ("見出しを太字にして", None),
    # ★ 見出し語が文中に在るだけでは狭めない（別の列の話かもしれない）
    ("見出しを太字にして、金額の列に色を付けて", None),
])
def test_only_a_named_column_header_is_narrowed(task, want):
    assert header_cell_target(task, "col:金額", 1, COLS) == want


def test_a_column_that_is_not_in_the_table_is_not_narrowed():
    """★ 実在の列名だけを見る（依頼文から名前を切り出さない・A' 原則）。"""
    assert header_cell_target("部門の見出しを太字に", "col:部門", 1, COLS) is None


def test_the_header_row_is_taken_from_the_book_not_assumed():
    """★ 見出し行は決め打ちしない（2 行目が見出しのブックが実在する）。"""
    assert header_cell_target("金額の見出しを太字に", "col:金額", 3, COLS) == "cell:3,5"


def test_the_word_must_sit_next_to_the_column_name():
    assert names_a_column_header("金額の見出し", "金額") is True
    assert names_a_column_header("見出しと金額", "金額") is False


# --- ② 配線（引数の解決を実際に通す）----------------------------------------

def _resolve(task, op="FILL_COLOR", target="col:金額", header_row=1):
    resolved = {"target": target}
    if op == "FILL_COLOR":
        resolved["color"] = "yellow"
    got = ailine._verify_bold(resolved, set(), "請求", {"請求": COLS}, op,
                              task, header_row)
    return got, resolved


@pytest.mark.parametrize("op", ["BOLD", "FILL_COLOR", "CENTER_ALIGN"])
def test_every_format_op_narrows_the_same_way(op):
    """★ 3 つの書式 op が同じ 1 箇所を通ることを縛る（片配線を作らない）。"""
    got, resolved = _resolve("金額の見出しを整えて", op=op)
    assert got is None, got
    assert resolved["target"] == "cell:1,5", resolved
    # ★ 宣言から列名を消さない ── `cell:1,5` だけでは人にどの列か伝わらず、
    #   残差の関所も「依頼の『金額』が解釈に出ていない」と正しく鳴いてしまう。
    assert resolved.get("_header_col") == "金額", resolved


def test_a_whole_column_request_is_left_alone():
    got, resolved = _resolve("金額の列を太字にして", op="BOLD")
    assert got is None, got
    assert resolved["target"] == "col:金額", resolved
    assert "_header_col" not in resolved, resolved


def test_the_declaration_shows_which_column_was_narrowed():
    """★ 狭めた事実と、狭めた先の列名を**両方**見せる（残差の関所と噛み合わせる）。"""
    for op in ("BOLD", "FILL_COLOR", "CENTER_ALIGN"):
        keys = [k for _label, k, _t in ailine._CONFIRM_FIELDS[op]]
        assert "_header_col" in keys, (op, keys)
