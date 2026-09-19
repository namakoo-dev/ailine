# -*- coding: utf-8 -*-
"""`--op` で固定した回の断りは、その op が**本当にすること**を言う（2026-09-19）。

★★ 出所（導通試験の最初の本物）: 案 D が読みの割れを止めて候補を示した後、
  **示した道を歩いたら別の断りに当たった** ── しかも文言が事実と違った:

      --op ADD_ROW で歩く
        → ？ 依頼文が『北斗精機』の行を指していますが、『行追加』は
             その列のデータ行を**全部**書き換えます

  `ADD_ROW` は行を足す op で、**列を全部書き換えたりしない**。
  ★ Namakoo の合格線「断りの文言が正確であること」に直撃する形。
  ★ しかも「通らない道を示した」ことになる（導通試験の verdict で言えば path_fails）。

★★ 根: 絞りに `plan_writes_beyond_one_cell` を使っていた。その集合には
  **行や列を足すだけの op** も入る ── 実測で 9 op 中 5 つ（行追加・行挿入・列追加・
  合計追加・セル分割）が、既存列を 1 つも書き換えないのにこの断りを浴びていた。
★ 直しは宣言から絞る ── `WRITE_EXISTING_COLUMN` と自分で言っている op だけ。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))
import ailine  # noqa: E402


@pytest.mark.parametrize("op", ["SET_COLUMN_VALUE", "SET_WHERE", "COMPUTE_COLUMN", "LOOKUP_FILL"])
def test_the_column_overwrite_warning_is_true_for_these(op):
    """★ この文言が当てはまる op ── 既存の列を書き換えると自分で宣言している。"""
    assert ailine._op_writes(op, ailine.WRITE_EXISTING_COLUMN), op


@pytest.mark.parametrize("op", ["ADD_ROW", "INSERT_ROWS", "ADD_COLUMN", "APPEND_TOTAL", "SPLIT_CELL"])
def test_the_column_overwrite_warning_is_a_lie_for_these(op):
    """★★ 事故そのもの: 行や列を**足すだけ**の op は、既存列を 1 つも書き換えない。

    ★ ここが真になると「その列のデータ行を全部書き換えます」が嘘になる。
    ★ 旧い絞り（plan_writes_beyond_one_cell）ではこの 5 つが全部引っ掛かっていた。
    """
    assert not ailine._op_writes(op, ailine.WRITE_EXISTING_COLUMN), op
    assert ailine.plan_writes_beyond_one_cell([{"op": op}]), (
        f"★ 旧い絞りがこの op を含まないなら、この試験の前提が変わっている: {op}")


def test_the_gate_is_narrowed_by_the_declaration_not_by_the_wide_gate():
    """★★ 配線: 絞りが宣言（WRITE_EXISTING_COLUMN）で、広い門に戻っていないこと。

    ★ 場所で決め打ちしない ── 製品の出所を配線経由で読む。
    """
    from _product_source import window_around
    body = window_around("1 か所だけ直すなら『1セル書換』を選んでください", before=900, after=200)
    assert body, "★ 探す場所が空（この検査が空回りする）"
    assert "_op_writes(forced_op, WRITE_EXISTING_COLUMN)" in body, (
        "★ 絞りが宣言から来ていない（広い門に戻っている疑い）")
    assert "plan_writes_beyond_one_cell([{\"op\": forced_op}])" not in body, (
        "★ 旧い広い門が残っている ── 行追加にまで嘘の文言が出る")
