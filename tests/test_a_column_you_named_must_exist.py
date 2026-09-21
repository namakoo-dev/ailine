# -*- coding: utf-8 -*-
"""人が名乗った列が無いなら、別の列に書いて `✓` を出さない（2026-09-21・盲検 5 体目 ④）。

★★ 出所（買い手の言葉）:

> 「差額が1以上の行の**チェック列**に「◎」を付けて」と頼んだら…「チェック列」が無いのに、
> 勝手に **`差額`列**（さっき苦労して作った突き合わせ結果）を選んでいました。
> 止めてくれたので助かりましたが、**画面が最初に勧めてくるのが `--overwrite`** です。

  ★ 別の冊では黙って `確認` 列に書いて **`✓ 機械検証済み`** を出していた（実測で再現）。
    合格線 ①2「**頼んでいないものを変えない**」に当たる。

★★ なぜ見逃していたか（読んで確かめた）: 仕分けは 3 段階（①照合できた／②無言／③矛盾）で、
  ③ なら ✓ を出さないと**既に決まっている**。ところが反証の材料を拾う `task_designators` は
  **実在物との照合だけ**で拾う ── `チェック列` は実在しないので**見えず**、誰も拾わなかった
  語として残らないため ②（無言）に落ちていた。★「出ないことは信号でない」の形。

★★ 効かせる範囲を絞った理由（実装前に数えた・Namakoo 決裁）:
  「◯◯列」は 2 通りに使われる ──
      既存の列を名指す        「在庫列に「◎」を付けて」（在庫は実在する）
      これから作る列を説明する「売上から原価を**引いた列**を作って」（『引いた』列は無くて当然）
  検体に後者が 16 件あった。だから **既存列にだけ書く op**（`SET_WHERE` / `SET_COLUMN_VALUE`）
  の時だけ当てる ── どの op がそれかは `OP_WRITE_TARGET` の宣言から導く（手で並べない）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from ailine_core import subject  # noqa: E402

COLUMNS = ["請求番号", "商品名", "請求金額", "発注金額", "差額", "確認"]


def _tiers(op: str, resolved: dict, task: str) -> dict:
    meta = {"sheets": ["請求"], "headers": {"請求": COLUMNS}, "header_rows": {"請求": 1}}
    got = ailine.classify_subject_provenance(op, dict(resolved), meta, task, a=None)
    return {v.slot.key: v.tier for v in got}


def test_the_organ_sees_a_column_that_does_not_exist():
    """★ 器官そのもの ── 名乗られたのに実在しない列を拾うこと。"""
    assert subject.named_but_missing_columns(
        "差額が1以上の行のチェック列に「◎」を付けて", COLUMNS) == ("チェック列",)


def test_an_existing_column_called_by_its_name_is_not_evidence():
    """★★ 実在する列を『◯◯列』と呼んだだけなら、反証にしないこと。

    ★ ここが緩むと「確認列に◎を付けて」という**正しい依頼**まで止まる。
    """
    assert subject.named_but_missing_columns(
        "差額が1以上の行の確認列に「◎」を付けて", COLUMNS) == ()


@pytest.mark.parametrize("task", [
    "売上から原価を引いた列を作って",
    "数量と単価をかけた列を作って",
    "単価と数量を掛けた列を作って",
])
def test_a_described_new_column_is_not_treated_as_a_name(task):
    """★★ 「作る列の説明」を列名と読まないこと（誤爆の本命）。

    ★ 器官そのものは『引いた』を拾う ── 拾わせないのは**呼び出し側の絞り**（下の試験）。
      ここでは「作る側の op では材料を渡さない」ことで守られると宣言しておく。
    """
    got = subject.named_but_missing_columns(task, ["売上", "原価", "数量", "単価"])
    assert got, "★ この器官は拾う（だから作る側の op には渡してはいけない）"


def test_only_ops_that_write_into_an_existing_column_use_it():
    """★★ 効かせる範囲は**宣言から導く**こと（手で並べない）。

    ★ 既存列にだけ書く op ＝ `SET_WHERE` / `SET_COLUMN_VALUE`。
      「作る」側（計算列・転記・列追加）に当てると 16 件の検体が誤爆する。
    ★ 一覧を手で書くと、op を足した日に静かにずれる。
    """
    only = {op for op, t in ailine.OP_WRITE_TARGET.items()
            if t.writes == (ailine.WRITE_EXISTING_COLUMN,)}
    assert only == {"SET_WHERE", "SET_COLUMN_VALUE"}, only
    from _product_source import window_around
    seg = window_around("named_but_missing_columns(task or", before=700, after=200)
    assert "OP_WRITE_TARGET" in seg, "★ 絞りを宣言から導いていない"
    assert "WRITE_EXISTING_COLUMN,)" in seg, "★ 『既存列にだけ書く』で絞っていない"


def test_the_accident_no_longer_earns_a_tick():
    """★★ 事故そのもの ── 名乗った列が無いなら ③（✓ を出さない）へ落ちること。"""
    got = _tiers("SET_WHERE",
                 {"col": "確認", "value": "◎", "cond_col": "差額", "cmp": "gte", "cond": 1},
                 "差額が1以上の行のチェック列に「◎」を付けて")
    assert got.get("col") == subject.CONTRADICTED, got


def test_naming_the_real_column_still_passes():
    """★★ 陰性対照 ── 実在する列を名乗った回は今までどおり ①（✓ 満額）。

    ★ これが無いと「止めた」のか「何でも止める」のか分からない。
    """
    got = _tiers("SET_WHERE",
                 {"col": "確認", "value": "◎", "cond_col": "差額", "cmp": "gte", "cond": 1},
                 "差額が1以上の行の確認列に「◎」を付けて")
    assert got.get("col") == subject.MATCHED, got


def test_saying_nothing_about_the_column_is_still_unspoken():
    """★ 列について**何も言っていない**回は ②（無言）のまま ── 止めない。

    ★ ②（証拠が無い）と ③（反証がある）を同じ扱いにすると、まっとうな run が軒並み止まる
      ── 設計が実測つきで警告している所。
    """
    got = _tiers("SET_WHERE",
                 {"col": "確認", "value": "◎", "cond_col": "差額", "cmp": "gte", "cond": 1},
                 "差額が1以上の行に「◎」を付けて")
    assert got.get("col") == subject.UNSPOKEN, got
