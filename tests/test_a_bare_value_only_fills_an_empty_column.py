# -*- coding: utf-8 -*-
"""値を「」で囲まなくても通す ── ただし**失うものが無い列だけ**（2026-09-07・第一段）。

★★ 出所（Namakoo）:「『』にしたのは暗黙の了解に近い。単語の説が qwen に読み取れるか
  怪しかったからだ。値なのか操作なのか、属性なのかを判別出来るなら『』はいらない」
  ★ そして「**『』を外すのは本当に用心してくれ**」。

★★ 測ってから決めた（→ docs/開発手法.md §6d）。未見の 20 本で:

      いまの製品のモデル（no thinking）  17/20   ★ 最良
      一回り大きいモデル（思考させる）   16/20   （6 倍遅い）
      手で書いた規則                     10/20   ★ 取るべき 10 本を 1 本も取れず

  外した中身は「操作を値と読む」「列名を値と読む」── **機械がタダで止められる形**。
  だから役割はこう分ける: **取り出しはモデル／拒否は機械**。

★★ 用心の形（賭け金で段を切る・モデルの自信では切らない）:
  一括書換は**列を丸ごと**書き換える。取り違えたときに失う量が最大の操作なので、
  緩めるのは**空の列だけ**。データの在る列は今までどおり「」を要求する。
  ★ 「」は 1 バイトも壊さない ── 要求から**逃げ道へ格下げ**しただけ。
"""
from __future__ import annotations

import openpyxl
import pytest

import ailine
from ailine_core import intent


def _book(tmp_path, memo_filled: bool):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名簿"
    ws.append(["氏名", "所属", "メモ", "締め日"])
    for i, (n, s) in enumerate([("田中", "営業"), ("鈴木", "経理")]):
        ws.append([n, s, ("済" if memo_filled else None), "2026/08/31"])
    p = tmp_path / ("filled.xlsx" if memo_filled else "empty.xlsx")
    wb.save(p)
    return p


def _run(book, task, col="メモ"):
    return ailine.verify_dsl_args(
        "SET_COLUMN_VALUE", {"col": col, "value": "確認済"},
        ailine.build_book_meta(book), task=task, vocab=ailine.load_vocab())


# --- 拒否の判定（純ロジック）------------------------------------------------

@pytest.mark.parametrize("value, rejected", [
    ("確認済", False),
    ("2026/09/30", False),
    ("所属", True),        # ★ 列名
    ("名簿", True),        # ★ シート名
    ("太字にする", True),  # ★ 操作
    ("空", True),          # ★ 消す操作（literal で『空』と書いてしまう）
    ("", True),
])
def test_only_a_real_value_gets_through(value, rejected):
    why = intent.why_not_a_value(
        value, ["氏名", "所属", "メモ", "締め日"], ["名簿"],
        ["太字", "並べ替え", "抽出", "集計"])
    assert (why is not None) is rejected, why


# --- 緩めた所 ---------------------------------------------------------------

def test_an_empty_column_accepts_a_bare_value(tmp_path):
    ok, res, _inf, err = _run(_book(tmp_path, memo_filled=False), "メモを全部確認済にして")
    assert ok, err
    assert res["value"] == "確認済", res


def test_the_source_tells_the_truth(tmp_path):
    """★ 囲んでいないのに「依頼文の引用」と言わないこと。

    ★ 実測（2026-09-07）: 最初の版はここが上書きされ、**利用者が囲んでいないのに
      『依頼文: 「確認済」』と出していた** ── 今日ずっと潰している形を自分で作った。
    """
    _ok, res, _inf, _err = _run(_book(tmp_path, memo_filled=False), "メモを全部確認済にして")
    said = (res.get("_sources") or {}).get("value") or ""
    assert "機械が取りました" in said and "空です" in said, said
    assert "依頼文: 「" not in said, said


# --- 緩めていない所（★ ここが用心の本体）------------------------------------

def test_a_column_with_data_still_requires_quotes(tmp_path):
    """★ 失うものが在る列は、今までどおり断る。"""
    ok, _res, _inf, err = _run(_book(tmp_path, memo_filled=True), "メモを全部確認済にして")
    assert not ok, "データの在る列に、囲まずに書き込んでいる"
    assert "囲んで" in err, err


def test_quoting_still_works_on_a_column_with_data(tmp_path):
    """★ 対で縛る ── 「」は 1 バイトも壊していない。"""
    ok, res, _inf, err = _run(_book(tmp_path, memo_filled=True), "メモを全部「確認済」にして")
    assert ok, err
    assert res["value"] == "確認済"
    assert "依頼文: 「" in ((res.get("_sources") or {}).get("value") or "")


def test_an_operation_word_is_never_taken_as_a_value(tmp_path):
    """★ モデルが操作名を値として返しても、機械が止めること（実測した誤りの形）。"""
    ok, _res, _inf, err = ailine.verify_dsl_args(
        "SET_COLUMN_VALUE", {"col": "メモ", "value": "太字"},
        ailine.build_book_meta(_book(tmp_path, memo_filled=False)),
        task="メモを太字にして", vocab=ailine.load_vocab())
    assert not ok, "操作の名前を値として書き込んでいる"


def test_a_column_name_is_never_taken_as_a_value(tmp_path):
    ok, _res, _inf, err = ailine.verify_dsl_args(
        "SET_COLUMN_VALUE", {"col": "メモ", "value": "所属"},
        ailine.build_book_meta(_book(tmp_path, memo_filled=False)),
        task="メモを所属にして", vocab=ailine.load_vocab())
    assert not ok, "列の名前を値として書き込んでいる"
