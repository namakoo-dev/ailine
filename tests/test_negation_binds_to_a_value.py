# -*- coding: utf-8 -*-
"""否定は「文に在るか」でなく「**何に付いているか**」で読む ── その番人。

★★ 出所（2026-09-06・`bench/negation_reading_probe.py` で本物の CLI を --dry で 14 件）:
  否定の判定が文全体を見ていたため、**逆向きの事故が 2 つ**在った。どちらも実測で再現した。

    Q1「所属が営業**でない**行のメモに○を付けて」 → 比べ方=等しい（4/4 再現）
       ★ 否定が消え、**営業の行**に書いた。
    Q2「**メモ以外**は変えずに、所属が営業の行のメモに○を付けて」 → のどれでもない（3/3）
       ★ 「以外」は列の話なのに条件の否定と読まれ、**営業でない行**に書いた。

  ★ Q2 の方が重い ── 書く行が**丸ごと入れ替わる**。

★ この試験は 4 つを縛る:
    ① 判定（値に付く / 列名に付く / どちらでもない）
    ② **配線**（条件つき書換と抽出の**両方**で効くこと ── 片方だけ直すのが事故の形）
    ③ 決められない時は**断る**（eq に落として逆のことをしない）
    ④ 語の集合の**含有関係**（`_EXCEPT_WORDS` は `removal_reading` と共有なので触らない）
"""
from __future__ import annotations

import openpyxl
import pytest

import ailine
from ailine_core import negation


@pytest.fixture
def book(tmp_path):
    """★ **非対称**にする（営業 3 行 / 以外 1 行）── 対称だと eq と nin で当たる行数が
    同じになり、「宣言は直ったが書く行は直っていない」を見逃す。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名簿"
    ws.append(["氏名", "所属", "担当", "メモ"])
    for r in [("田中", "営業", "主任", ""), ("鈴木", "経理", "副任", ""),
              ("佐藤", "営業", "副任", ""), ("山田", "営業", "主任", "")]:
        ws.append(r)
    p = tmp_path / "meibo.xlsx"
    wb.save(p)
    return p


# --- ① 判定そのもの（純ロジック）-------------------------------------------

@pytest.mark.parametrize("task, want", [
    ("所属が営業以外の行のメモに「○」を付けて", negation.NEGATED),
    ("所属が営業でない行のメモに「○」を付けて", negation.NEGATED),
    ("所属が営業ではない行のメモに「○」を付けて", negation.NEGATED),
    ("所属が営業じゃない行のメモに「○」を付けて", negation.NEGATED),
    ("所属が営業を除く行のメモに「○」を付けて", negation.NEGATED),
    ("所属が営業を除いた行のメモに「○」を付けて", negation.NEGATED),
    # ★ 列名に付いた否定は、条件の否定ではない
    ("メモ以外は変えずに、所属が営業の行のメモに「○」を付けて", negation.PLAIN),
    ("担当を除いた列は触らずに、所属が営業の行のメモに「○」を付けて", negation.PLAIN),
    ("所属が営業の行のメモに「○」を付けて", negation.PLAIN),
    # ★ どちらにも付いていない ── 決めない
    ("所属がエイギョウ以外の行のメモに「○」を付けて", negation.UNCLEAR),
])
def test_negation_is_read_by_what_it_sticks_to(task, want):
    assert negation.reading(task, ["営業"], ["氏名", "所属", "担当", "メモ"]) == want


def test_the_shared_word_list_is_never_narrowed():
    """④ `_EXCEPT_WORDS` は `removal_reading`（「味噌汁の行を除いて」）と共有。

    ★ そこへ語を足すと別の判断が黙って変わるので、**上位集合をこちらが持つ**。
      片方だけ更新される（片配線）と、読み直しに届かない否定が生まれる。
    """
    assert set(ailine._EXCEPT_WORDS) <= set(negation.NEGATION_WORDS)
    assert set(negation.SHARED_WORDS) == set(ailine._EXCEPT_WORDS), (
        "共有語の写しがずれている（negation.SHARED_WORDS と ailine._EXCEPT_WORDS）")


# --- ②③ 配線: 条件つき書換と抽出の**両方**------------------------------------

def _set_where(book, task):
    return ailine.verify_dsl_args(
        "SET_WHERE", {"col": "メモ", "cond_col": "所属", "cmp": "eq", "value": "○"},
        ailine.build_book_meta(book), task=task, vocab=ailine.load_vocab())


def _extract(book, task):
    return ailine.verify_dsl_args(
        "EXTRACT", {"col": "所属", "cmp": "eq", "value": "営業"},
        ailine.build_book_meta(book), task=task, vocab=ailine.load_vocab())


@pytest.mark.parametrize("task, cmp_, rows", [
    ("所属が営業でない行のメモに「○」を付けて", "nin", [3]),
    ("所属が営業を除く行のメモに「○」を付けて", "nin", [3]),
    # ★ 列名に付いた「以外」で条件を反転させない（書く行が丸ごと入れ替わる事故）
    ("メモ以外は変えずに、所属が営業の行のメモに「○」を付けて", "eq", [2, 4, 5]),
])
def test_the_conditional_write_reads_the_negation_by_binding(book, task, cmp_, rows):
    ok, res, _inf, err = _set_where(book, task)
    assert ok, err
    assert res["cmp"] == cmp_, res
    assert res["_match_rows"] == rows, res["_match_rows"]


@pytest.mark.parametrize("task, cmp_", [
    ("所属が営業でない行を抜き出して", "nin"),
    ("メモ以外は変えずに、所属が営業の行を抜き出して", "eq"),
])
def test_the_extraction_reads_the_negation_by_binding(book, task, cmp_):
    """★ 片方だけ直すのが事故の形なので、**別の口**も同じ試験で縛る。"""
    ok, res, _inf, err = _extract(book, task)
    assert ok, err
    assert res["cmp"] == cmp_, res


@pytest.mark.parametrize("task, want", [
    ("所属が営業以外の行を抜き出して", ("所属", ["営業"])),
    # ★ 列名に付いた否定では、読み直し自身が「否定の抽出ではない」と返すこと
    ("メモ以外は変えずに、所属が営業の行を抜き出して", (None, None)),
    ("担当を除いた列は触らずに、所属が営業の行を抜き出して", (None, None)),
])
def test_the_reread_itself_refuses_a_negation_bound_to_a_column(book, task, want):
    """★ 3 つ目の口 ── 読み直しが「〜以外を抜き出す」と読む所。

    ★ ここを叩く検体が無かったため、この配線を殺しても試験は緑のままだった
      （2026-09-06 の変異試験で発覚）。**番人でなく試験が見ていなかった**方の抜け。
    """
    got = ailine.except_extraction_reading(ailine.build_book_meta(book), "名簿", task)
    assert got == want, got


def test_an_undecidable_negation_is_refused_not_guessed(book):
    """③ 決められないなら断る ── eq に落とすと、依頼と逆のことをして ✓ が出る。"""
    ok, _res, _inf, err = _set_where(book, "所属がエイギョウ以外の行のメモに「○」を付けて")
    assert not ok
    assert "否定" in err and "決められません" in err, err
