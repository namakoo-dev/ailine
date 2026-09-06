# -*- coding: utf-8 -*-
"""依頼に在るのに解釈に出ない列名が在れば ✓ を名乗らない ── その番人。

★★ なぜ在るか（2026-09-06・実測で決めた・Namakoo 決裁 B）:
  事後条件が見るのは「宣言 vs 実体」だけで、**依頼から落ちた分は原理的に見えない**。
  本物の走行 3179 件を調べると、条件つき書換を *行追加* と誤読した 28 件は
  **post=pass** と言っていた。同じ依頼の前後で自然実験になっていた:

      09-04 06:26〜09-05 03:34  op=ADD_ROW    解釈に『所属』が出ない
      09-05 06:01 以降          op=SET_WHERE  解釈に『所属』が出る

  鳴った 28 件はすべて本物の欠陥（誤爆 0）。再現: `bench/residue_gate_on_history.py`。

★ この試験は 3 つを一度に縛る:
    ① 判定そのもの（本物の欠陥で鳴り、直った版で黙る）
    ② **配線**（✓ を出す唯一の関数を実際に通して ⚠ が出る）
    ③ **1 箇所しか無いこと**（呼び出し側 4 箇所に書き写されていない ── 片配線を作らない）
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

import openpyxl
import pytest

import ailine
from ailine_core import residue

HEADERS = {"氏名", "所属", "内線", "メモ"}
TASK = "所属が営業の行のメモに「○」を付けて"

#: 実際に履歴に残っていた 2 つの解釈行（前＝誤読 / 後＝正しい）。★ 手で創作しない
DECL_BROKEN = "操作:行追加 挿入位置:2 位置の根拠:『営業』の行＝2行目 入れる値:メモ=○"
DECL_FIXED = ("操作:条件つき書換 書き込む列:メモ 書き込む値:○（依頼文: 「○」） "
              "条件を見る列:所属 比べ方:等しい 条件:『所属』が 営業 等しい 当てはまる行:1 行（2行目）")


def test_it_fires_on_the_real_defect_and_is_silent_on_the_fix():
    """① 判定 ── 同じ依頼・同じ機械の、直る前と後。"""
    broken = residue.unaccounted_request_words(
        TASK, DECL_BROKEN, ailine._op_match_pool("ADD_ROW"), HEADERS)
    fixed = residue.unaccounted_request_words(
        TASK, DECL_FIXED, ailine._op_match_pool("SET_WHERE"), HEADERS)
    assert broken == ["所属"], broken
    assert fixed == [], fixed


def test_the_shapes_it_cannot_see_are_stated_not_hidden():
    """★ 捕まえないと**測って決めた**形。ここが赤くなったら「見えるようになった」の意味。

    ・落ちたのが**値だけ**（列名は宣言に在る）
    ・**否定の反転**（「営業以外」を「営業」と読む ── 語が 1 つも落ちない）
      ★ これは 09-05 06:06 の実走行に在った本物の誤りで、実際に黙っていた。
    """
    only_value = residue.unaccounted_request_words(
        "取引先が東西商事の行を消して", "操作:行削除 削除位置:3 条件を見る列:取引先",
        ailine._op_match_pool("DELETE_ROWS"), {"取引先", "品名"})
    inverted = residue.unaccounted_request_words(
        "所属が営業以外の行のメモに「○」を付けて", DECL_FIXED,
        ailine._op_match_pool("SET_WHERE"), HEADERS)
    assert only_value == [], only_value
    assert inverted == [], inverted


def _run_finish_apply(tmp_path, task: str, scope: str, op: str, capsys):
    """✓ を出す唯一の関数を**本当に通す**（--copy 経路なので原本置換はしない）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名簿"
    ws.append(["氏名", "所属", "内線", "メモ"])
    ws.append(["田中", "営業", "100", ""])
    book = tmp_path / "in.xlsx"
    wb.save(book)
    out = tmp_path / "in.out.xlsx"
    out.write_bytes(book.read_bytes())
    work = tmp_path / "w"
    work.mkdir(exist_ok=True)
    a = argparse.Namespace(inplace=False, task=task, json=False, model="m",
                           keep_backups=3, copy=True)
    result = {"op": op}
    ailine._finish_apply(a, book, out, work, result, machine_verified=True,
                         scope=scope, scope_note="", warning_count=0)
    return capsys.readouterr().out


def test_the_gate_is_actually_wired_into_the_only_place_that_prints_the_check(tmp_path, capsys):
    """② 配線 ── 判定が正しくても、呼ばれていなければ何も守らない（在っても鳴らない）。"""
    shown = _run_finish_apply(tmp_path, TASK, DECL_BROKEN, "ADD_ROW", capsys)
    assert "『所属』" in shown and "⚠" in shown, shown
    assert "✓" not in shown, shown          # ★ B: ✓ は名乗らない

    clean = _run_finish_apply(tmp_path, TASK, DECL_FIXED, "SET_WHERE", capsys)
    assert "『所属』" not in clean, clean    # ★ 正しい回は黙る


def test_the_judgment_lives_in_exactly_one_place():
    """③ 片配線を作らない ── 呼び出し側 4 箇所に書き写されていないこと。

    ★ 走査の未達（2026-09-05）と同じ畳み方: 判断は `_finish_apply` の中に 1 つ、
      呼び出し側は材料を渡すだけ。★ ここが 2 以上になったら、**片方だけ直す**事故が
      起きる形に戻っている（同じ形の再来を 08-21〜08-24 に 3 波観測している）。
    """
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    assert src.count("unaccounted_request_words") == 1, "判定が 2 箇所以上に散っている"
    # ★ 呼び出し側（_finish_apply(...) を呼ぶ行）の周辺に判定が無いこと
    for m in re.finditer(r"_finish_apply\(a, book", src):
        near = src[max(0, m.start() - 1200):m.start()]
        assert "unaccounted_request_words" not in near, "呼び出し側が判定を持っている"
