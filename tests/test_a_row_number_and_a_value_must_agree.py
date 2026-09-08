# -*- coding: utf-8 -*-
"""依頼が「N 行目」と値の両方を指して食い違ったら ✓ を出さない（2026-09-08）。

★★ 出所（曖昧な行の測定・実害のある false ✓）:

    表     2行目 ナット / 3行目 **ボルト** / 4行目 ナット / 6行目 ナット
    依頼   「**3行目のナット**を削除して」        ← 依頼そのものが矛盾している
    解釈   操作:行削除 削除位置:3 行数:1(推定)   ← ★『ナット』がどこにも出ていない
    実物   **ボルト**が消えた ／ 出力 **✓ 機械検証済み**

  ★ 機械は番号だけ取り、値を黙って捨てた。残差の関所が黙るのは『ナット』が
    **見出しでなく値**だから ── 「値だけが落ちた」家系が実害の形で出た最初の例。

★ 見るのは**依頼と実表**だけで、実行した操作は見ない ── 依頼が自己矛盾なら、
  何をしたとしても「頼まれたとおり」とは言えない。だから op を 1 つも列挙しない。
"""
from __future__ import annotations

import argparse

import openpyxl
import pytest

import ailine
from ailine_core.row_conflict import value_not_in_the_named_row

ROWS = [["品名", "棚", "単価"],
        ["ナット", "A-1", 4800],
        ["ボルト", "A-2", 800],
        ["ナット", "B-1", 1200],
        ["ワッシャ", "B-2", 300],
        ["ナット", "C-1", 900]]


def _book(path, rows=ROWS, sheet="部品表"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


# --- ① 判定（実表を読む純ロジック）------------------------------------------

@pytest.mark.parametrize("task, row_no, want", [
    # ★ 実測した事故そのもの（3 行目は ボルト）
    ("3行目のナットを削除して", 3, "ナット"),
    ("3行目のナットの単価を1500にして", 3, "ナット"),
    # ★ 対の試験: 一致している回は黙る
    ("2行目のナットを削除して", 2, None),
    ("3行目のボルトを削除して", 3, None),
    # ★ これから作る値は表に無いので黙る（「新品」は追加する値）
    ("3行目の下に新品を追加して", 3, None),
    # ★ 列名は値ではない（見出し行は見ない）
    ("3行目の単価を1500にして", 3, None),
    # ★ 行番号が無ければ黙る
    ("ナットを削除して", None, None),
])
def test_a_named_value_must_be_in_the_named_row(tmp_path, task, row_no, want):
    assert value_not_in_the_named_row(task, row_no, _book(tmp_path / "b.xlsx"),
                                      "部品表") == want


def test_a_row_that_does_not_exist_is_someone_elses_job(tmp_path):
    """★ その行が無いのは別の関所の受け持ち ── ここは黙る（二重に鳴らさない）。"""
    assert value_not_in_the_named_row("99行目のナットを削除して", 99,
                                      _book(tmp_path / "b2.xlsx"), "部品表") is None


def test_an_unreadable_book_stays_silent(tmp_path):
    """★ 測れない回は黙る（測れないものを鳴らさない）。"""
    assert value_not_in_the_named_row("3行目のナットを削除して", 3,
                                      tmp_path / "ない.xlsx", "部品表") is None


def test_the_longest_matching_value_is_taken(tmp_path):
    """★ 部分文字列で取りこぼさない（「青りんご」と「りんご」が両方在る表）。"""
    rows = [["品名", "個数"], ["りんご", 1], ["青りんご", 2], ["みかん", 3]]
    got = value_not_in_the_named_row("2行目の青りんごを削除して", 2,
                                     _book(tmp_path / "b3.xlsx", rows, "表"), "表")
    assert got == "青りんご", got


# --- ② 配線（✓ を出す唯一の関所を実際に通す）--------------------------------

def _run_finish_apply(tmp_path, task, capsys, name="in"):
    book = _book(tmp_path / f"{name}.xlsx")
    out = tmp_path / f"{name}.out.xlsx"
    out.write_bytes(book.read_bytes())
    work = tmp_path / f"w{name}"
    work.mkdir(exist_ok=True)
    a = argparse.Namespace(inplace=False, task=task, json=False, model="m",
                           keep_backups=3, copy=True, _target_sheet="部品表")
    ailine._finish_apply(a, book, out, work, {"op": "DELETE_ROWS"},
                         machine_verified=True,
                         scope="操作:行削除 削除位置:3 行数:1", scope_note="",
                         warning_count=0)
    return capsys.readouterr().out


def test_the_gate_is_wired_into_the_only_place_that_prints_the_check(tmp_path, capsys):
    shown = _run_finish_apply(tmp_path, "3行目のナットを削除して", capsys, name="bad")
    assert "その行に『ナット』はありません" in shown, shown
    assert "✓" not in shown, shown


def test_a_request_that_agrees_keeps_its_check(tmp_path, capsys):
    """★ 対の試験 ── 食い違っていない回は今までどおり ✓ が出る。"""
    clean = _run_finish_apply(tmp_path, "2行目のナットを削除して", capsys, name="ok")
    assert "はありません" not in clean, clean
    assert "✓" in clean, clean
