# -*- coding: utf-8 -*-
"""新しく作ったシートが**空**なら ✓ を出さない（2026-09-08・盲検 B）。

★★ 出所（実害のある false ✓）:

    依頼   「合計より上の行だけ抽出して新しいシートにコピーして」
    実行   操作:抽出 対象列:商品 条件:等しい 値:合計   ← 条件を誤変換
    事後条件 「4行中**0行**が一致 → 0行を抽出」        ← ★ 自分で数えている
    出力   **✓ 機械検証済み**／ 実物は見出しだけの空シート

  ★ 検出は在って**帰結が無い**（同日の「壊れた式を原本へ書く」と同じ形・2 度目）。
    抽出としては嘘をついていない（0 行を正しく抽出した）が、依頼は満たされていない。

★★ 「0 件なら落とす」ではない ── 4 op を実際に呼んで掃き出した:

    EXTRACT   0 行抽出  → pass  ★ 欠陥（成果物が空）
    DEDUP     重複 0 件 → pass  ← **正しい**（除くものが無いのは正常・表は残る）
    SET_WHERE 0 行一致  → fail  ← 既に守っている
    AGGREGATE 0 群      → fail  ← 既に守っている

  ★ 分かれ目は件数でなく**成果物**。だから見るのは「新しいシートに中身があるか」だけで、
    op 名は 1 つも列挙しない（新しい op を足しても自動で守られる）。
"""
from __future__ import annotations

import argparse

import openpyxl
import pytest

import ailine
from ailine_core.new_sheet import empty_new_sheets

ROWS = [["商品", "件数", "売上"], ["りんご", 3, 100], ["みかん", 5, 200]]


def _book(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    wb.save(path)
    return path


# --- ① 判定（純ロジック）-----------------------------------------------------

def test_a_new_sheet_with_only_a_header_is_named(tmp_path):
    before = _book(tmp_path / "b.xlsx", {"売上表": ROWS})
    after = _book(tmp_path / "a.xlsx",
                  {"売上表": ROWS, "商品が合計": [["商品", "件数", "売上"]]})
    assert empty_new_sheets(before, after) == ["商品が合計"]


def test_a_new_sheet_with_rows_is_left_alone(tmp_path):
    """★ 対の試験 ── 中身のある新シートには鳴らない（重複除去の正常系がこれ）。"""
    before = _book(tmp_path / "b2.xlsx", {"売上表": ROWS})
    after = _book(tmp_path / "a2.xlsx", {"売上表": ROWS, "重複除去": ROWS})
    assert empty_new_sheets(before, after) == []


def test_an_existing_empty_sheet_is_not_blamed(tmp_path):
    """★ もともと在った空シートは、この操作が作ったものではない。"""
    before = _book(tmp_path / "b3.xlsx", {"売上表": ROWS, "空": [["見出し"]]})
    after = _book(tmp_path / "a3.xlsx", {"売上表": ROWS, "空": [["見出し"]]})
    assert empty_new_sheets(before, after) == []


def test_an_unreadable_pair_stays_silent(tmp_path):
    """★ 測れない回は黙る（測れないものを鳴らさない）。"""
    assert empty_new_sheets(tmp_path / "ない.xlsx", tmp_path / "これも無い.xlsx") == []


# --- ② 掃き出しの結論を焼き込む（0 件が全部欠陥ではない）----------------------

def test_removing_nothing_is_still_a_correct_removal(tmp_path):
    """★ DEDUP は重複 0 件でも合格でよい ── 成果物（表）が残っているから。

    ★ この試験が在るのは、次に誰かが「0 件なら落とす」と一般化しないため。
    """
    from ailine_core.postconditions.move import check_dedup
    before = _book(tmp_path / "b4.xlsx", {"売上表": ROWS})
    after = _book(tmp_path / "a4.xlsx", {"売上表": ROWS, "重複除去": ROWS})
    st, _note = check_dedup(after, {"_new_sheet": "重複除去", "keys": ["商品"]}, 1, before)
    assert st == "pass"
    assert empty_new_sheets(before, after) == []


# --- ③ 配線（✓ を出す唯一の関所を実際に通す）--------------------------------

def _run_finish_apply(tmp_path, after_sheets, capsys, name="in"):
    book = _book(tmp_path / f"{name}.xlsx", {"売上表": ROWS})
    out = _book(tmp_path / f"{name}.out.xlsx", after_sheets)
    work = tmp_path / f"w{name}"
    work.mkdir(exist_ok=True)
    a = argparse.Namespace(inplace=False, task="合計より上の行だけ抽出して", json=False,
                           model="m", keep_backups=3, copy=True)
    ailine._finish_apply(a, book, out, work, {"op": "EXTRACT"}, machine_verified=True,
                         scope="操作:抽出 対象列:商品 条件:等しい 値:合計",
                         scope_note="", warning_count=0)
    return capsys.readouterr().out


def test_the_gate_is_wired_into_the_only_place_that_prints_the_check(tmp_path, capsys):
    shown = _run_finish_apply(
        tmp_path, {"売上表": ROWS, "商品が合計": [["商品", "件数", "売上"]]},
        capsys, name="empty")
    assert "中身がありません" in shown, shown
    assert "✓" not in shown, shown


def test_a_run_that_produced_rows_keeps_its_check(tmp_path, capsys):
    """★ 対の試験 ── 中身のある回は今までどおり ✓ が出る。"""
    clean = _run_finish_apply(
        tmp_path, {"売上表": ROWS, "抽出": [["商品", "件数", "売上"], ["りんご", 3, 100]]},
        capsys, name="ok")
    assert "中身がありません" not in clean, clean
    assert "✓" in clean, clean
