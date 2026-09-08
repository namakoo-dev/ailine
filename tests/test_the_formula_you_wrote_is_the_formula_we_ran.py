# -*- coding: utf-8 -*-
"""依頼が**式そのものを書いた**のに別の計算をしたら ✓ を出さない ── その番人（2026-09-08）。

★★ 出所（盲検の検品が 20 件中**唯一の false ✓** として拾い、こちらで再現した）:

    表     商品 / 売上 / 原価              ← ★ 『利益』という列は**無い**
    依頼   「利益率（**利益÷売上**）の列を追加して」
    実行   操作:計算列 演算対象:**売上 と 原価** 演算子:/ 新しい列の名前:利益率
    実物   りんご **1.714**（頼んだ式なら 0.417）
    出力   **✓ 機械検証済み**

  ★ 既存の 2 つの関所はどちらも鳴らない:
      残差（列名）  『利益』は新しい列の名前『利益**率**』に**部分一致して消費される**
      効果の種類    計算列を頼んで計算列を実行 ── 食い違わない
    宣言と実体は完全に一致しているので事後条件も通る。★ また依頼だけが見られていない。

★ 直さない・止めない ── ⚠ を出して ✓ を降ろすだけ（residue/intent と同じ作法）。
"""
from __future__ import annotations

import argparse

import openpyxl
import pytest

import ailine
from ailine_core import arith


# --- ① 判定（純ロジック）-----------------------------------------------------

@pytest.mark.parametrize("task, decl, fires", [
    # ★ 実測した事故そのもの
    ("利益率（利益÷売上）の列を追加して",
     "操作:計算列 演算対象:売上 と 原価 演算子:/ 新しい列の名前:利益率", True),
    # ★ 頼んだとおりに割った回は黙る（対の試験 ── 鳴りっぱなしの番人は番人でない）
    ("利益率（利益÷売上）の列を追加して",
     "操作:計算列 演算対象:利益 と 売上 演算子:/ 新しい列の名前:利益率", False),
    # ★ 割り算は**順序が意味を変える**
    ("利益率（利益÷売上）の列を追加して",
     "操作:計算列 演算対象:売上 と 利益 演算子:/ 新しい列の名前:利益率", True),
    # ★ 演算子そのものが違う
    ("利益率（利益÷売上）の列を追加して",
     "操作:計算列 演算対象:利益 と 売上 演算子:* 新しい列の名前:利益率", True),
    # ★ 掛け算は入れ替えても同じ ── 順序では鳴らさない
    ("単価×件数の列を作って",
     "操作:計算列 演算対象:件数 と 単価 演算子:* 新しい列の名前:金額", False),
    # ★ 引き算は順序が意味を変える
    ("売上-原価の列を作って",
     "操作:計算列 演算対象:原価 と 売上 演算子:- 新しい列の名前:利益", True),
])
def test_the_formula_in_the_request_is_compared_with_the_one_declared(task, decl, fires):
    assert bool(arith.calculation_mismatch(task, decl)) is fires


@pytest.mark.parametrize("task, decl", [
    # ★★ 9,128 件の実走行で出た**唯一の誤爆**を焼き込む: 長音符『ー』を引き算と
    #   読み、「請求明細シ ー ト」で鳴っていた。記号表から外して 0 件になった。
    ("請求明細シートの数量と単価をかけた金額を出して",
     "操作:計算列 演算対象:数量 と 単価 演算子:* 対象列:金額"),
    # ★ 日付の区切りを式と読まない（両側が数字なら語ではない）
    ("締め日を2026/09/30にして", "操作:計算列 演算対象:売上 と 原価 演算子:/ 対象列:x"),
    # ★ 宣言が演算対象を持たない回は黙る（社名の `/` 等を式と読む危険がある）
    ("A社/B社の行を抜き出して", "操作:抽出 対象列:取引先 条件:等しい 値:A社"),
    # ★ 記号が 2 つ以上ある式は、どれを比べるか機械が決められないので黙る
    ("（売上-原価）÷売上の列を作って",
     "操作:計算列 演算対象:売上 と 原価 演算子:/ 新しい列の名前:利益率"),
])
def test_what_is_not_a_formula_does_not_ring(task, decl):
    assert arith.calculation_mismatch(task, decl) is None


def test_a_column_the_table_does_not_have_is_still_caught():
    """★ この事故の核心 ── 表に『利益』が無いので、機械は**別の列で代用した**。
       依頼側の語が実在の見出しかどうかは見ない（見ると、まさにこの回で黙る）。"""
    assert arith.calculation_mismatch(
        "利益率（利益÷売上）の列を追加して",
        "操作:計算列 演算対象:売上 と 原価 演算子:/ 新しい列の名前:利益率") == "利益÷売上"


# --- ② 配線（✓ を出す唯一の関所を実際に通す）--------------------------------

def _run_finish_apply(tmp_path, task, scope, capsys, name="in"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上表"
    ws.append(["商品", "売上", "原価"])
    ws.append(["りんご", 1200, 700])
    book = tmp_path / f"{name}.xlsx"
    wb.save(book)
    out = tmp_path / f"{name}.out.xlsx"
    out.write_bytes(book.read_bytes())
    work = tmp_path / f"w{name}"
    work.mkdir(exist_ok=True)
    a = argparse.Namespace(inplace=False, task=task, json=False, model="m",
                           keep_backups=3, copy=True)
    ailine._finish_apply(a, book, out, work, {"op": "COMPUTE_COLUMN"},
                         machine_verified=True, scope=scope, scope_note="",
                         warning_count=0)
    return capsys.readouterr().out


def test_the_gate_is_wired_into_the_only_place_that_prints_the_check(tmp_path, capsys):
    shown = _run_finish_apply(
        tmp_path, "利益率（利益÷売上）の列を追加して",
        "操作:計算列 演算対象:売上 と 原価 演算子:/ 新しい列の名前:利益率",
        capsys, name="bad")
    assert "利益÷売上" in shown and "⚠" in shown, shown
    assert "✓" not in shown, shown


def test_the_run_that_did_what_was_asked_keeps_its_check(tmp_path, capsys):
    """★ 対の試験 ── 頼んだ式どおりに計算した回は、今までどおり ✓ が出る。"""
    clean = _run_finish_apply(
        tmp_path, "利益率（利益÷売上）の列を追加して",
        "操作:計算列 演算対象:利益 と 売上 演算子:/ 新しい列の名前:利益率",
        capsys, name="good")
    assert "実行した計算はそれと違います" not in clean, clean
    assert "✓" in clean, clean
