# -*- coding: utf-8 -*-
"""頼まれた操作の**種類**と食い違ったら ✓ を出さない ── その番人（2026-09-07）。

★★ 出所（外部の検品が最重の所見として拾い、こちらで再現した）:

    依頼   「ヤマノ食品の行を**削除して**」
    実行   操作:**抽出** → 新しいシートを作り、元の 6 行はそのまま
    出力   **✓ 機械検証済み**

  事後条件は「抽出として正しいか」を確かめるので通る ── **宣言と実体は一致していて、
  依頼だけが落ちている**。★ 既存の関所は**列名**しか見ないので、動詞は拾えなかった。

★ 併せて、その既存の関所が黙っていた**別の穴**も塞いだ:

    「原価**列**の右隣に備考列を追加して」→ 残差は『原価列』で、見出し『原価』と
    完全一致せず**鳴らなかった**。実害は「備考が末尾に入り、解釈行が
    『依頼文に位置の指定が無いため』と**嘘をついて** ✓ を出す」だった。

★ どちらも直さない・止めない ── ⚠ を出して ✓ を降ろすだけ。
"""
from __future__ import annotations

import openpyxl
import pytest

import ailine
from ailine_core import intent, residue


def _pools():
    return {op: [p for p in ailine._op_match_pool(op) if p] for op in ailine.OP_META}


def _removes():
    return {op: (ailine.WRITE_REMOVE in
                 (getattr(ailine.OP_WRITE_TARGET.get(op), "writes", ()) or ()))
            for op in ailine.OP_META}


# --- ① 効果の種類の食い違い（純ロジック）------------------------------------

@pytest.mark.parametrize("task, op, fires", [
    ("ヤマノ食品の行を削除して", "EXTRACT", True),      # ★ 実測した事故そのもの
    ("品名が重複している行を消して", "DEDUP", True),     # ★ 同じ形（新シートを作るだけ）
    ("ヤマノ食品の行を削除して", "DELETE_ROWS", False),  # 取り除く op なら食い違わない
    ("数量に単価をかけた小計の列を追加して", "COMPUTE_COLUMN", False),  # ★ どちらも取り除かない
    ("重複行を削除して重複を除く", "DEDUP", False),      # ★ 自分の語彙が名指しされている
])
def test_a_removal_request_that_removes_nothing_is_named(task, op, fires):
    got = intent.removal_asked_but_not_done(task, op, _pools(), _removes())
    assert bool(got) is fires, got


def test_op_names_are_not_what_we_compare():
    """★ op 名で比べると誤爆する ── 効果の種類で見ていることを縛る。

    「列を追加して」は ADD_COLUMN の語彙に当たるが、実行した COMPUTE_COLUMN も
    列を書く op なので**食い違いではない**。実測でここが 22 件を占めていた。
    """
    assert intent.removal_asked_but_not_done(
        "数量に単価をかけた小計の列を追加して", "COMPUTE_COLUMN",
        _pools(), _removes()) == []


# --- ② 見出し + 構造の語（『原価列』→『原価』）------------------------------

@pytest.mark.parametrize("task, want", [
    ("原価列の右隣に備考列を追加して", ["原価"]),   # ★ 昨日の関所が黙っていた形
    ("原価の右に備考の列を追加して", ["原価"]),
    ("原価率の列を追加して", []),                    # ★ 別の語まで拾わない
])
def test_a_header_with_a_structure_word_still_counts(task, want):
    decl = "操作:列追加 新しい列の名前:備考 入れる位置:末尾＝5列目（依頼文に位置の指定が無いため）"
    got = residue.unaccounted_request_words(
        task, decl, ailine._op_match_pool("ADD_COLUMN"), {"品名", "売上", "原価", "利益"})
    assert got == want, got


# --- ③ 配線（✓ を出す唯一の関所を実際に通す）--------------------------------

def _run_finish_apply(tmp_path, task, scope, op, capsys, name="in"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["品名", "売上", "原価", "利益"])
    ws.append(["りんご", 1000, 600, 400])
    book = tmp_path / f"{name}.xlsx"
    wb.save(book)
    out = tmp_path / f"{name}.out.xlsx"
    out.write_bytes(book.read_bytes())
    work = tmp_path / f"w{name}"
    work.mkdir(exist_ok=True)
    import argparse
    a = argparse.Namespace(inplace=False, task=task, json=False, model="m",
                           keep_backups=3, copy=True)
    ailine._finish_apply(a, book, out, work, {"op": op}, machine_verified=True,
                         scope=scope, scope_note="", warning_count=0)
    return capsys.readouterr().out


def test_the_mismatch_gate_is_wired_into_the_only_place_that_prints_the_check(
        tmp_path, capsys):
    shown = _run_finish_apply(
        tmp_path, "ヤマノ食品の行を削除して",
        "操作:抽出 対象列:取引先 条件:等しい 値:ヤマノ食品", "EXTRACT", capsys)
    assert "取り除きません" in shown, shown
    assert "✓" not in shown, shown

    clean = _run_finish_apply(
        tmp_path, "ヤマノ食品の行を削除して",
        "操作:行削除 削除位置:3 行数:1", "DELETE_ROWS", capsys, name="ok")
    assert "取り除きません" not in clean, clean
