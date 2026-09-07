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


def _effects():
    """op ごとの**効果の種類**（登録簿そのまま）。"""
    return {op: set(getattr(ailine.OP_WRITE_TARGET.get(op), "writes", ()) or ())
            for op in ailine.OP_META}


#: 解釈行の代わり ── ★ **実物に合わせる**。宣言が薄いと「当たった語が宣言に出ている」
#: という拒否（残差と同じ考え）を素通りさせ、判定でなく検体の粗を測ることになる。
DECLS = {
    "COMPUTE_COLUMN": "操作:計算列 新しい列の名前:小計 演算対象:数量 演算子:* 単価",
    "SORT": "操作:並べ替え 対象:金額 順:降順",
    "DELETE_ROWS": "操作:行削除 削除位置:3 行数:1",
    "EXTRACT": "操作:抽出 対象列:取引先 条件:等しい 値:ヤマノ食品",
    "DEDUP": "操作:重複除去 対象列:品名",
    "ADD_ROW": "操作:行追加 挿入位置:2 入れる値:品名=棚",
    "SWAP": "操作:入れ替え 入れ替える一方:机 もう一方:棚",
}


def _decl(op):
    return DECLS.get(op, f"操作:{ailine.OP_LABELS.get(op, op)}")


# --- ① 効果の種類の食い違い（純ロジック）------------------------------------

@pytest.mark.parametrize("task, op, fires", [
    ("ヤマノ食品の行を削除して", "EXTRACT", True),      # ★ 実測した事故そのもの
    ("品名が重複している行を消して", "DEDUP", True),     # ★ 同じ形（新シートを作るだけ）
    ("ヤマノ食品の行を削除して", "DELETE_ROWS", False),  # 取り除く op なら食い違わない
    ("数量に単価をかけた小計の列を追加して", "COMPUTE_COLUMN", False),  # ★ どちらも取り除かない
    ("重複行を削除して重複を除く", "DEDUP", False),      # ★ 自分の語彙が名指しされている
])
def test_a_removal_request_that_removes_nothing_is_named(task, op, fires):
    got = intent.op_effect_mismatch(task, {op}, _decl(op), _pools(), _effects())
    assert bool(got) is fires, got


@pytest.mark.parametrize("task, op, fires", [
    # ★★ 2026-09-07: 効果の行列が掴んだ本物の欠陥 ── 「交換して」が 23 回中 1 回
    #   **行追加**に化け、行が 4 → 5 に増えたのに成功と報告していた。
    #   ★ 取り除き限定の版では黙る（入れ替えも行追加も取り除かない）。
    ("机の行と棚の行を交換して", "ADD_ROW", True),
    ("机の行と棚の行を交換して", "SWAP", False),          # ★ 正しく入れ替えた回は黙る
    ("あかね商事の行とうえだ物産の行を交換して", "ADD_ROW", True),
])
def test_a_swap_that_became_an_insert_is_named(task, op, fires):
    got = intent.op_effect_mismatch(task, {op}, _decl(op), _pools(), _effects())
    assert bool(got) is fires, got


def test_op_names_are_not_what_we_compare():
    """★ op 名で比べると誤爆する ── 効果の種類で見ていることを縛る。

    「列を追加して」は ADD_COLUMN の語彙に当たるが、実行した COMPUTE_COLUMN も
    列を書く op なので**食い違いではない**。実測でここが 22 件を占めていた。
    """
    assert intent.op_effect_mismatch(
        "数量に単価をかけた小計の列を追加して", {"COMPUTE_COLUMN"},
        "操作:計算列 新しい列の名前:小計 演算対象:数量 単価", _pools(), _effects()) == []


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


def test_a_swap_that_became_an_insert_loses_the_check(tmp_path, capsys):
    """★ 実物の宣言で、✓ が消えることまで縛る（判定だけ正しくても意味が無い）。

    ★ 実測の宣言をそのまま使う ── 「操作:行追加 挿入位置:2 位置の根拠:『机』の行＝2行目
      入れる値:品名=棚」。ここに『交換』は出ていないので、依頼の語が落ちている。
    """
    shown = _run_finish_apply(
        tmp_path, "机の行と棚の行を交換して",
        "操作:行追加 挿入位置:2 位置の根拠:『机』の行＝2行目 入れる値:品名=棚",
        "ADD_ROW", capsys, name="swapbad")
    assert "交換" in shown and "⚠" in shown, shown
    assert "✓" not in shown, shown


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
