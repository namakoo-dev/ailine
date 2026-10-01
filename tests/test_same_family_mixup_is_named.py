# -*- coding: utf-8 -*-
"""**同じ系統**の op どうしの取り違えにも ✓ を出さない ── その番人（2026-10-01）。

★★ 出所: 効果の種類の食い違い（intent.op_effect_mismatch）は、拒否②「2 つの op の効果が
  重なれば黙る」のせいで、同じ系統の中の取り違えに**原理的に鳴らなかった**。
  台帳（tests/warning_register.json の blind_spots）に本物が 2 件残っていた:

    「税込み金額の順番を逆にして」→ 操作:入れ替え（税込み金額 ↔ 締め日）で ✓
    「件数の合計も合計行に入れて」→ 操作:行追加（値は文字列『合計』）で ✓

★ 拒否②を外すだけでは戻らない ── ②を入れた理由（「列を追加して」で計算列が走る
  上位下位の誤爆）が同じ成功回で 13 件戻ってくる。だから②は**狭めた**:
  黙るのは、実行した op が依頼の op の**一種**と宣言から読める時だけ。
    ・効果: 依頼の op に、実行した op の持たない効果が在る（INSERT_ROWS は ADD_ROW の一部）
    ・語彙: 実行した op の語が依頼の op の語を含み、差の部分が依頼に在る
            （『重複行を削除』⊃『行を削除』で、依頼に『重複』が在る）
★ 測った数（読むだけの別スクリプト・報告に全件）: 実走行の成功回 1,312 件で新しく鳴るのは
  上の 2 件だけ／失敗回まで含めた 1,618 件で 10 件・全部が本物の取り違え／正しい op を付けた
  凍結検体 190 件で 0 件。
"""
from __future__ import annotations

import argparse

import openpyxl
import pytest

import ailine
from ailine_core import intent


def _pools():
    return {op: [p for p in ailine._op_match_pool(op) if p] for op in ailine.OP_META}


def _effects():
    """op ごとの**効果の種類**（登録簿そのまま）。"""
    return {op: set(getattr(ailine.OP_WRITE_TARGET.get(op), "writes", ()) or ())
            for op in ailine.OP_META}


def _fires(task, ops, decl):
    ops = {ops} if isinstance(ops, str) else set(ops)
    return intent.op_effect_mismatch(task, ops, decl, _pools(), _effects())


# --- ① 系統ごとに 1 つ以上の本物の取り違え → 鳴る ----------------------------
#: 宣言は**実物の形**に合わせる（薄いと拒否③を素通りさせて検体の粗を測ることになる）。
#: 実走行の宣言が在るものはそのまま写した。

MIXUPS = [
    # 書式だけ: 太字を頼んで背景色
    ("見出しを太字にして", "FILL_COLOR", "操作:背景色 対象:row:1 色:yellow", "太字"),
    # 並べ替え: 並べ替えを頼んで入れ替え
    ("金額で並べ替えて", "SWAP",
     "操作:入れ替え 入れ替える一方:金額 もう一方:在庫 何を入れ替えるか:列（見出しで一致）", "並べ替え"),
    # ★ 実走行の本物（台帳の 1 件目・宣言はそのまま）
    ("税込み金額の順番を逆にして", "SWAP",
     "解釈: 操作:入れ替え 入れ替える一方:税込み金額 もう一方:締め日 何を入れ替えるか:列（見出しで一致）",
     "順番"),
    # 消す: 列を消してと頼んで行削除
    ("備考の列を消して", "DELETE_ROWS", "操作:行削除 削除位置:3 行数:1", "列を消して"),
    # 消す: 名指しの行を消してと頼んで重複行の削除（★ 含むだけで黙らせると見逃す形）
    ("鈴木の行を消して", "DEDUP_DELETE", "操作:重複行の削除 判定キー:氏名", "行を消して"),
    # 新しいシート: 集計を頼んで抽出
    ("取引先ごとに金額を集計して", "EXTRACT",
     "操作:抽出 対象列:取引先 条件:等しい 値:ヤマノ食品", "集計"),
    # 既存列に書く: 計算を頼んで転記
    ("単価と数量の掛け算で金額を出して", "LOOKUP_FILL",
     "操作:転記 キー列:品名 参照先:マスタ 書き込む列:金額", "掛け算"),
    # ★ 実走行の本物（台帳の 2 件目・宣言はそのまま）
    ("件数の合計も合計行に入れて", "ADD_ROW",
     "解釈: 操作:行追加 挿入位置:6 位置の根拠:『合計』の行＝6行目 入れる値:件数=合計 "
     "入れる位置:『合計』の行＝6行目", "合計行"),
]


@pytest.mark.parametrize("task, op, decl, word", MIXUPS)
def test_a_mixup_within_the_same_family_is_named(task, op, decl, word):
    got = _fires(task, op, decl)
    assert word in got, (task, op, got)


# --- ② 同じ依頼で正しい op → 黙る（対で縛る）-------------------------------

RIGHT_OP = {
    "見出しを太字にして": ("BOLD", "操作:太字 対象:row:1"),
    "金額で並べ替えて": ("SORT", "操作:並べ替え 対象:金額 順:昇順"),
    "税込み金額の順番を逆にして": ("SORT", "操作:並べ替え 対象:税込み金額 順:降順"),
    "備考の列を消して": ("DELETE_COLUMN", "操作:列削除 対象:備考"),
    "鈴木の行を消して": ("DELETE_ROWS", "操作:行削除 削除位置:3 位置の根拠:『鈴木』の行＝3行目 行数:1"),
    "取引先ごとに金額を集計して": ("AGGREGATE", "操作:集計 グループ:取引先 値:金額"),
    "単価と数量の掛け算で金額を出して": ("COMPUTE_COLUMN", "操作:計算列 演算対象:単価 と 数量 演算子:*"),
    "件数の合計も合計行に入れて": ("APPEND_TOTAL", "操作:合計追加 対象:件数"),
}


@pytest.mark.parametrize("task", sorted(RIGHT_OP))
def test_the_right_op_for_the_same_request_stays_quiet(task):
    op, decl = RIGHT_OP[task]
    assert _fires(task, op, decl) == [], (task, op)


def test_every_family_has_a_mixup_and_its_right_op():
    """★ 検体の側の番人: 系統（効果の集合）ごとに、鳴る検体と黙る対が揃っていること。"""
    eff = _effects()
    have = {frozenset(eff[op]) for _, op, _, _ in MIXUPS}
    for rep in ("BOLD", "SORT", "DELETE_ROWS", "AGGREGATE"):   # 書式・並べ替え・消す・新シート
        assert frozenset(eff[rep]) in have, rep
    assert any(ailine.WRITE_EXISTING_COLUMN in f for f in have)  # 既存列に書く
    assert {t for t, _, _, _ in MIXUPS} == set(RIGHT_OP)


# --- ③ 上位下位（実行した op が依頼の op の一種）→ 黙る -------------------------

@pytest.mark.parametrize("task, ops, decl", [
    # 効果で読む一種: 計算列は「列を作る」の一種（②を入れた理由・実測 22 件）
    ("数量に単価をかけた小計の列を追加して", "COMPUTE_COLUMN",
     "操作:計算列 新しい列の名前:小計 演算対象:数量 演算子:* 単価"),
    # 効果で読む一種: 空行の挿入は「行を足す」の一種（実走行の宣言そのまま）
    ("丸和物流と近江スチールの間に1行足して", "INSERT_ROWS",
     "解釈: シート:『8月請求』(1枚目) 操作:行挿入 挿入位置:3 "
     "位置の根拠:『丸和物流』（2行目）の下＝3行目（『近江スチール』との間）"),
    # 語彙で読む一種: 依頼に『重複』が在る（道具自身が案内する --op DEDUP_DELETE の形）
    ("重複行を消して", "DEDUP_DELETE", "操作:重複行の削除 判定キー:品名"),
    # 語彙で読む一種: 依頼に『クロス』が在る（凍結検体の正解）
    ("商品×月のクロスで集計して", "PIVOT", "操作:ピボット 行:商品 列:月 値:売上"),
])
def test_a_kind_of_the_requested_op_stays_quiet(task, ops, decl):
    assert _fires(task, ops, decl) == [], (task, ops)


# --- ④ 複合計画: 別の段の語は、その段が実行していれば黙る ---------------------

def test_another_steps_word_in_a_plan_stays_quiet():
    task = "60以上を抜き出してから売上シートを現場ごとに集計して"
    assert _fires(task, {"PLAN", "EXTRACT", "AGGREGATE"}, "") == []
    # ★ 対で縛る: 集計の段が無ければ（抽出だけ）、同じ依頼で鳴る
    assert "集計" in _fires(task, {"EXTRACT"}, "")


# --- ⑤ 配線（✓ を出す唯一の関所を実際に通す）--------------------------------

def _run_finish_apply(tmp_path, task, scope, op, capsys, name="in", ops=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "8月請求"
    ws.append(["取引先", "金額", "税込み金額", "締め日"])
    ws.append(["丸和物流", 1000, 1100, "月末"])
    book = tmp_path / f"{name}.xlsx"
    wb.save(book)
    out = tmp_path / f"{name}.out.xlsx"
    out.write_bytes(book.read_bytes())
    work = tmp_path / f"w{name}"
    work.mkdir(exist_ok=True)
    a = argparse.Namespace(inplace=False, task=task, json=False, model="m",
                           keep_backups=3, copy=True)
    result = {"op": op, **({"ops": ops} if ops else {})}
    ailine._finish_apply(a, book, out, work, result, machine_verified=True,
                         scope=scope, scope_note="", warning_count=0)
    return capsys.readouterr().out


def test_a_same_family_mixup_loses_the_check(tmp_path, capsys):
    """★ 実走行の宣言で、✓ が消え、断り文が**嘘をつかない**ことまで縛る。

    ★ 入れ替えは列を動かしているので「取り除きません（元の表はそのまま残っています）」と
      書けば、それ自体が嘘になる ── 実行した操作の名前を出す。
    """
    shown = _run_finish_apply(
        tmp_path, "税込み金額の順番を逆にして",
        "操作:入れ替え 入れ替える一方:税込み金額 もう一方:締め日 何を入れ替えるか:列（見出しで一致）",
        "SWAP", capsys, name="swap")
    assert "⚠" in shown and "順番" in shown, shown
    assert "『入れ替え』です" in shown, shown
    assert "取り除きません" not in shown, shown
    assert "✓" not in shown, shown


def test_the_right_op_keeps_the_check(tmp_path, capsys):
    clean = _run_finish_apply(
        tmp_path, "税込み金額の順番を逆にして",
        "操作:並べ替え 対象:税込み金額 順:降順", "SORT", capsys, name="sort")
    assert "順番』と読めます" not in clean, clean
    assert "✓" in clean, clean


def test_a_plan_hands_every_step_to_the_judge(tmp_path, capsys):
    """★ 複合計画は op が『PLAN』で、段は ops に在る ── 段を渡し損ねると正しい計画で鳴る。

    ★ 検体は 2026-09-07 に複合計画で誤爆した依頼そのもの（『列を足して』は計算列の段の語・
      並べ替えの段の語『並べ替え』が依頼に在るので拒否①で黙るのが正しい）。
      ★ 段の操作名が解釈行に出ている形（『集計して』と 操作:集計）は拒否③でも黙るので、
        段を渡し損ねても緑のままになる ── 変異試験で 1 度すり抜けた。
    """
    plan = _run_finish_apply(
        tmp_path, "単価で並べ替えして、数量と単価を掛けた金額の列を足して",
        "操作:並べ替え 対象:単価 順:昇順; 操作:計算列 演算対象:数量 と 単価 演算子:*",
        "PLAN", capsys, name="plan", ops=["SORT", "COMPUTE_COLUMN"])
    assert "と読めますが" not in plan, plan
    # ★ 対で縛る: 計画の段が依頼と食い違えば鳴り、名前は段の操作名で出す（『PLAN』と書かない）
    bad = _run_finish_apply(
        tmp_path, "金額と締め日の間に税込み金額列を追加",
        "操作:計算列 演算対象:金額 演算子:*; 操作:行挿入 挿入位置:2", "PLAN", capsys,
        name="planbad", ops=["COMPUTE_COLUMN", "INSERT_ROWS"])
    assert "『計算列・行挿入』です" in bad, bad
    assert "PLAN" not in bad and "✓" not in bad, bad
