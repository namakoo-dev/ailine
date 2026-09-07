# -*- coding: utf-8 -*-
"""査定者が挙げた false ✓ を、**3 件とも**出さないこと（2026-09-07）。

★★ 出所: 外部の査定が「私が受け取った ✓ は 8 件、うち 3 件（37.5%）は頼んだことと
  違うことをして ✓ を出した」と書いた、その 3 件。相手の条件は
  「`✓` の意味を『宣言どおり』から『依頼どおり』に締めて、false ✓ が 0 になったら」。

    ① ヤマノ食品の行を削除して          → 抽出（新シート・元は無傷）で ✓
    ② 田中さんの分だけ残して他は消して  → 同上
    ③ 金額に円マークを付けて            → 桁区切り（¥ 無し）で ✓

★ ①は「効果の種類」で捕まえた（取り除く語 vs remove を書かない op）。
★ ②は素通りしていた ── op の照合語彙が「**行を**消して」の形しか持たず、
  「他は消して」に当たらなかった。**裸の動詞**まで見るようにした。
  ★ 広げる前に測った: 4,556 件の実走行で 0.37% → 0.70%。増えた 15 件はすべて
    DEDUP（新シートを作るだけ＝元の重複は残る）で、既に本物と数えた家系。誤爆 0。
★ ③は**宣言しか見ていなかった**。`style != "thousands"` を弾く番人は在るのに、
  LLM が「円マーク」を thousands へ**正規化して**返すので素通りする。
  ── 三項（依頼・宣言・実体）のうち、また依頼が落ちる形。依頼文の側を見る。
"""
from __future__ import annotations

import pytest

import ailine
from ailine_core import intent


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


@pytest.mark.parametrize("task, op", [
    ("ヤマノ食品の行を削除して", "EXTRACT"),          # ★ 査定の①
    ("田中さんの分だけ残して他は消して", "EXTRACT"),  # ★ 査定の②（裸の動詞）
    ("品名が重複している行を消して", "DEDUP"),        # ★ 同じ家系（新シートを作るだけ）
])
def test_the_three_false_checks_are_caught(task, op):
    assert intent.op_effect_mismatch(task, {op}, _decl(op), _pools(), _effects()), task


@pytest.mark.parametrize("task, op", [
    ("ヤマノ食品の行を削除して", "DELETE_ROWS"),   # ★ 取り除く op なら黙る
    ("金額の大きい順に並べ替えて", "SORT"),
    ("数量に単価をかけた小計の列を追加して", "COMPUTE_COLUMN"),
])
def test_it_stays_quiet_when_nothing_is_wrong(task, op):
    assert not intent.op_effect_mismatch(task, {op}, _decl(op), _pools(), _effects()), task


@pytest.mark.parametrize("task, want", [
    ("金額に円マークを付けて", "通貨記号（¥）"),           # ★ 査定の③
    ("金額に円マークとカンマつけて", "通貨記号（¥）"),
    ("単価をパーセント表示にして", "百分率（%）"),
    ("金額に桁区切りを付けて", None),                       # ★ 持っている書式は通す
    ("日付の列に桁区切りを付けて", None),
])
def test_a_format_we_do_not_have_is_named_not_silently_swapped(task, want):
    got = intent.format_asked_but_not_supported(
        task, ["金額", "単価", "日付", "締め日", "取引先"])
    assert got == want, got


def test_a_column_named_like_a_format_does_not_trip_it():
    """★ 列名に当たる語では鳴らさない ── 「通貨」という列を持つ表は実在しうる。

    ★ この検体を足すまで、除外を殺しても試験は緑のままだった（2026-09-07 の変異試験で
      発覚）。**番人でなく試験が見ていなかった**方の抜け ── 今日 3 度目。
    """
    cols = ["品名", "通貨", "小数", "金額"]
    assert intent.format_asked_but_not_supported("通貨の列に桁区切りを付けて", cols) is None
    assert intent.format_asked_but_not_supported("小数の列に桁区切りを付けて", cols) is None
    # ★ 対で縛る: 同じ語でも、列に無ければ書式の指定として鳴る
    assert intent.format_asked_but_not_supported(
        "金額を通貨表示にして", ["品名", "金額"]) == "通貨記号（¥）"


def test_the_format_refusal_reaches_the_real_path(tmp_path):
    """★ 判定が正しくても、配線されていなければ何も守らない。"""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["品名", "金額"])
    ws.append(["りんご", 1200])
    p = tmp_path / "b.xlsx"
    wb.save(p)
    ok, _res, _inf, err = ailine.verify_dsl_args(
        "NUMBER_FORMAT", {"col": "金額", "style": "thousands"},
        ailine.build_book_meta(p), task="金額に円マークを付けて", vocab=ailine.load_vocab())
    assert not ok, "円マークを頼まれて桁区切りを当てている"
    assert "通貨記号" in err and "桁区切り" in err, err

    ok2, _r2, _i2, err2 = ailine.verify_dsl_args(
        "NUMBER_FORMAT", {"col": "金額", "style": "thousands"},
        ailine.build_book_meta(p), task="金額に桁区切りを付けて", vocab=ailine.load_vocab())
    assert ok2, err2
