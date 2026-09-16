# -*- coding: utf-8 -*-
"""★★ 条件つき書換を、表の**構造を変える op** に化けさせない（全 op を 1 本で縛る）。

★ なぜ 1 本にまとめてあるか（2026-09-16 の掃き出し）:
  盲検の買い手役 2 体目が 0 円と答えた唯一の理由は、
  「納期が2026/09/30より後の行のチェック列に★を入れて」が**行追加**に化け、
  受注台帳に取引先も金額も空の行が EXIT=0 で入ったことだった。
  朝の直しは `ADD_ROW` **1 op だけ**に関所を置いた。ところが同じ文を
  表の構造を変える 9 op 全部に通すと **4 op が通っていた**
  （DELETE_COLUMN／INSERT_ROWS／APPEND_TOTAL／COMPUTE_COLUMN）。
  ★ だから検体も op ごとに書かない ── **名簿から回して全 op を 1 本で縛る**。
    新しい構造 op を足した人は、名簿に足した瞬間にこの試験の対象になる。

★ 線の引き方（実測で決めた・詳細は ailine_core/intent.py のコメント）:
  断るのは「**書き込む動詞**が在る」かつ「**行き先の列**を名指し、または**条件の語**」。
  助詞だけで断った初版は、正当な依頼 195 件のうち 56 件を誤って断った。
"""
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from ailine_core import compare_words, intent as intent_mismatch  # noqa: E402

STRUCTURAL_OPS = sorted(intent_mismatch.STRUCTURAL_EFFECT)

_META_ORDERS = {
    "sheets": ["受注台帳"],
    "headers": {"受注台帳": ["受注番号", "取引先", "納期", "金額", "チェック", "状態"]},
    "header_rows": {"受注台帳": 1},
    "rows": {"受注台帳": 12},
}

#: 事故の形 ── どれも「条件に合う行の、ある列に値を置く」依頼であって、行や列は増減しない。
ACCIDENTS = [
    "納期が2026/09/30より後の行のチェック列に★を入れて",        # ★ 買い手の実文そのもの
    "納期が2026/09/30より後の行だけ、チェック列を★に書き換えて",  # ★ 言い直しても同じだった
    "金額が50000以上の行のチェック列に★を入れて",
    "状態が未手配の行の備考列に「至急」と入れて",
    # ★ 「列」と書かない形 ── **条件の語の枝**だけが受け持つ。
    "金額が50000以上の行に★を入れて",
    "数量が1000より多い行に「大口」と入れて",
]

#: 陰性対照 ── その op を**本当に**頼む言い方。1 件も落とさないこと。
#: ★ 出典は repo が既に持っている検体（bench/tests から集めた 195 件の代表）。
#:   ここを落とすと、買い手は正しい依頼を断られて仕事が進まない ── 誤爆は誤動作より高くつく。
LEGITIMATE = {
    "ADD_ROW": ["5行目に丸山工業の行を追加して",
                 "北斗精機の行の下に、取引先「西村工業」の行を追加して、項目は事務机、件数は2、単価は15000にして"],
    "INSERT_ROWS": ["2行目に空の行を1つ挿入して", "3行目に空行を2行挿入して"],
    "APPEND_TOTAL": ["単価列の合計行に単価の合計を書いて", "最終行に金額の合計行を追加して"],
    "ADD_COLUMN": ["備考という列を作って", "すいかの右に列を追加して"],
    "COMPUTE_COLUMN": ["数量に単価を掛けた列を作って", "金額に1.1を掛けた列を追加して"],
    "LOOKUP_FILL": ["2冊を照合して", "商品コードで商品マスタを引いて商品名の列を埋めて"],
    "SPLIT_CELL": ["URL の列を改行で分けて", "住所をスペースで分けて2列にして"],
    "DELETE_ROWS": ["数量が100未満の行を削除して", "金額が空の行を削除して"],
    "DELETE_COLUMN": ["備考の列を削除して", "原価列を削除して"],
}


def test_the_roster_of_structural_ops_is_not_empty():
    """★ 名簿が空でも他の試験は全部緑になる（空回りの検出）。"""
    assert len(STRUCTURAL_OPS) >= 9, STRUCTURAL_OPS


@pytest.mark.parametrize("op", STRUCTURAL_OPS)
@pytest.mark.parametrize("task", ACCIDENTS)
def test_no_structural_op_accepts_a_conditional_write(op, task):
    """★ 9 op × 6 文 = 54 通り。1 つでも通れば、頼んでいない行や列が台帳に残る。"""
    ok, _r, _i, err = ailine.verify_dsl_args(op, {}, _META_ORDERS, task=task,
                                              target_sheet="受注台帳")
    assert not ok, f"{op} が条件つき書換を受け取った: {task}"
    effect = intent_mismatch.STRUCTURAL_EFFECT[op][0]
    assert effect in err, f"断り文がこの op の帰結（{effect}）を言っていない: {err}"
    assert "のように" in err, f"断るだけで、通る書き方を示していない: {err}"


@pytest.mark.parametrize("op", STRUCTURAL_OPS)
def test_every_structural_op_still_accepts_its_own_real_request(op):
    """★ 陰性対照 ── 名簿の全 op に、正当な言い方の検体が要る（無い op は赤）。"""
    tasks = LEGITIMATE.get(op)
    assert tasks, f"{op} の正当な依頼の検体が無い（誤爆を測れていない）"
    for task in tasks:
        refusal = intent_mismatch.refuse_structural_op_that_is_really_a_write(
            task, op, comparison_in_task=compare_words.read(task).hit)
        assert refusal is None, f"正当な依頼を止めた: {op} / {task} → {refusal}"


def test_the_refusal_quotes_the_words_the_person_wrote():
    """★ 断り文は、人が自分の依頼文の中に見つけられる語を名指しすること。"""
    err = intent_mismatch.refuse_structural_op_that_is_really_a_write(
        "状態が未手配の行の備考列に「至急」と入れて", "ADD_ROW", comparison_in_task=True)
    assert "備考列" in err, err
