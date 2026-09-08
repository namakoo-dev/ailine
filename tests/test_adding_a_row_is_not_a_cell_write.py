# -*- coding: utf-8 -*-
"""「〜の下に … 足して」は**行を増やす**依頼（2026-09-09）。

★★ 出所（到達率の器を広い op まで厚くした最初の一撃）:

    依頼   「みかんの**下に**「ぶどう」の行を**足して**」
    読み直し「『一括書換』でなく『1セル書換』として読み直しました
             ── 依頼文が『みかん』の行を名指ししています」
    実物   ★ **みかんの商品名が『ぶどう』に上書きされる**（行は増えず、みかんが消える）

  ★ 「下に」も「足して」も見ていない ── 三項の「依頼」が落ちた形だが、今回は
    **読み直し自身が落としている**（一段目の誤りを直そうとして別の誤りに着地した）。
  ★ ADD_ROW は走行の **18.1%**＝いちばん広い op。到達率の器の分母に無かったので
    誰も気づいていなかった（09-08 の作業前から在ることを A/B で確認済み）。

★★ 分かれ目は**向き**（実測でこう割れた）:

      行を増やす   〜の**下**に / **上**に / **間**に  ＋ 足す・追加・入れる
      1 セル       〜の**右**に / **左**に / **隣**に  （2026-09-08 に直した回・塞がない）

★ 1 セル書換へ着地する門は**2 つ**あるので、門ごとに書き足さず**入口で 1 度だけ**
  判定して両方が読む（片方だけ直る形を作らない）。
"""
from __future__ import annotations

import pytest

import ailine
from ailine_core.row_placement import task_places_a_new_row


def _asks(task: str) -> bool:
    return task_places_a_new_row(task, ailine._ANCHOR_AFTER, ailine._ANCHOR_BEFORE,
                                 ailine._re_between)


@pytest.mark.parametrize("task, want", [
    # ★ 実測した事故そのもの
    ("みかんの下に「ぶどう」の行を足して", True),
    ("りんごの上に1行追加して「もも」と入れて", True),
    ("りんごとみかんの間に1行入れて", True),
    ("みかんの下に空行を挿入して", True),
    # ★ 対の試験: 1 セルの言い方は塞がない（09-08 に直した回）
    ("みかんの右に東棟", False),
    ("鈴木の隣に東棟", False),
    # ★ 対の試験: 値の書換は行を増やさない
    ("梨の売上を2000にして", False),
    ("7行目の担当を「佐藤」にして", False),
    ("丸和物流の項目を「配送」にして", False),
    # ★★ 変異試験で分かった弱い所を埋めた（2026-09-09）: 下の 3 件が無いと、
    #   「動詞を見ない」変異も「向きを見ない」変異も**赤くならなかった**
    #   ── 塞いでいる条件を踏む検体が 1 つも無かった（検体の側の穴）。
    ("みかんの下にある行を削除して", False),   # ★ 位置の語は在るが**置く動詞が無い**
    ("みかんの右に「東棟」を入れて", False),   # ★ 置く動詞は在るが**向きが 1 セル側**
    ("みかんの下に東棟", False),               # ★ どちらも無い
    ("みかんの下の行を削除して", False),
])
def test_the_direction_tells_a_row_from_a_cell(task, want):
    assert _asks(task) is want


def test_the_words_come_from_one_place():
    """★ 位置の語彙を 2 箇所に持たない ── 判定は ailine.py の語彙を受け取るだけ。

    ★ 「無いこと」だけを見ると、探す場所が空でも通る（恒真）。**在ること**と対にする
      ── 2026-09-09 に既存の番人（test_guard_ledger）がこの粗を捕まえた。
    """
    import pathlib

    import ailine_core.row_placement as rp
    src = pathlib.Path(rp.__file__).read_text(encoding="utf-8")
    # ★ 在ること: 動詞の語彙は core が持つ（ここが空なら試験は無意味）
    assert "PLACING_VERBS" in src and "足し" in src, src[:200]
    # ★ 無いこと: 位置の語は ailine.py 側にだけ在る
    body = src.split('"""', 2)[-1]      # ★ docstring の例示は数えない
    assert "の下に" not in body and "の上に" not in body, (
        "位置の語を core 側にも書いている（増やす時に片方だけになる）")


def test_both_gates_read_the_same_judgement():
    """★★ 1 セル書換へ着地する門は 2 つ。**入口で 1 度**判定して両方が読む。

    ★ 門ごとに書き足すと、また片方だけ直る（09-08 に 8 回見た形）。
      ここは配線の形を静的に縛る ── 実機を起こさずに、片配線を機械で止める。
    """
    import pathlib
    src = (pathlib.Path(ailine.__file__)).read_text(encoding="utf-8")
    lands = src.count('"op": "SET_CELL_VALUE"')
    reads = src.count("_wants_new_row")
    assert lands == 2, f"1 セル書換へ着地する所が {lands} 箇所（2 のはず）── 増えたら配線も要る"
    # ★ 1 回の代入 + 2 つの門で読む = 3
    assert reads == 3, f"_wants_new_row の出現が {reads}（代入 1 + 門 2 = 3 のはず）"
