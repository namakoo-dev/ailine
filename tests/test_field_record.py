# -*- coding: utf-8 -*-
"""区分の導出（確 / 単 / 割 / 無）と、空欄に理由が付くこと。

★★ この試験の主眼は 3 つ。どれも**実際に踏んだ事故**から来ている。

  ① 結合セルの写しを 2 つ目の根拠に数えない（2026-09-10）
     実測: 畳まずに数えると 40 冊すべて「確」。アンカーで畳むと 確 14 / 単 26。
     26 冊は**自分の写しと一致して「裏が取れた」**と言っていた。
     ★ この repo は買い手の $0 条件として「検算が本体と同じ口から出るなら
       それは検算でなく感想」を掲げている。その検算を測る採点器がそれをやっていた。

  ② 空欄に理由が付かない経路を作らせない（2026-09-11 凍結・Namakoo）
     空欄だけだと「取れなかった」と「矛盾していた」が同じ顔になる。
     ★ 型で守る（理由を持たずに空欄の Record を組み立てられない）。

  ③ ★ 一致だけで「確」と言わない（2026-09-11、未見 51 冊の実測）
     初版は「開いた 2 つの口が一致した」だけで 確 を立てていた。
     実測で誤区分になった 12 件すべてで、**開いていない第三の口**が食い違っていた
     （源泉徴収の振込金額／備考の再掲／明細の合計／2 枚目のシート／通貨／
       そもそも納品書か見積書か）。
     ★ 一致は「他の口が黙っている」ことを意味しない ── **開いていないだけ**。
     → 確 は「一致した ＋ 食い違う口が無いことを**見た**（swept）」を意味する。
"""
from __future__ import annotations

import pytest

from ailine_core.field_record import (CONFIRMED, NONE_FOUND, SINGLE, SPLIT,
                                       Evidence, Record, both_sides, describe,
                                       grade, grade_of, value)


def ev(at, val, rule="規則", how="読み方"):
    return Evidence(rule=rule, value=val, at=at, how=how)


def rec(field, evidences, **kw):
    """★ 掃き出し済みとして組む（区分の中身を試したい試験の既定）。"""
    kw.setdefault("swept", True)
    kw.setdefault("swept_how", "（試験）他に食い違う数字が無いことを確かめた")
    return Record(field, evidences, **kw)


# ── 区分の導出 ────────────────────────────────────────────
def test_two_independent_sources_that_agree_are_confirmed():
    r = rec("請求額", (ev("C11", 3300), ev("H39", 3300)))
    assert grade(r) == CONFIRMED
    assert value(r) == 3300


def test_one_source_is_single_and_still_gives_a_value():
    r = rec("請求額", (ev("C11", 3300),))
    assert grade(r) == SINGLE
    assert value(r) == 3300, "★ 単 は値を出す（矛盾していないだけ・裏が無い）"


def test_sources_that_disagree_are_split_and_give_no_value():
    r = rec("請求額", (ev("C11", 3300), ev("H39", 3301)),
            blank_reason="上部の請求額と帯の合計が 1 円ずれています")
    assert grade(r) == SPLIT
    assert value(r) is None, "★ 割 は値を出さない（2026-09-11 に凍結した判断）"


def test_no_source_is_none_found():
    r = rec("請求元", (), blank_reason="連絡先のラベルが見つかりません")
    assert grade(r) == NONE_FOUND
    assert value(r) is None


# ── ★ ①恒真の再演を防ぐ ────────────────────────────────────
def test_the_same_cell_read_twice_is_not_two_sources():
    """★ 結合の写しを 2 つ目の根拠に数えない（2026-09-10 の恒真）。

    同じアンカー番地から 2 回読んでも、独立した根拠は 1 つ。
    ★ これが「確」になると、**自分の写しと一致して裏が取れた**ことになる。
    """
    r = rec("請求額", (ev("C11", 3300, how="ラベルの右"),
                       ev("C11", 3300, how="結合の写し")))
    assert grade(r) == SINGLE, (
        "★ 同じセルを 2 回数えて『確』になった ── 恒真が再演している")


def test_three_reads_from_two_cells_are_two_sources():
    """3 回読んでも、出所が 2 つなら 2 つ。"""
    r = rec("請求額", (ev("C11", 3300), ev("C11", 3300), ev("H39", 3300)))
    assert grade(r) == CONFIRMED


def test_numbers_are_compared_as_numbers():
    """★ 3300 と 3300.0 は同じ値（型で割れると偽の『割』が出る）。

    ★ ただしこの 1 本だけでは `round()` を叩いていない ── Python では
      `3300 == 3300.0` が最初から真なので、丸めを外しても緑のままになる
      （変異試験で素通りして分かった）。下の小数の試験と対で意味を持つ。
    """
    r = rec("請求額", (ev("C11", 3300), ev("H39", 3300.0)))
    assert grade(r) == CONFIRMED


def test_tiny_float_noise_does_not_split_a_matching_pair():
    """★ 小数の揺れで偽の『割』を出さない。

    ★ 実物の消費税は `ROUND(x, 1)` で小数が出る（検体の実測）。
      浮動小数の計算で末尾がわずかにずれた 2 口を『食い違い』と言うと、
      **正しい冊に印が立つ**（誤報 ── 黙って失敗する次に悪い）。
    """
    r = rec("請求額", (ev("C11", 3300.001), ev("H39", 3300.004)))
    assert grade(r) == CONFIRMED, "★ 小数の末尾のずれで割になった（誤報）"


def test_a_real_difference_in_the_decimals_still_splits():
    """★ 逆向き ── 丸めが緩すぎて本物の差を見逃さないこと（負の被覆）。"""
    r = rec("請求額", (ev("C11", 3300.0), ev("H39", 3300.5)),
            blank_reason="0.5 の差があります")
    assert grade(r) == SPLIT, "★ 0.5 の差を丸めて『確』にした（見逃し）"


def test_text_is_compared_after_trimming():
    r = rec("請求元", (ev("G5", "あかね商事"), ev("B3", " あかね商事 ")))
    assert grade(r) == CONFIRMED


# ── ★ ③一致だけでは 確 と言わない ────────────────────────────
def test_agreement_without_a_sweep_is_only_single():
    """★ 掃き出していないなら、一致していても言えるのは 単 まで。

    ★ 2026-09-11 の実測: 上部の請求額欄と帯の合計が一致した 12 冊で、
      **開いていない第三の口**（源泉徴収の振込金額・備考の再掲・明細の合計・
      2 枚目のシート・通貨・帳票種別）が食い違っていた。
      一致は「他が黙っている」ことではない ── **開いていないだけ**。
    """
    r = Record("請求額", (ev("C11", 3300), ev("H39", 3300)))   # swept 既定は False
    assert grade(r) == SINGLE, (
        "★ 掃き出していないのに『確』と言った ── 見ていないから自信がある形")
    assert value(r) == 3300, "★ 単 なので値は出す（黙るのではない）"


def test_the_reason_says_the_sweep_was_not_done():
    """★ 一致しているのに 単 なら、**なぜ裏が取れていないか**を人に言う。"""
    r = Record("請求額", (ev("C11", 3300), ev("H39", 3300)))
    text = describe(r)
    assert "確かめられていません" in text, f"★ 黙って単にしている: {text}"
    assert "C11" in text and "H39" in text, f"★ 一致した口を示していない: {text}"


def test_a_contradiction_found_elsewhere_forces_a_blank():
    """★ 値そのものは一致していても、別の所で食い違いを見つけたら値は出さない。

    ★ 2026-09-11 の実測: 明細の合計と小計が合わない冊で、**食い違いを見つけて
      いるのに区分は「単」**という記録が出ていた。理由の文字列にだけ書いて、
      導出は `evidences` しか見ていなかった（片配線）。
    """
    r = Record("請求額", (ev("C11", 3300), ev("H39", 3300)),
               blank_reason="明細の合計と小計が合いません",
               conflict=True, conflict_why="明細の合計と小計が合いません")
    assert grade(r) == SPLIT, "★ 食い違いが区分に届いていない（片配線）"
    assert value(r) is None


def test_grade_of_is_the_only_derivation():
    """★ `grade()` は `grade_of()` に丸投げしていること（偽造で先読みさせない）。"""
    evs = (ev("C11", 3300), ev("H39", 3300))
    assert grade_of(evs, swept=True) == CONFIRMED
    assert grade_of(evs, swept=False) == SINGLE
    assert grade_of(evs, swept=True, conflict=True) == SPLIT
    assert grade_of((), swept=True) == NONE_FOUND


# ── ★ ②空欄には理由が要る ──────────────────────────────────
def test_a_blank_cannot_be_built_without_a_reason():
    """★ 型で守る ── 理由を持たない空欄の記録は組み立てられない。

    ★「後から理由を足す」形にすると、片方の経路だけ理由なしで空欄を書く事故が生える。
    """
    with pytest.raises(ValueError, match="理由が無い"):
        Record("請求元", ())                      # 無 なのに理由が無い
    with pytest.raises(ValueError, match="理由が無い"):
        Record("請求額", (ev("C11", 100), ev("H39", 200)))   # 割 なのに理由が無い
    with pytest.raises(ValueError, match="理由が無い"):
        # ★ 食い違いを立てたのに理由が無い経路も塞ぐ
        Record("請求額", (ev("C11", 100),), conflict=True)


def test_a_value_bearing_record_needs_no_reason():
    """確 と 単 は理由が要らない（空欄ではないので）。"""
    rec("請求額", (ev("C11", 3300), ev("H39", 3300)))
    rec("請求額", (ev("C11", 3300),))


# ── 両側の数字と、人に見せる 1 行 ──────────────────────────
def test_split_shows_both_sides():
    """★ 買い手の信用条件④「両側の数字を並べる」。"""
    r = rec("請求額", (ev("C11", 3300, how="上部ラベルの右"),
                       ev("H39", 3301, how="帯の合計")),
            blank_reason="1 円ずれています")
    sides = both_sides(r)
    assert {(at, v) for at, v, _ in sides} == {("C11", 3300), ("H39", 3301)}


def test_describe_says_something_specific_for_every_grade():
    """★ 説明が定数文字列でないこと ── 区分ごとに違い、番地を含む。"""
    recs = [
        rec("請求額", (ev("C11", 3300), ev("H39", 3300))),
        rec("請求額", (ev("C11", 3300),)),
        rec("請求額", (ev("C11", 3300), ev("H39", 3301)), blank_reason="ずれ"),
        rec("請求元", (), blank_reason="ラベルが無い"),
    ]
    texts = [describe(r) for r in recs]
    assert len(set(texts)) == 4, f"★ 説明が区分ごとに違わない: {texts}"
    for t in texts[:3]:
        assert any(c.isdigit() for c in t), f"★ 番地も値も含まない説明: {t}"
