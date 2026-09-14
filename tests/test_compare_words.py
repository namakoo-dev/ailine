# -*- coding: utf-8 -*-
# 辞書を「正確に引く」機構の番人（2026-09-15・設計と凍結した予測は docs/DESIGN-20260915-辞書を正確に引く.md）。
#
# ★ なぜ在るか: 言い回し 120 件の盲検で辞書に口語を足したあと、Namakoo「辞書登録後は正確に
#   引いてこれる仕組みも必要だ」。引き方を実測したら欠陥が 4 つ、どれも**辞書を増やすほど悪化**:
#     「20時間を超えない人」→ gt（逆）／「3000以上5000未満」→ gte（未満を黙って捨てる）／
#     「10を切っていない品番」→ None → LLM の不等号がそのまま通る
#
# 番人は 4 つ（辞書が育っても壊れない形）:
#   1. 1 語 1 検体 ── 全レコードの陽性を機械で回す（検体の無い語は入らない）
#   2. 到達できない語は入らない ── 自分の検体で**その語が**勝つ
#   3. 語彙の共食いの凍結 ── 408 句の読み（bench/compare_words_freeze.json）が 1 件も動かない
#   4. 出所の無い語は入らない
import json
import os
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import compare_words as C  # noqa: E402
sys.path.insert(0, str(Path(__file__).resolve().parent))
from _product_source import count_in_product  # noqa: E402 ── ★ 本体を場所で決め打ちしない

FREEZE = REPO / "bench" / "compare_words_freeze.json"
REGEN = os.environ.get("AILINE_REGEN_COMPARE_FREEZE") == "1"


# ── 辞書そのものの番人 ────────────────────────────────────────────────────────

def test_every_word_has_a_specimen_and_wins_it():
    """1・2: 全レコードに陽性 1 件が在り、その検体で**その語が**勝つ（他の語に覆われていない）。"""
    bad = []
    for w in C.WORDS:
        if not w.example:
            bad.append((w.text, "検体が無い"))
            continue
        r = C.read(w.example)
        if r.cmp != w.cmp or r.word != w.text:
            bad.append((w.text, f"自分の検体で勝てない: {r}"))
    assert not bad, bad


def test_every_word_has_a_source_and_a_known_guard():
    """4: 出所の無い語は入らない（頭から思いついた語を足さない）。ガードは 3 種だけ。"""
    assert all(w.source for w in C.WORDS), [w.text for w in C.WORDS if not w.source]
    assert all(w.guard in C.GUARDS for w in C.WORDS), [w.text for w in C.WORDS if w.guard not in C.GUARDS]
    assert len({w.text for w in C.WORDS}) == len(C.WORDS), "同じ字面のレコードが 2 つある"


# ── 引き方の番人（凍結した予測・5/5）────────────────────────────────────────────

def test_longest_match_wins_at_the_same_place(monkeypatch):
    """同じ位置から始まる 2 語は**長い方が勝つ**（「超え」が「超えない」を横取りしない）。

    ★ 正直に書く: 出荷している辞書には同じ位置から始まる 2 語が**無い**ので、この規則は今日の
      辞書では発火しない（変異試験で緑のままだった＝在っても鳴らない番人）。規則が要るのは
      **否定形を語として登録する日**（反転を定義する日）── だからその日の形の辞書を検体で作る。
    """
    extra = C.Word("超えない", "lte", C.NUM_BEFORE, source="検体用（この試験の中だけ）",
                   example="金額が5000を超えない行")
    monkeypatch.setattr(C, "WORDS", C.WORDS + (extra,))
    r = C.read("金額が5000を超えない行を抜き出して")
    assert (r.cmp, r.word, r.negated) == ("lte", "超えない", False), r
    # ★ 出荷の辞書では「を超える」（先に始まる）が「超え」（重なる）を畳む ── 語は 1 つだけ報告される
    r2 = C.read("金額が5000を超える行を抜き出して")
    assert (r2.cmp, r2.word) == ("gt", "を超える"), r2


@pytest.mark.parametrize("task,inverted", [
    ("残業時間が20時間を超えない人を抜き出して", "以下"),
    ("在庫数が10個を下回らない品番を抜き出して", "以上"),
    ("金額が5000以上ではない行", "未満"),
])
def test_a_negated_comparison_is_asked_back_not_flipped(task, inverted):
    """① 否定は**反転させず聞き返す** ── 黙って反転させるのは直している事故と同じ形。"""
    r = C.read(task)
    assert r.cmp is None and r.hit and r.negated, r
    assert inverted in r.ambiguous and "否定" in r.ambiguous, r.ambiguous


def test_two_different_comparisons_are_refused_as_a_range():
    """② 『以上』と『未満』の両方 ── 黙って片方を捨てない。"""
    r = C.read("金額が3000以上5000未満の行を抜き出して")
    assert r.cmp is None and r.hit and not r.negated, r
    assert "以上" in r.ambiguous and "未満" in r.ambiguous and "範囲" in r.ambiguous, r.ambiguous


@pytest.mark.parametrize("task,want,word", [
    ("金額が1000以上で部門が営業と同じ行の備考に「○」を付けて", "gte", "以上"),
    ("部門が営業と同じで金額が1000以上の行", "gte", "以上"),       # 並びに依らず数値の側
])
def test_a_numeric_comparison_beside_an_equality_is_two_conditions_not_a_range(task, want, word):
    """★ 曖昧と読むのは**数値の比較が 2 つ**（範囲）だけ ── 数値＋等しい は 2 条件の依頼で、
       比較は数値の側（2 組目は実表の値で読む既存の道・test_two_conditions_are_not_half_done）。"""
    r = C.read(task)
    assert (r.cmp, r.word, r.ambiguous) == (want, word, ""), r


def test_two_non_numeric_words_keep_the_old_first_match_until_measured():
    """保留（実文が無い）: 「を含む」と「と同じ」が並ぶ回は旧来どおり先に出た語。
       ★ 発火条件: 実文が 1 件出た日に、範囲と同じ線で断るかを測ってから決める。"""
    r = C.read("備考に東京を含む行で状態が完了と同じもの")
    assert (r.cmp, r.word) == ("contains", "を含む"), r


def test_the_same_comparison_twice_is_not_ambiguous():
    r = C.read("金額が5000以上の行を、以上で抜き出して")
    assert r.cmp == "gte" and not r.ambiguous, r


def test_no_match_reads_nothing_and_says_so():
    r = C.read("在庫数が10を切っていない品番")     # ★ 辞書の漏れ ── 引けない回は決めない
    assert r == C.Reading(), r
    assert C.read("") == C.Reading() and C.read(None) == C.Reading()


@pytest.mark.parametrize("task,want", [
    ("単価が1000より高い行", "gt"),
    ("20件を超えた分だけ抜き出して", "gt"),
    ("金額が5000以上の行", "gte"),
])
def test_clean_readings_are_unchanged(task, want):
    assert ailine.extract_cmp_from_task(task) == want


# ── 3: 語彙の共食いの凍結（408 句）─────────────────────────────────────────────

def _now() -> list:
    frozen = json.loads(FREEZE.read_text(encoding="utf-8"))
    out = []
    for row in frozen["readings"]:
        r = C.read(row["phrase"])
        out.append({**row, "cmp": r.cmp, "word": r.word, "hit": r.hit,
                    "negated": r.negated, "ambiguous": bool(r.ambiguous)})
    return frozen, out


def test_the_frozen_readings_of_408_phrasings_do_not_move():
    """★ 辞書を変えた前後で、凍結した句の読みが**1 件も動かない**。動くなら意図した件だけのはずで、
       それは AILINE_REGEN_COMPARE_FREEZE=1 で再生成し、git diff を人が読んでから commit する
       （公開面の凍結と同じ作法）。★ 雑音床 0（純関数なので LLM の揺れは無い）。"""
    frozen, now = _now()
    if REGEN:
        FREEZE.write_text(json.dumps({**frozen, "readings": now}, ensure_ascii=False, indent=1) + "\n",
                          encoding="utf-8")
        pytest.skip("凍結を再生成した ── git diff を読むこと")
    moved = [(a["phrase"], {k: a[k] for k in ("cmp", "word", "hit", "negated", "ambiguous")},
              {k: b[k] for k in ("cmp", "word", "hit", "negated", "ambiguous")})
             for a, b in zip(frozen["readings"], now) if a != b]
    assert not moved, f"{len(moved)} 件の読みが動いた（意図した件なら凍結を再生成）: {moved[:5]}"
    assert len(frozen["readings"]) >= 400, "凍結の句が減っている"


# ── 呼び出し側の規則（EXTRACT / SET_WHERE / フォルダ抽出が同じ器官を同じ規則で呼ぶ）──────

def _meta():
    return {"sheets": ["Sheet"], "headers": {"Sheet": ["商品", "金額", "印"]},
            "header_rows": {"Sheet": 1}}


def test_extract_refuses_a_negated_comparison_instead_of_flipping():
    ok, r, _i, err = ailine.verify_dsl_args(
        "EXTRACT", {"col": "金額", "cmp": "gt", "value": 5000}, _meta(),
        task="金額が5000を超えない行を抜き出して")
    assert not ok and "否定" in err and "以下" in err, (ok, err)


def test_extract_refuses_a_range_instead_of_keeping_the_first_word():
    ok, r, _i, err = ailine.verify_dsl_args(
        "EXTRACT", {"col": "金額", "cmp": "gte", "value": 3000}, _meta(),
        task="金額が3000以上5000未満の行を抜き出して")
    assert not ok and "範囲" in err, (ok, err)


def test_extract_does_not_take_the_llm_inequality_when_the_dictionary_is_silent():
    """④ 引けなかった回に黙って LLM に負けない ── 三項（依頼／宣言／実体）の依頼側が欠けている。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "EXTRACT", {"col": "金額", "cmp": "lt", "value": 10}, _meta(),
        task="金額が10を切っていない行を抜き出して")
    assert not ok and "比較の語" in err and "未満" in err, (ok, err)


def test_extract_without_a_task_still_takes_the_dsl_as_is():
    """★ 依頼文が無い経路（DSL 直渡し・ゴールデン）は触らない ── 接地する相手が無い。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "EXTRACT", {"col": "金額", "cmp": "lt", "value": 10}, _meta(), task="")
    assert ok and r["cmp"] == "lt" and r["value"] == 10.0, (ok, err, r)


def test_extract_eq_with_a_number_is_not_touched_by_the_rule():
    """規則 ④ は**数値の比較**だけ ── eq は辞書に無くても通る（名指しの抽出の道）。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "EXTRACT", {"col": "金額", "cmp": "eq", "value": 5000}, _meta(),
        task="金額が5000の行を抜き出して")
    assert ok and r["cmp"] == "eq", (ok, err, r)


def test_set_where_refuses_a_negated_comparison():
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "印", "cond_col": "金額", "cmp": "gt", "cond_value": 5000, "value": "◎"},
        _meta(), task="金額が5000を超えない行の印に『◎』を付けて")
    assert not ok and "否定" in err, (ok, err)


def test_set_where_does_not_take_the_llm_inequality_when_the_dictionary_is_silent():
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "印", "cond_col": "金額", "cmp": "gte", "cond_value": 5000, "value": "◎"},
        _meta(), task="金額が5000の行の印に『◎』を付けて")
    assert not ok and "比較の語" in err, (ok, err)


def test_set_where_with_a_dictionary_word_is_unchanged():
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "印", "cond_col": "金額", "cmp": "gte", "cond_value": 5000, "value": "◎"},
        _meta(), task="金額が5000以上の行の印に『◎』を付けて")
    assert ok and r["cmp"] == "gte" and r["cond_value"] == 5000.0, (ok, err, r)


def test_the_three_deciders_share_one_organ():
    """★ 片配線の番人: 比較を決める 3 経路（EXTRACT／SET_WHERE／フォルダ抽出）が全部
       compare_words.read を呼び、旧辞書（_EXTRACT_CMP_WORDS）が本体に残っていない。"""
    assert count_in_product("compare_words.read(") >= 7, count_in_product("compare_words.read(")
    assert count_in_product("_EXTRACT_CMP_WORDS") == 0, "旧辞書が本体に残っている（辞書が 2 つになる）"
    assert count_in_product("compare_words.unconfirmed(") == 3, "規則 ④ が 3 経路に配線されていない"
    assert count_in_product("if _cmp_read.ambiguous:") == 3, "否定・曖昧の断りが 3 経路に配線されていない"
