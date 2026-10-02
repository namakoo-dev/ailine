# -*- coding: utf-8 -*-
"""断定の文は、畳んだ関数の外に増えない ── 増えたら理由を書くまで赤（2026-10-02・形 7 の番人）。

★★ 前提: 盲検の欠陥の形 7（断り・警告の文言が偽）は、件ごとに直しても、**次に書く断定の文**が
  同じ形で嘘になる。書いた時点では嘘になる場面が見えないので、人の注意では止まらない。
  ★ だから「断定の語」を含む文を、書かれた場所から機械で導き（tests/claim_wording_core.py）、
    理由つきの台帳（tests/claim_wording_register.json）と**等号で**縛る:
      ・台帳に無い文が現れたら赤（新しい断定の文は、言う根拠を台帳に書くまで通らない）
      ・台帳に在って実装に無い文も赤（死んだ免除を残さない）
      ・件数が違っても赤

★ この試験自身を疑う（陽性対照・変異）:
    ① 直す前の文（HEAD にあった形）を実際に与えて、語が**鳴る**こと
    ② 直した後の形（探したシートを言う口が頭に付く）では**鳴らない**こと
    ③ 断定の文を 1 つ足した一時ファイルで、台帳との比較が**赤になる**こと
"""
from __future__ import annotations

import functools
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import claim_wording_core as core  # noqa: E402


@functools.lru_cache(maxsize=1)
def _real_tally() -> dict:
    """製品コード全体の走査（重いので 1 回だけ）。"""
    return core.tally(core.scan())


# --- 本体 -------------------------------------------------------------------------------------

def test_every_assertive_sentence_is_registered_with_its_reason():
    diff = core.compare(_real_tally(), core.load_register())
    assert not diff["unlisted"], (
        "断定の語を含む文が台帳に無い（新しい文は、**何を観測して／どの宣言表に当たって**言っているかを"
        "tests/claim_wording_register.json に書くこと。書けないなら文を直す ── 見たことと推し量りを分ける）: "
        f"{diff['unlisted']}")
    assert not diff["wrong_count"], f"件数が台帳と違う（台帳, 実装）: {diff['wrong_count']}"
    assert not diff["stale"], f"台帳に在って実装に無い文（死んだ免除）: {diff['stale']}"


def test_every_register_entry_has_a_reason_and_an_unlock():
    data = json.loads(core.REGISTER.read_bytes().decode("utf-8"))
    for e in data["allowed"]:
        assert len(e.get("reason", "")) >= 20, e["func"]
        assert len(e.get("unlock", "")) >= 6, e["func"]
        assert e["word"] in core.WORDS, e
        assert isinstance(e["count"], int) and e["count"] >= 1, e


# --- ① 陽性対照: 直す前の文に語が鳴る ---------------------------------------------------------------

#: HEAD にあった（直す前の）形。語ごとに 1 つ以上。
BEFORE = {
    "closed_two_way": "この道具が書いた記録がありません（人が置いたファイルか、途中で失敗した run が残した作業結果のどちらかです）",
    "after_change": "（{out.name} は ailine が作った物ですが、そのあと変更されています。",
    "placed_by_human": "（人が置いたファイルか、途中で失敗した run が残した作業結果のどちらかです）",
    "mixup_or": "── 列の取り違えか、別のソフトの書き出しが混ざっています",
    "not_supported": "この道具は『文字色』に対応していません",
    "not_supported_now": "列の値をそのまま書き換える操作は今のところ対応していません。",
    "no_such_op": "この道具に『{label}』を打ち消す操作はありません",
    "tool_lacks": "曜日や「今日から見て」のような読み方はこの道具にはありません。",
    "not_a_value": "『{v}』は列の名前です（書き込む値ではありません）",
    "absent_or_unchanged": "★ 依頼で言及された『列B』は存在しません/変更されていません",
    "scopeless_col": "列『{name}』がこの表にありません（ある列: A）",
    "scopeless_row": "『{name}』という行が見つかりません",
}

#: 位置の語の言い方（『X』という列がありません）も鳴ること（直す前の形・anchor.py）。
BEFORE_EXTRA = ["『{name}』という列がありません（ある列: A、B）"]

#: 直した後の形（探したシートを言う句が同じ文に載る）。★ 範囲を言わない不在としては鳴らないこと。
AFTER = [
    "列『{s}』がありません{searched(sheet)}。ある列: {known}",
    "『{name}』という行が見つかりません{searched(sheet)}",
    "列『{wanted}』がこの表にありません（{where}ある列: {known}）",
    "『{wanted}』という列がありません（{where}ある列: {known}）",
]


def _fires(sentence: str, word: str) -> bool:
    return bool(core.WORDS[word].search(sentence))


@pytest.mark.parametrize("word", sorted(core.WORDS))
def test_each_word_fires_on_the_sentence_it_was_made_from(word):
    assert word in BEFORE, f"語 {word} に陽性対照の文が無い"
    assert _fires(BEFORE[word], word), (word, BEFORE[word])


@pytest.mark.parametrize("sentence", BEFORE_EXTRA)
def test_the_anchor_style_column_absence_fires_too(sentence):
    assert _fires(sentence, "scopeless_col"), sentence


@pytest.mark.parametrize("sentence", AFTER)
def test_the_fixed_forms_do_not_fire(sentence):
    """直した形は、同じ文に探した範囲が載っているので、範囲を言わない不在としては鳴らない
    （★ 語そのものは鳴る ── 範囲が載っているかを別に見ている）。"""
    assert any(_fires(sentence, w) for w in core._SCOPE_WORDS), "語が鳴らない（対照が死んでいる）"
    assert core.SCOPE_STATED.search(sentence), sentence


# --- ③ 変異: 断定の文を 1 つ足すと赤 ---------------------------------------------------------------

def _scan_source(tmp_path, source: str) -> list:
    f = tmp_path / "mutant.py"
    f.write_bytes(source.encode("utf-8"))
    return core.scan([f])


def test_mutation_adding_one_assertive_sentence_turns_the_comparison_red(tmp_path):
    base = dict(_real_tally())
    extra = _scan_source(tmp_path, 'def f():\n    print("この道具は『文字色』に対応していません")\n')
    assert extra, "走査が一時ファイルの断定の文を拾えない（番人が死んでいる）"
    merged = dict(base)
    for k, n in core.tally(extra).items():
        merged[k] = merged.get(k, 0) + n
    diff = core.compare(merged, core.load_register())
    assert diff["unlisted"], "断定の文を足したのに、台帳との比較が赤にならない"


def test_mutation_an_absence_without_its_scope_is_caught_and_with_it_is_not(tmp_path):
    bare = _scan_source(tmp_path, 'def f(x):\n    return f"列『{x}』がありません"\n')
    assert {r["word"] for r in bare} == {"scopeless_col"}, bare
    scoped = _scan_source(tmp_path, 'def f(x, s):\n    return f"列『{x}』がありません（探したシート: 『{s}』）"\n')
    assert scoped == [], scoped


def test_mutation_a_second_sentence_in_a_registered_function_changes_the_count(tmp_path):
    reg = core.load_register()
    (file, func, word), entry = next(iter(reg.items()))
    base = dict(_real_tally())
    base[(file, func, word)] = base.get((file, func, word), 0) + 1       # 同じ関数に 1 文足した体
    diff = core.compare(base, reg)
    assert diff["wrong_count"], "登録済みの関数に文を足しても件数が合わないと言わない"


def test_mutation_a_register_entry_without_a_sentence_is_stale():
    reg = dict(core.load_register())
    some_file = next(iter(reg))[0]
    ghost = (some_file, "no_such_function", "not_supported")
    reg[ghost] = {"count": 1}
    diff = core.compare(_real_tally(), reg)
    assert ghost in diff["stale"]


def test_a_docstring_or_a_comment_is_not_a_sentence_the_user_reads(tmp_path):
    found = _scan_source(
        tmp_path,
        'def f():\n    """この道具は『X』に対応していません（説明）"""\n'
        '    # この道具は『Y』に対応していません\n    return 1\n')
    assert found == [], found


def test_a_concatenated_sentence_is_read_as_one(tmp_path):
    """行をまたいで連結した文でも、1 つの文として語を拾う。"""
    found = _scan_source(
        tmp_path, 'def f(x):\n    return ("列の値をそのまま書き換える操作は"\n            f"今のところ{x}対応していません")\n')
    assert found == [] or all(x["word"] for x in found)
    found2 = _scan_source(
        tmp_path, 'def f():\n    return ("列の値をそのまま書き換える操作は"\n            "今のところ対応していません。")\n')
    assert {x["word"] for x in found2} >= {"not_supported_now"}, found2
