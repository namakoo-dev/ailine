# -*- coding: utf-8 -*-
"""この道具が書かない属性を頼まれたら、近い操作で代わりに実行しない。

★★ なぜ在るか（2026-09-09・語彙外の台が初回の実行で拾った）:

    依頼   「品名の**文字色**を赤にして」
    実行   操作:背景色 対象:col:品名 色:red
    出力   **✓ 機械検証済み**   ── 3 回とも同じ（運ではなく決定論的）

  既存の 2 つの関所はどちらも**原理的に**見えなかった:
    residue.unaccounted_request_words … 実在する**列名**しか見ない（『文字色』は列名でない）
    intent.op_effect_mismatch         … 文字色も背景色も同じ**効果の種類**（書式のみ）
  ★ 粒度が 1 段足りない ── 列でも効果でもなく **属性**を見る器官が要る。

★★ この試験のいちばん大事な仕事は「鳴ること」ではなく、
   **名簿が持っている能力を殺していないこと**を確かめる方だ。
   拒否の名簿は広すぎても『断りすぎ』で済むが、うっかり `太字` を入れれば
   正しい依頼が全部断られる。だから op の照合語彙との重なりを機械で見る。
"""
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine
from ailine_core.intent import WE_DO_NOT_DO, asked_for_what_we_do_not_do
from _product_source import count_in_product  # noqa: E402 ── ★ 本体決め打ちで読まない


def test_the_list_does_not_eat_an_ability_we_have():
    """★ 名簿の語が、実際に持っている op の照合語彙と重ならないこと。

    重なった瞬間、その op は**永久に断られる**（黙って能力が消える）。
    """
    ours = set()
    for meta in ailine.OP_META.values():
        for key in ("label", "without_the_word"):
            if isinstance(meta.get(key), str):
                ours.add(meta[key])
        for key in ("synonyms", "match_phrases", "pool_phrases", "requires_word"):
            ours.update(str(w) for w in (meta.get(key) or ()))
    ours.update(str(v) for v in ailine.OP_LABELS.values())

    collisions = []
    for word in WE_DO_NOT_DO:
        for mine in ours:
            if mine and (word in mine or mine in word):
                collisions.append((word, mine))
    assert not collisions, (
        "『できない属性』の名簿が、持っている op の語彙と重なっている ── "
        f"その op は今後ずっと断られる: {collisions}")


@pytest.mark.parametrize("task, want", [
    ("品名の文字色を赤にして", "文字の色"),
    ("フォント色を変えて", "文字の色"),
    ("見出しを斜体にして", "斜体"),
    ("品名に取り消し線を引いて", "取り消し線"),
    ("行の高さを30にして", "行の高さ"),
    ("文字サイズを大きくして", "文字の大きさ"),
    ("ウィンドウ枠を固定して", "ウィンドウ枠の固定"),
    ("印刷範囲をA1:D5に設定して", "印刷範囲"),
    ("用紙の向きを変えて", "用紙の向き"),
])
def test_we_name_the_attribute_we_cannot_write(task, want):
    assert asked_for_what_we_do_not_do(task) == want


@pytest.mark.parametrize("task", [
    "見出しに背景色を付けて", "見出しを太字にして", "けい線を引いて",
    "列幅を自動調整して", "単価に桁区切りを付けて", "単価で降順に並べ替えて",
    "元に戻して",              # ★『取り消し』を含むが『取り消し線』ではない
    "単価が3000以上の行を抽出して",   # ★ 二重語 ── 『フィルタ』は別名で持っている能力
    "部門ごとに金額を集計して",       # ★ 同上 ── 『グループ化』は集計で出来る
])
def test_we_stay_quiet_for_what_we_can_do(task):
    """★ 逆向き ── 出来ることで鳴ったら、それは能力を殺す番人になる。"""
    assert asked_for_what_we_do_not_do(task) is None


def test_a_real_column_by_that_name_is_a_target_not_an_attribute():
    """『書体』という列が実在する表なら、それは対象であって属性の指定ではない。"""
    assert asked_for_what_we_do_not_do("書体の列を太字にして", ["書体", "品名"]) is None
    assert asked_for_what_we_do_not_do("書体を変えて", ["品名"]) == "フォントの種類"


def test_the_judgement_lives_in_exactly_one_place():
    """★ 判定は 1 箇所 ── 呼び出し側に op ごとの if を配らない（片配線を作らない）。"""
    assert count_in_product("asked_for_what_we_do_not_do(") == 2, (
        "定義 1 + 呼び出し 1 のはず（増えていたら、判定が散っている）")


def test_the_gate_actually_stops_the_plan():
    """★ 文字列ではなく**振る舞い**を見る。

    ★★ なぜこの形なのか（2026-09-09・自分の番人が変異で赤くならなかった）:
      初版は「呼び出しの文字列が製品に在るか」しか見ていなかったので、
      `if False and 判定(...)` と殺しても数は変わらず**緑のまま**だった。
      ── 在っても鳴らない、を番人自身がやった。呼ぶのは実物の関所そのもの。
    """
    import argparse
    meta = {"headers": {"見積": ["品名", "単価"]}, "sheets": ["見積"]}
    fill = [{"op": "FILL_COLOR", "args": {"target": "col:品名", "color": "red"}}]

    _plan, rc = ailine._reread_the_plan(
        argparse.Namespace(task="品名の文字色を赤にして"), meta, fill)
    assert rc == 3, "書けない属性を頼まれたのに、計画がそのまま通った"

    _plan, rc = ailine._reread_the_plan(
        argparse.Namespace(task="品名に背景色を付けて"), meta, fill)
    assert rc is None, "出来る依頼まで止めている（能力を殺す番人になっている）"
