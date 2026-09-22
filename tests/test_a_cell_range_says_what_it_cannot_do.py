# -*- coding: utf-8 -*-
"""セル範囲（A1:C5）を頼まれたら、**何が言えないのか**を言う（2026-09-22）。

★★ 出所（Namakoo 自身の実測・`~/.ailine/history.jsonl` 18,024 run 中 7 件）:

    2026-09-18  ok=True  op=DRAW_BORDERS   「A1:C5 を「済」にして」

  **「済」と入れてほしいのに罫線を引いて成功と報告していた。**
  ★ この事故そのものは **09-18 に別の角度で塞がれている** ── 引用値の残差検査
    （「依頼が書くと言っている『済』が、実行した解釈のどこにも出ていません」）。
    ★ 語の長さを広げて直さず、依頼者が引用符で括った値を項として渡す、という処方。

★★ 残っていた穴: **断りが的外れ**だった。いまは止まるが、画面はこう言う ──

    ？ この依頼を 2 回読んだところ、読み方が分かれました
      読み方 1: 背景色   読み方 2: 一括書換        ← どちらも頼んだことでない

  ★ 断るのは正しい。だが**何が言えないのか**を言わないと、人は言い直せない。

★★ op は足さなかった（Namakoo 決裁「1 で実装して、必要になれば 2 番に切り替える」）:
  ・需要は 18,024 run 中 **2 run**（範囲に値を入れる依頼）。盲検 5 体からは **0 件**
  ・プロンプトに op を足す代償は実測済み ── OPS_DOC に 1 行で 245 件中 3 件落ちる
    （2026-09-21・帯は 241〜245）
  ★ 発火条件: 範囲の依頼が盲検で出るか、実利用で増えたら 2 番（op を足す）へ。

★ 設計:
  ・範囲を受け取れる op は **`OP_SCHEMA` の宣言から導く**（手で並べない）
  ・代わりの言い方は **`example_task_for` から取る**（実機の番人が毎回通ることを確かめている）
    ── 自分で作文しない。通らない例を示すのは、示さないより悪い。
  ・置き場は**共通の断りの描画口**（`render_refusal`）── 断りの経路は 1 つでないので、
    経路ごとに書き足さない。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from ailine_core.examples import example_task_for  # noqa: E402


# --- 範囲を見つける -------------------------------------------------------------------

@pytest.mark.parametrize("task, want", [
    ("A1:C5 を「済」にして", "A1:C5"),
    ("A1：C5 を「済」にして", "A1:C5"),          # ★ 全角コロン
    ("印刷範囲をA1:D5に設定して", "A1:D5"),
    ("AA10:AB20 を結合して", "AA10:AB20"),
])
def test_it_finds_the_range_the_person_named(task, want):
    assert ailine.cell_range_in_task(task) == want


@pytest.mark.parametrize("task", [
    "状態の列を「済」にして",
    "2行目の備考を「確認済」に書き換えて",
    "金額が1000以上の行の備考に「○」を付けて",
    "品名で並べ替えて",
])
def test_it_stays_quiet_when_no_range_was_named(task):
    """★★ 陰性対照 ── 範囲を言っていない依頼に口を出さない。

    ★ ここが緩むと、通る依頼の断り画面に毎回よけいな 2 行が付く（★ が毎回出る病）。
    """
    assert ailine.cell_range_in_task(task) is None
    assert ailine.cell_range_note(task) == []


# --- 何が言えないのかを言う -----------------------------------------------------------

def test_the_note_names_the_range_and_what_can_take_it():
    """★ 事故そのもの ── 頼まれた範囲を名指しし、範囲を扱える操作を言う。"""
    lines = ailine.cell_range_note("A1:C5 を「済」にして")
    assert lines, "範囲を名指しされたのに黙っている"
    text = chr(10).join(lines)
    assert "A1:C5" in text, text
    assert "セル結合" in text, f"範囲を扱える操作を言っていない: {text}"


def test_the_ops_that_take_a_range_are_derived_not_written():
    """★★ 範囲を扱える op を**宣言から導く**こと（手で並べない）。

    ★ 範囲を扱う op が増えた日に、この案内が静かに古くなるのを防ぐ。
    """
    takes = {op for op, sch in ailine.OP_SCHEMA.items() if "range" in (sch or ())}
    assert takes == {"MERGE"}, f"範囲を受け取る op が変わった: {sorted(takes)}"
    from _product_source import window_around
    seg = window_around("def cell_range_note", before=0, after=1400)
    assert "OP_SCHEMA" in seg, "宣言から導いていない"
    assert '"MERGE"' not in seg, "★ 範囲を扱える op を字面で書いている"


def test_the_alternatives_come_from_the_verified_examples():
    """★★ 「こう言えば通る」は**実機で確かめた例**から取ること（作文しない）。

    ★ `tests/test_examples_actually_work.py` が `example_task_for` の中身を毎回
      実機で打って確かめている ── そこから引けば、案内が嘘にならない。
    ★ この repo の結論:「道具が自分の示した例を自分で断っていた。
      導線が嘘なら、導線が無いより悪い」。
    """
    text = chr(10).join(ailine.cell_range_note("A1:C5 を「済」にして"))
    for op in ("SET_COLUMN_VALUE", "SET_CELL_VALUE"):
        ex = example_task_for(op)
        assert ex and ex in text, f"{op} の実測済みの例が案内に出ていない: {ex}"
    from _product_source import window_around
    seg = window_around("def cell_range_note", before=0, after=1400)
    assert "example_task_for" in seg, "例を作文している（実測済みの表から取ること）"


def test_the_note_is_shown_on_the_common_refusal():
    """★★ 断りの**共通の描画口**に載っていること。

    ★ 断りの経路は 1 つでない（読み方が割れた回／対象の形式で止まった回…）。
      経路ごとに書き足すと、次に増えた経路で黙る（この repo の片配線そのもの）。
    """
    lines = ailine.render_refusal(
        "FILL_COLOR", {"target": "A1:C5", "color": "green"},
        "対象『A1:C5』の形式が不明です", task="A1:C5 を「済」にして")
    text = chr(10).join(lines)
    assert "セル範囲" in text, f"共通の断りに案内が出ていない: {text}"


def test_the_refusal_stays_clean_for_an_ordinary_request():
    """★ 陰性対照 ── 範囲を言っていない断りに、この案内は付かない。"""
    lines = ailine.render_refusal(
        "SORT", {"col": "存在しない", "order": "asc"}, "列『存在しない』がありません",
        task="存在しない列で並べ替えて")
    assert not [ln for ln in lines if "セル範囲" in ln], lines


def test_the_screen_is_not_markdown():
    """★ 画面に強調記号をそのまま出さないこと（端末は Markdown を解さない）。"""
    text = chr(10).join(ailine.cell_range_note("A1:C5 を「済」にして"))
    assert "**" not in text, f"強調記号が画面に出ている: {text}"
