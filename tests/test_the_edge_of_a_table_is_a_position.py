# -*- coding: utf-8 -*-
"""「一番下」「最後」は**位置**であって行の名前ではない（2026-09-09）。

★★ 出所（Namakoo「言い間違えを除けば位置語の指定が精度に直結する」）:
  位置の語彙を軸ごとに並べたら、**列にだけ端が在り行には無かった**。

      概念        列                      行
      隣（後）    6 語                    5 語
      隣（前）    5 語                    3 語
      2 つの間    _re_between（共有）     同左        ← ここだけ対称
      端          _COL_HEAD/_COL_TAIL     ★ 無い

  ★ 09-08 に列側へ端を足したので、非対称が**広がって**いた。
  ★ 行と列の非対称は、この repo で 3 度目（間・端・…）── 列と**同じ形**で持つ。

★★ 実測で分かった本物の欠陥: LLM は位置を**語**で返すことがある（挿入位置=「最後」）。
  旧版はそれを**行の名前**として実表に探しに行き、
  「『最後』という行が見つかりません」と断っていた ── 位置語を値として扱っていた形。

★ 住所を解く所は **2 つ**あるので、両方に同じ語彙を持たせる（片方だけだと
  「一番下に足して」は通るのに「最後の行を削除して」が断られる ── 実測でそうなった）。
"""
from __future__ import annotations

import openpyxl
import pytest

import ailine

ROWS = [["商品", "分類", "売上"],
        ["りんご", "果物", 1200],
        ["みかん", "果物", 800],
        ["にんじん", "野菜", 1500],
        ["だいこん", "野菜", 600]]


@pytest.fixture()
def meta(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    return {"sheets": ["売上"], "headers": {"売上": list(ROWS[0])},
            "header_rows": {"売上": 1}, "path": str(p)}


# --- ① 語彙が軸で対称になっていること ---------------------------------------

def test_both_axes_have_the_same_position_concepts():
    """★ 概念の欠けを機械で見る ── 片方の軸にだけ足すと、また非対称が広がる。"""
    for name in ("_ANCHOR_AFTER", "_ANCHOR_BEFORE", "_ANCHOR_TOP", "_ANCHOR_BOTTOM",
                 "_COL_AFTER", "_COL_BEFORE", "_COL_HEAD", "_COL_TAIL"):
        got = getattr(ailine, name, None)
        assert got, f"{name} が無い（位置の概念が軸で欠けている）"
        assert all(isinstance(w, str) and w for w in got), (name, got)


# --- ② 端の語が「位置」として解けること -------------------------------------

@pytest.mark.parametrize("task, want_row, why", [
    ("一番下に「もも」の行を入れて", 6, "表の終わりの次"),
    ("一番上に「もも」の行を足して", 2, "見出しの次"),
    ("最後に1行足して", 6, "表の終わりの次"),
])
def test_an_edge_word_becomes_a_row_number(meta, task, want_row, why):
    got, note = ailine.resolve_row_anchor(task, meta, "売上", 1)
    assert got == want_row, (got, note)
    assert why in (note or ""), note


def test_a_named_row_still_wins_over_an_edge_word(meta):
    """★ 端は**隣より後**に見る（両方書いてあったら具体的な方を採る・列側と同じ順）。"""
    got, note = ailine.resolve_row_anchor("みかんの下に「もも」を足して", meta, "売上", 1)
    assert got == 4, (got, note)
    assert "みかん" in (note or ""), note


def test_an_edge_word_is_not_looked_up_as_a_row_name(meta):
    """★★ 実測した欠陥そのもの ── LLM が返した『最後』を行の名前として探していた。"""
    got, note = ailine._resolve_named_row(meta, "売上", "最後")
    assert got is not None, note
    assert "見つかりません" not in (note or ""), note


def test_a_real_row_name_is_unaffected(meta):
    """★ 対の試験 ── 表に在る名前は今までどおり行として解ける。"""
    got, note = ailine._resolve_named_row(meta, "売上", "みかん")
    assert got == 3, (got, note)


def test_both_address_resolvers_know_the_edge_words():
    """★★ 住所を解く所は 2 つ。片方だけに持たせると、実測どおり片側が断る。

    ★ 配線の形を静的に縛る（実機を起こさずに片配線を止める）。
    """
    import pathlib
    src = pathlib.Path(ailine.__file__).read_text(encoding="utf-8")
    uses = src.count("_ANCHOR_TOP + _ANCHOR_BOTTOM")
    assert uses >= 2, f"端の語を見ている住所の解決が {uses} 箇所（2 つとも要る）"
