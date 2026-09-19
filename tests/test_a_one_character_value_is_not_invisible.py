# -*- coding: utf-8 -*-
"""1 文字の値も、行番号との矛盾として見る（2026-09-19）。

★★ 出所（2026-09-18 夜の ⚠ 家系の監査・Namakoo「漏れは塞ぐけど実際に欠陥とは限らないのか
  調べることはできる？」→ 8 つの盲点を全部調べた回に**新しく見つかった穴**）:

    表     2行目 記号=A / 3行目 記号=B
    依頼   「2行目の記号を B にして」     ← B は 3 行目に在る（依頼が自己矛盾）
    関所   黙る                          ← ★ len(v) >= 2 で 1 文字を見ていなかった

  ★ 今日『済』で塞いだのと**同じ 1 文字の盲点**。この repo に 3 箇所あり
    （残差ゲート／引用値／ここ）、これが最後の 1 つ。
  ★ ◎ ○ × 済 可 A B は帳簿でいちばん普通の値 ── 引用値 313 件中 68 件が 1 文字だった。

★★ なぜ `len(v) >= 2` が在ったか（外す前に理由を測った）:

    表の 3 行目に 値 '0' が在る状態で 「2行目の数量を**10**にして」
      → 素朴に 1 文字を許すと '0' が '10' の部分文字列として一致し、**誤爆**する

  ★ だから「1 文字を許す」のではなく **文字種の境界で切る**（A1 記法で使ったのと同じ形）:
      数字の値は前後が数字でない時だけ／英字の値は前後が英字でない時だけ 一致とみなす。
  ★ 記号（◎ ○ × 済）は境界を持たないので、そのまま一致する。

★ 測った限界（★ 実データに無い形なので**直さず開示する**）:
    「3行目の**C**コードを見て」のように英字 1 文字にカタカナが続く形は拾ってしまう。
    ★ 依頼文 1,942 件を調べて**この形は 0 件**だった ── 発明した検体に合わせて
      規則を曲げない（曲げると実在する『Bコース』のような値を落とす）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

from ailine_core.row_conflict import value_not_in_the_named_row  # noqa: E402

ROWS = [["取引先", "記号", "数量", "担当"],
        ["大東金属", "A", 10, "佐藤"],
        ["みどり商事", "B", 0, "鈴木"],
        ["丸山重工", "C", 5, "田中"]]


@pytest.fixture
def book(tmp_path):
    p = tmp_path / "在庫.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    wb.close()
    return p


# --- 事故そのもの ------------------------------------------------------------------------

@pytest.mark.parametrize("task, row_no, want", [
    ("2行目の記号をBにして", 2, "B"),          # ★ B は 3 行目（事故そのもの）
    ("2行目の記号を「B」にして", 2, "B"),      # ★ 引用つきでも同じ
    ("2行目のメモをCにして", 2, "C"),          # ★ C は 4 行目
    ("2行目の担当を鈴木にして", 2, "鈴木"),     # 2 文字以上（従来どおり）
])
def test_a_one_character_value_from_another_row_is_caught(book, task, row_no, want):
    assert value_not_in_the_named_row(task, row_no, book) == want


# --- 誤爆殺し（★ len(v) >= 2 が守っていたもの）----------------------------------------------

@pytest.mark.parametrize("task, row_no, why", [
    ("2行目の数量を10にして", 2, "★ '0'（3行目）が '10' の部分文字列として一致する形"),
    ("2行目の数量を15にして", 2, "★ '5'（4行目）が '15' の部分文字列"),
    ("2行目の取引先をAAA商事にして", 2, "★ 'A' は 2 行目自身の値・かつ 'AAA' の一部"),
    # ★★ 変異試験が名指しした穴（2026-09-19）: 上の 'A' は**2 行目自身の値**なので
    #   そもそも候補から外れており、**英字の境界を一度も試していなかった**。
    #   'C' は 4 行目の値なので候補に残る ── ここが英字側の本物の陰性対照。
    ("2行目のコードをCCにして", 2, "★ 'C'（4行目）が 'CC' の一部として一致する形"),
    ("2行目の記号をABCにして", 2, "★ 'B'（3行目）が 'ABC' の一部"),
    ("2行目の担当を佐藤にして", 2, "佐藤は 2 行目自身の値（矛盾していない）"),
])
def test_a_substring_match_does_not_ring(book, task, row_no, why):
    assert value_not_in_the_named_row(task, row_no, book) is None, why


def test_a_standalone_digit_still_rings(book):
    """★ 境界で切るだけ ── 単体で書かれた数字は従来どおり拾う（0 は 3 行目の値）。"""
    assert value_not_in_the_named_row("2行目の数量を0にして", 2, book) == "0"


# --- 既にある性質を壊していないこと --------------------------------------------------------

def test_the_longest_value_still_wins(tmp_path):
    """★ 「青りんご」と「りんご」が両方在る表で、短い方だけを拾わない（既存の不変）。"""
    p = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in (["商品", "棚"], ["みかん", "A"], ["青りんご", "B"], ["りんご", "C"]):
        ws.append(r)
    wb.save(p)
    wb.close()
    assert value_not_in_the_named_row("2行目の青りんごを消して", 2, p) == "青りんご"


def test_a_header_value_is_still_ignored(book):
    """★ 見出し行の語は値ではない（既存の絞り）── 『記号』は列名。"""
    assert value_not_in_the_named_row("2行目の記号を書き換えて", 2, book) is None


def test_the_known_limit_is_recorded(book):
    """★★ 開示: 英字 1 文字にカタカナが続く形は拾ってしまう（直さないと決めた）。

    ★ 実データ 1,942 件にこの形は 0 件。発明した検体に合わせて規則を曲げると、
      実在しうる『Bコース』のような値を落とす方が高くつく。
    ★ ここが将来変わったら（実データに出たら）、この試験が**気づかせる**。
    """
    assert value_not_in_the_named_row("2行目のCコードを見て", 2, book) == "C"
