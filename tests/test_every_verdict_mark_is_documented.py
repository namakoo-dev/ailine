# -*- coding: utf-8 -*-
"""画面に出る**判定の印**が、README の一覧に全部載っているか（2026-09-08）。

★★ 出所（盲検の検品が CONFUSING として挙げた）:

    README の一覧   `✓ / △ / ×` の 3 つ
    実際に出た      `⚠ ...に適用しましたが、機械保証はありません`  ← ★ 一覧に無い
                    `？ ...`（断り）                                ← 本文にはあるが一覧に無い

  「これは △ の仲間？失敗？」と初めての人が迷う ── **在るのに文書が知らない**形。

★ この番人は「印の語彙」だけを見る（文言までは縛らない）。新しい印を足したのに
  README を直し忘れたら赤くなる ── 直す場所を機械が指す。
★ 逆向きも縛る: README にだけ在って**実物が一度も出さない印**も赤にする
  （文書が実物より豪華になるのを止める ── 昨日 3 件直したのと同じ形）。
"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
README = ROOT / "README.md"

#: 判定の見出しに使う印。★ ここに足したら README にも足すことになる。
MARKS = "✓△⚠×？"

#: 印を出す実装（見出しを組む場所）。
SOURCES = (ROOT / "src" / "ailine" / "__init__.py",
           ROOT / "src" / "ailine_core" / "claim.py",
           ROOT / "src" / "ailine_core" / "cli_render.py")


def _documented() -> set:
    """README 冒頭の一覧（``` で囲んだ**最初の塊**）の各行の先頭文字。

    ★ 先頭文字を印の集合で絞ってはいけない ── 絞ると、載っていない印の行は
      塊の外に落ちて**逆向きの試験が恒真になる**（2026-09-08 の変異試験で実測。
      にせの印を足しても緑のままだった）。塊は柵（```）で切り、中身は全部読む。"""
    body = README.read_text(encoding="utf-8")
    block = re.search(r"```\n(.*?)```", body, re.S)
    assert block, "README 冒頭の判定一覧が見つからない（書式が変わった？）"
    return {ln[0] for ln in block.group(1).splitlines() if ln.strip()}


def _used() -> set:
    """実装が**行頭の印**として書いている文字（`"✓ ...` / `f"\n⚠ ...` の形）。"""
    pat = re.compile(r'(?:f?")(?:\n)?([' + MARKS + r'])\s')
    used = set()
    for src in SOURCES:
        used |= set(pat.findall(src.read_text(encoding="utf-8")))
    return used


def test_the_readme_lists_every_mark_the_tool_prints():
    used, doc = _used(), _documented()
    assert used - doc == set(), (
        f"画面に出るのに README の一覧に無い印: {sorted(used - doc)}")


def test_the_readme_does_not_list_a_mark_the_tool_never_prints():
    used, doc = _used(), _documented()
    assert doc - used == set(), (
        f"README の一覧に在るのに実装が出さない印: {sorted(doc - used)}")
