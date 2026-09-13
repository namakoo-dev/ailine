# -*- coding: utf-8 -*-
"""画面の主役（4 つの区分）が、買い手の読む文書に在ること（2026-09-13・買い手の初見 C2）。

★★ なぜ要るか: `ailine forms` の画面と出力の主役は 4 つの区分なのに、README にも
  ENGINEERING にも**その語が 1 度も出てこなかった**（grep 0 件）。買い手は画面で初めて
  見る語で「値が空なのは壊れているのか」を判断できない。★ 出るのに説明が無いのは、
  出ないのと同じくらい危ない（人は自分の理解の方を疑う）。

★ 書き写しにしない: 凡例の 1 行は `field_record.grade_legend()` が作るものと**同一**か
  を機械で突き合わせる（意味を直した日に、文書だけ古く残らない）。
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core import field_record                                    # noqa: E402

README = REPO / "README.md"


def test_the_legend_in_the_readme_is_the_one_the_machine_prints():
    text = README.read_text(encoding="utf-8")
    legend = field_record.grade_legend()
    assert legend in text, (
        f"README に凡例が無い（機械が出すのは『{legend}』）── "
        "画面に出る語は、買い手が読む文書にも在ること")


def test_every_grade_word_is_explained_with_what_the_cell_shows():
    """★ 分母つき ── 4 語すべてが、セルがどうなるか（値／空欄）と一緒に説明されている。"""
    text = README.read_text(encoding="utf-8")
    rows = [ln for ln in text.splitlines() if ln.startswith("| ") and "|" in ln[2:]]
    explained = []
    for grade in field_record.GRADE_ORDER:
        hit = [ln for ln in rows if ln.startswith(f"| {grade} |")]
        assert len(hit) == 1, f"区分『{grade}』の説明行が {len(hit)} 本ある"
        assert "値を出す" in hit[0] or "空欄" in hit[0], f"セルがどうなるか書いていない: {hit[0]}"
        explained.append(grade)
    assert len(explained) == len(field_record.GRADE_ORDER) == 4, explained
