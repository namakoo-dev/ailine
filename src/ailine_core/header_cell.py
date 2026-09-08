# -*- coding: utf-8 -*-
"""「<列名>の**見出し**に」と言われたら、列ぜんぶでなく**その 1 セル**。

★★ なぜ在るか（2026-09-08・盲検 B が false ✓ として拾った）:

    依頼   「金額の**見出し**に色を付けて太字にして」
    宣言   操作:背景色 対象:col:金額 ／ 操作:太字 対象:col:金額
    実物   E1〜E8（データ行も合計式の行も）が黄色＋太字
    出力   **✓ 機械検証済み**

  ★ 依頼の「見出し」が宣言のどこにも出ていない ── 三項の「依頼」が落ちた形。
    頼んでいない範囲へ静かに広がるのは、この道具がいちばん嫌う壊れ方。

★★ 08-27 の判断は**反転しない**: 「見出しを太字にして」（列を言わない）は
  **見出し行ぜんぶ**の意味でもありうるので曖昧 ── 勝手に狭めない。
  ★ 分かれ目は**列が名指しされているか**。列の見出しは 1 つしかないので曖昧でない。

      「見出しを太字に」        → 狭めない（曖昧・08-27 の判断のまま）
      「金額の見出しに色を」    → 1 セル（列が決まれば見出しは 1 つ）

★ 実在の列名だけを見る（依頼文から名前を切り出さない ── A' 原則）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations

import re

#: 「見出し」を指す語。★ ailine_core/subject.py と同じ語彙（片方だけ増えないよう、
#: 増やす時は両方を見ること）。
HEADER_WORDS = ("見出し", "ヘッダー", "ヘッダ")

#: 列名と見出し語の間に挟まってよいもの（助詞・空白）。
_BETWEEN = r"[のな\s　]{0,3}"


def names_a_column_header(task: str, column: str) -> bool:
    """依頼文が「<その列>の見出し」と言っているか。

    ★ 列名が**見出し語の直前**に在ることまで見る ── 「見出し」が文のどこかに在るだけで
      狭めると、「見出しを太字にして、金額の列に色を付けて」のような回で
      金額の列まで 1 セルに縮む（頼んだ範囲を勝手に狭めるのも事故）。
    """
    if not task or not column:
        return False
    pat = re.escape(column) + _BETWEEN + "(?:" + "|".join(HEADER_WORDS) + ")"
    return bool(re.search(pat, task))


def header_cell_target(task: str, target: str, header_row: int, columns) -> str | None:
    """`col:X` を `cell:見出し行,X の列番号` へ絞るべきなら、その新しい target を返す。

    絞らない回は None（呼び出し側は今までどおり）。
    """
    if not target.startswith("col:"):
        return None
    name = target[4:]
    names = [str(c) for c in (columns or [])]
    if name not in names or not names_a_column_header(task, name):
        return None
    return f"cell:{int(header_row)},{names.index(name) + 1}"
