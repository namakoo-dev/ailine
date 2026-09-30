# -*- coding: utf-8 -*-
"""並べ替えの向き（昇順／降順）を**依頼文から**読む。

★★ なぜ在るか（2026-09-30・盲検 7 体目の致命）:

    依頼   「請求日の新しい順に並べ替えて」（表は古い順に並んでいる）
    解釈   順:昇順 ／（文書に変化は検出されなかった）／ ✓ 機械検証済み

  ★ 向き（order）を決めていたのは LLM だけで、依頼文と突き合わせる所が無かった。
    事後条件は「宣言どおりに並んだか」（宣言↔実体）しか見ないので、向きを取り違えても
    ✓ になる ── 依頼の項が欠けていた（判定には 依頼／宣言／実体 の 3 項が要る）。
  ★ 値は模型に決めさせず依頼文から literal で取る（区切り・書き込む値と同じ作法）。

★ 読めない時は None（推測しない）:
    - 語が 1 つも無い
    - 両向きの語が残る（「昇順・降順」）
  打ち消された語（「新しい順ではなく」）は数えない ──「新しい順ではなく古い順に」は昇順。
★ 語は**列挙で増やさない**。足す前に、足した語が何を奪うかを測ること。
★ ailine を import しない（持ち出せる部品）。
"""
from __future__ import annotations

import re

#: 語 → 向き。1 語 1 行。「請求日の新しい順」「金額が大きい順」は部分一致で拾う。
DIRECTION_WORDS = (
    ("降順", "desc"),
    ("大きい順", "desc"),
    ("高い順", "desc"),
    ("多い順", "desc"),
    ("新しい順", "desc"),
    ("最新順", "desc"),
    ("遅い順", "desc"),
    ("長い順", "desc"),
    ("昇順", "asc"),
    ("小さい順", "asc"),
    ("低い順", "asc"),
    ("安い順", "asc"),
    ("少ない順", "asc"),
    ("古い順", "asc"),
    ("早い順", "asc"),
    ("短い順", "asc"),
)

#: 語の直後に付く打ち消し（「新しい順ではなく」「新しい順じゃなくて」「新しい順にではなく」）。
_NEGATION_RE = re.compile(r"\s*に?\s*(?:では|じゃ|で)な")


def read_direction(task: str | None) -> tuple[str, str] | None:
    """依頼文から向きを読む。戻りは (向き, 根拠の語)。向きは "asc" / "desc"。

    読めない時（語が無い／両向きが残る）は None。
    """
    if not task:
        return None
    found: dict[str, str] = {}
    for word, direction in DIRECTION_WORDS:
        for m in re.finditer(re.escape(word), task):
            if _NEGATION_RE.match(task, m.end()):
                continue
            found.setdefault(direction, word)
    if len(found) != 1:
        return None
    direction, word = next(iter(found.items()))
    return direction, word
