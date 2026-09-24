"""依頼文の引用符（「」『』"" 等）の中の語 ── 位置の解決・入れ替え・新しい列名・セルへの書き込みが共有する土台。

★ 2026-09-24 に src/ailine/__init__.py から**移しただけ**（本文は 1 文字も変えていない）。
★ ailine を import しない（持ち出せる部品）。本体は from ailine_core.quotes import ... で束ね直す。
"""
from __future__ import annotations

import re


# --- ★ A' 原則(致命3・W10e): SET_COLUMN_VALUE が書き込む定数値を LLM から切り離す ------
#   依頼文の引用符（「」『』""''）で囲まれた文字列を機械抽出する。extract_rate_factor と
#   同じ考え方 — ちょうど1つに絞れる時だけ確定・0件/2件以上は CLARIFY に委ねる（None）。
_QUOTE_PATTERNS = (
    re.compile(r"「([^」]+)」"),
    re.compile(r"『([^』]+)』"),
    re.compile(r'"([^"]+)"'),
    re.compile(r"'([^']+)'"),
)


def extract_quoted_literal(text: str) -> str | None:
    """依頼文全体を通して、引用符で囲まれた文字列がちょうど1つだけ見つかった場合に
       その中身を返す。0個/2個以上は曖昧とみなし None（機械確定を諦める＝呼び出し側が
       CLARIFY にする）。"""
    if not text:
        return None
    found = []
    for pat in _QUOTE_PATTERNS:
        found.extend(m.group(1) for m in pat.finditer(text))
    if len(found) == 1:
        return found[0]
    return None


def _task_outside_quotes(task: str) -> str:
    """引用符で囲まれた所を空白に潰した依頼文（＝**値でない部分**だけ）。

    ★ 「」の中は「そのセルに書く文字列」── 対象（列・行）の名指しとして読むと、
      値の一部分がたまたま列名と一致しただけで、頼んでいない列が対象になる。
    """
    out = str(task or "")
    for pat in _QUOTE_PATTERNS:
        out = pat.sub(lambda m: " " * len(m.group(0)), out)
    return out
