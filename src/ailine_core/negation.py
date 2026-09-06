# -*- coding: utf-8 -*-
"""否定を「文に在るか」でなく「**何に付いているか**」で読む（2026-09-06）。

★★ なぜ在るか（実測 `bench/negation_reading_probe.py`・本物の CLI を --dry で 11 件）:
  否定の判定が**文全体**を見ていた（`any(w in task)`）ため、逆向きの事故が 2 つ在った。

    Q1 語彙の穴 ── 見ていたのは 3 語だけ（"以外" / "を除いた" / "を抜いた"）。
       「営業でない」「ではない」「じゃない」「を除いて」は否定と読まれず、
       LLM の `eq` が通った ── **依頼と逆の行に書いて ✓**（4/4 再現）。

    Q2 結び先の穴 ── 「メモ**以外**は変えずに、所属が営業の行のメモに○を付けて」で
       `nin` が強制され、**営業でない行に書いた**（3/3 再現）。「以外」は列の話なのに、
       条件の否定として読まれていた。★ 実害はこちらが重い（書く行が丸ごと入れ替わる）。

★ だから判定は 3 通りに割る。**足りない時に決めないこと**が肝で、
  「決められないなら断る」は既存の振る舞い（2026-09-05）をそのまま保つ:

      値に付いている    「営業以外」          → NEGATED  否定として読む
      列名に付いている  「メモ以外は変えずに」→ PLAIN    条件の話ではない
      どちらでもない    「エイギョウ以外」    → UNCLEAR  決めない（呼び出し側が断る）

★ 語の集合について: `ailine.py` の `_EXCEPT_WORDS` は `removal_reading`
  （「味噌汁の行を除いて」を削除と読む判断）と**共有**されている。そこへ語を足すと
  別の判断が黙って変わるので、**共有の集合を含む別の集合**をここに持ち、
  含有関係を番人が機械で縛る（tests/test_negation_binds_to_a_value.py）。

★ ailine を import しない（可搬性の番人 test_line_budget.py が機械で守る層）。
"""
from __future__ import annotations

NEGATED = "negated"      #: 否定として読む
PLAIN = "plain"          #: 否定ではない（条件は素直に読む）
UNCLEAR = "unclear"      #: 決めない ── 呼び出し側は**断る**こと

#: `ailine._EXCEPT_WORDS` と同じ 3 語（★ 番人が「これを含むこと」を機械で確かめる）
SHARED_WORDS = ("以外", "を除いた", "を抜いた")

#: 条件の否定として読む語（★ 実測で反転を確認した言い方だけを足した ──
#: 「再現しない言い方は触らない」を測る前に凍結してある）
NEGATION_WORDS = SHARED_WORDS + ("でない", "ではない", "じゃない", "を除いて", "を除く")


def _sticks_to(task: str, things) -> bool:
    """否定語が、渡した語のどれかの**直後**に付いているか。"""
    return any(f"{t}{w}" in task for t in (things or ()) if t for w in NEGATION_WORDS)


def reading(task: str, values, column_names) -> str:
    """条件が否定かを NEGATED / PLAIN / UNCLEAR で返す。

    values      … 条件列に**実在する**値のうち、依頼文が名指ししているもの
    column_names… その表の列名（★「メモ以外」のように列に付く否定を分けるため）

    ★ 順番に意味がある: 値への付き方を先に見る。列名と同じ文字列の値が在っても、
      条件として名指しされた値の方が強い。
    """
    text = task or ""
    if not any(w in text for w in NEGATION_WORDS):
        return PLAIN
    if _sticks_to(text, values):
        return NEGATED
    if _sticks_to(text, column_names):
        return PLAIN
    return UNCLEAR
