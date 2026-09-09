# -*- coding: utf-8 -*-
"""「空行を挿してから値を入れる」2 段を、1 段に畳む。

★★ なぜ在るか（2026-09-09・盲検 D が false refusal として拾った）:

    依頼   「北斗精機の行の**下に**、取引先「西村工業」の行を追加して、
             項目は事務机、件数は 2、単価は 15000 にして」
    計画   1段目 INSERT_ROWS at=8 count=1   ← 空行を挿す
           2段目 ADD_ROW     at=9 values={取引先:西村工業, 項目:事務机, …}
    出力   **？ 1 回の依頼で行を 2 回足そうとしています**（断り）

  ★ モデルは間違っていない ── 人が手でやる手順（空行を挿してから埋める）を
    そのまま書いている。ADD_ROW は**それ自体が行を作る**ので 1 段目は要らず、
    機械の「二重宣言だ」という判定も正しい。
    ★ つまり「断りが正しくて、しかし人が困る」形。**畳めば 1 段で済む。**
  ★ 3/3 で同じ計画が返った（モデルの気まぐれではなく、この言い方の常態）。

★★ 畳んでよい条件を狭く取る（迷ったら畳まない）:

    ・1 段目が INSERT_ROWS で、2 段目が ADD_ROW
    ・2 段目の行が、1 段目で空けた範囲の中に在る
    ・1 段目が空ける行数と、2 段目が埋める行数（＝1）が釣り合う

  ★ 釣り合わない回（3 行空けて 1 行だけ埋める）は**畳まない** ── 人は本当に
    空行が欲しいのかもしれない。黙って解釈を狭めない。
★ 畳んだことは必ず画面に出す（呼び出し側が言う ── fold_identical_steps と同じ作法）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations


def _int(v, default=None):
    try:
        return int(str(v).strip())
    except Exception:
        return default


def fold_insert_then_add(plan) -> tuple:
    """`INSERT_ROWS` の直後に、空けた行へ `ADD_ROW` する形を 1 段へ畳む。

    戻り値: (畳んだ計画, 畳んだなら人に見せる 1 行 or None)
    """
    steps = list(plan or [])
    if len(steps) != 2:
        return steps, None
    first, second = steps[0] or {}, steps[1] or {}
    if str(first.get("op")) != "INSERT_ROWS" or str(second.get("op")) != "ADD_ROW":
        return steps, None
    a1 = dict(first.get("args") or {})
    a2 = dict(second.get("args") or {})
    at1, at2 = _int(a1.get("at")), _int(a2.get("at"))
    count = _int(a1.get("count"), 1)
    if at1 is None or at2 is None or count != 1:
        return steps, None
    # ★ 空けた 1 行そのものへ入れる回だけ（挿した位置か、その直後）
    if at2 not in (at1, at1 + 1):
        return steps, None
    if not (a2.get("values") or {}):
        return steps, None      # ★ 値が無いなら、それは本当に空行が欲しい回
    # ★★ 2026-09-09（畳んだ直後に実測して直した）: 位置に **LLM の数字を持ち込まない**。
    #   最初は `at1` を採ったが、モデルの数え方は実表とずれており（北斗精機は 6 行目
    #   なのに at=8）、「北斗精機の**下に**」が**上**へ入った。
    #   ★ 位置は依頼文と実表から機械が決める（A' 原則）── 畳みは「値を運ぶ」だけにして、
    #     `at` は落とす。落ちた後は resolve_row_anchor が依頼文から解き直す。
    _values = {k: v for k, v in a2.items() if k != "at"}
    return [{"op": "ADD_ROW", "args": _values}], (
        "（空行を挿してから値を入れる 2 段を、1 段にまとめました "
        "── 行を足すのは 1 回です）")
