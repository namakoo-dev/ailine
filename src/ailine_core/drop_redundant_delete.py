# -*- coding: utf-8 -*-
"""「合計行を**除いて**並べ替えて」の削除段は、頼まれていない。

★★ なぜ在るか（2026-09-09・盲検 C と D が**独立に**同じ所を指した）:

    依頼   「合計行を**除いて**売上の多い順に並べ替えて」
    計画   1段目 行削除（合計行）＋ 2段目 並べ替え
    実物   ★ **合計行が消える**（⚠ なので ✓ は出さないが、undo しないと戻らない）

  ★ 「除いて」は「**対象から外して**」であって「削除して」ではない。
  ★ しかも並べ替えは**元から合計行を外す**（`_skip_rows`・「データ行でないため
    並べ替えません」と画面に出る）。抽出・条件つき書換も同じ。
    → つまりこの削除段は「機械が既にやることを人が言い足しただけ」で、**落としてよい**。

★★ 計画だけでは区別できないと実測した（同じ 2 段が返る・順番すら安定しない）:

    「合計行を除いて…並べ替えて」    DELETE_ROWS + SORT
    「合計の行を消してから…並べ替えて」DELETE_ROWS + SORT   ← ★ こちらは削除が正しい

  ★ だから**依頼文の語**で分ける ── 「除いて／以外」は修飾、「削除・消す」は操作。
    両方在る回（「合計行を削除して、それから…」）は**落とさない**（消す意図が明示）。

★ 落としたことは必ず画面に出す（黙って段を減らさない ── 他の畳みと同じ作法）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations

#: 「対象から外す」と読める語（削除ではない）。
EXCLUDING_WORDS = ("除いて", "除き", "を除く", "以外", "抜いて", "抜きで", "のぞいて")

#: 「本当に消す」と読める語。★ 1 つでも在れば落とさない（消す意図が明示されている）。
REMOVING_WORDS = ("削除", "消して", "消す", "取り除")


def drop_delete_that_was_only_a_qualifier(plan, task: str, skips_rows_itself) -> tuple:
    """先頭の行削除が「除いて」の言い換えに過ぎないなら落とす。

    skips_rows_itself: op 名 → その op が合計行等を**自分で対象外にする**か（呼び出し側の宣言）。
    戻り値: (直した計画, 人に見せる 1 行 or None)
    """
    steps = list(plan or [])
    t = task or ""
    if len(steps) != 2:
        return steps, None
    if str((steps[0] or {}).get("op")) != "DELETE_ROWS":
        return steps, None
    if not skips_rows_itself(str((steps[1] or {}).get("op"))):
        return steps, None
    if not any(w in t for w in EXCLUDING_WORDS):
        return steps, None
    if any(w in t for w in REMOVING_WORDS):
        return steps, None      # ★ 消す意図が明示されている ── 落とさない
    return [steps[1]], ("（『除いて』は対象から外す意味なので、行を消す段は外しました "
                        "── この操作は合計行などを元から対象にしません）")
