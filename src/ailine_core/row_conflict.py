# -*- coding: utf-8 -*-
"""依頼が「N 行目」と、表に実在する値の**両方**を名指していて、その値が N 行目に無い。

★★ なぜ在るか（2026-09-08・曖昧な行の測定で出た）:

    表     2行目 ナット / 3行目 **ボルト** / 4行目 ナット / 6行目 ナット
    依頼   「**3行目のナット**を削除して」        ← 依頼そのものが矛盾している
    解釈   操作:行削除 削除位置:3 行数:1(推定)   ← ★『ナット』がどこにも出ていない
    実物   **ボルト**が消えた
    出力   **✓ 機械検証済み**

  ★ 機械は番号だけを取り、値を黙って捨てた。残差の関所が黙るのは、
    『ナット』が**見出しでなく値**だから（あの関所は列名しか見ない）── 開けたまま
    持つと決めていた「値だけが落ちた」家系が、実害のある形で出た最初の例。

★★ 見るのは**依頼と実表**だけで、実行した操作は見ない ── 依頼が自己矛盾なら、
  何をしたとしても「頼まれたとおり」とは言えない。だから op を 1 つも列挙しない。

★ 誤爆を避ける 2 つの絞り:
  ・値は**表に実在するもの**だけ（依頼文から名前を切り出さない・A' 原則）
    → 「3行目の下に**新品**を追加して」は新品が表に無いので黙る（これから作る値）
  ・見出し行の値は見ない → 「3行目の**単価**を…」の『単価』は列名であって値ではない

★ 直さない・止めない ── ⚠ を出して ✓ を降ろすだけ（他の関所と同じ作法）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations

from pathlib import Path

from ailine_core.book_view import BookView


def _values_by_row(path, sheet: str | None, header_row: int) -> dict:
    """行番号（1 起点）→ その行に在る値の集合（見出し行は含めない）。"""
    out: dict = {}
    with BookView(Path(path)) as bv:
        ws = bv.sheet(sheet)
        for row in ws.iter_rows(min_row=header_row + 1):
            got = {str(c.value).strip() for c in row
                   if c.value is not None and str(c.value).strip() != ""}
            if got:
                out[row[0].row] = got
    return out


def value_not_in_the_named_row(task: str, row_no, path, sheet=None,
                               header_row: int = 1) -> str | None:
    """依頼が名指しした値が、依頼が名指しした行に無ければ、その値を返す。

    ★ 決められない材料（行番号が無い・表が読めない）なら None（黙る）。
    """
    if not task or not row_no:
        return None
    try:
        rows = _values_by_row(path, sheet, header_row)
    except Exception:
        return None
    here = rows.get(int(row_no))
    if here is None:
        return None                 # ★ その行が無いのは別の話（他の関所の受け持ち）
    elsewhere = set()
    for r, vals in rows.items():
        if r != int(row_no):
            elsewhere |= vals
    # ★ 長い値から当てる（「青りんご」と「りんご」が両方在る表で短い方だけ当たるのを防ぐ）
    for v in sorted(elsewhere - here, key=len, reverse=True):
        if len(v) >= 2 and v in task:
            return v
    return None
