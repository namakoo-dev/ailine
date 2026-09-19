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

import re
from pathlib import Path

from ailine_core.book_view import BookView

#: 1 文字の値の「隣」に来てはいけない文字（同じ文字種が続くなら、それは別の語の一部）。
#:   数字の値 → 前後が数字 ／ 英字の値 → 前後が英字。半角と全角の両方を見る。
_SAME_KIND = {"digit": "0-9０-９", "alpha": "A-Za-zＡ-Ｚａ-ｚ"}


def mentions_value(task: str, value: str) -> bool:
    """依頼文が**この値そのもの**を言っているか（部分文字列の巻き添えを外す）。

    ★★ 2026-09-19（⚠ の家系を監査していて見つけた穴）: ここは長らく
      `len(v) >= 2 and v in task` で、**1 文字の値を一度も見ていなかった**:

          表     2行目 記号=A ／ 3行目 記号=B
          依頼   「2行目の記号を B にして」    ← B は 3 行目に在る（依頼が自己矛盾）
          関所   黙る                        ← ★ 1 文字なので見ていない

      ◎ ○ × 済 可 A B は帳簿でいちばん普通の値（引用値 313 件のうち 68 件が 1 文字）。
      同じ日に『済』で塞いだ残差ゲートの盲点と**同じ形**で、repo に 3 箇所あった最後の 1 つ。

    ★★ `len(v) >= 2` は無意味な制限ではなかった ── **外す前に理由を測った**:

          表の 3 行目に 値 '0' が在る状態で 「2行目の数量を**10**にして」
            → 素朴に 1 文字を許すと '0' が '10' の部分文字列として一致し、**誤爆**する

      だから「1 文字を許す」のではなく **文字種の境界で切る**（A1 記法の読みと同じ形）:
      数字は前後が数字でない時だけ／英字は前後が英字でない時だけ 言及とみなす。
      記号（◎ ○ × 済）は境界を持たないので、そのまま一致する。

    ★ 測って**直さないと決めた限界**: 「3行目の**C**コードを見て」のように英字 1 文字へ
      カタカナが続く形は拾ってしまう。依頼文 1,942 件を調べて**この形は 0 件**だった ──
      発明した検体に合わせて規則を曲げると、実在しうる『Bコース』のような値を落とす方が高くつく。
    """
    v = str(value or "")
    if not v or v not in (task or ""):
        return False
    if len(v) >= 2:
        return True
    kind = "digit" if v.isdigit() else ("alpha" if v.isascii() and v.isalpha() else None)
    if kind is None:
        return True                 # ★ 記号・漢字・かなは境界を持たない
    cls = _SAME_KIND[kind]
    return re.search(f"(?<![{cls}]){re.escape(v)}(?![{cls}])", task) is not None


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
        if mentions_value(task, v):
            return v
    return None
