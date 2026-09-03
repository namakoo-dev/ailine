"""compare_blocked — 「数値として比べられなかったセル」を**開示だけ**する器官。

★ なぜ在るか（2026-08-24 第三波 H3・盲検）: 金額列が文字列（"1,000" / "△1,500" /
全角）で入っている実物のフォルダに「金額が50000以上の行を抜き出して」を掛けると、
80,000 も 90,000 も一致せず **「計 0 行一致」で静かに終わる**。predicate が文字列を
不一致にするのは正しい（型を黙って変換しない＝憲法）。欠けていたのは判定ではなく
**理由を言う口**だった ── 「出ないことは信号でない」。

★ この module は**一度も判定に使わない**。使うと ✓ の意味が変わる（機械が勝手に
"△1,500" を -1500 と読んだことになる）。返すのは人へ見せる文字列だけ。

★ 片配線への備え: 単一ブック経路（ailine.check_extract 側）とフォルダ経路
（extract_multi）の**両方**から呼ぶ。両経路を同時に縛る番人は
tests/test_string_amounts.py（片方だけ直すと赤くなる）。

★ ailine を import しない（移植可能性の番人）。
"""
from __future__ import annotations

import re
import unicodedata

# gte/lte/gt/lt だけが「数値として比べる」比較（eq/contains は文字列でも成立しうる）。
NUMERIC_CMPS = frozenset({"gte", "lte", "gt", "lt"})

# 会計の実物で見る飾り: 通貨記号・単位・空白・桁区切り・和文の負号（△▲）・
# 括弧の負数。★ ここに足すのは「見たことがある形」だけ（推測で広げない）。
_STRIP_RE = re.compile(r'[¥￥$,\s円]')
_MINUS_MARKS = ("△", "▲", "−", "―", "ー", "‐", "-")


def looks_numeric(value) -> bool:
    """人が見れば数値だが、セルの型は文字列 ── という値か。
       ★ 判定には使わない（開示専用）。真なら『数字に見えるのに文字列』。"""
    if not isinstance(value, str):
        return False
    s = unicodedata.normalize("NFKC", value).strip()
    if not s:
        return False
    for mark in _MINUS_MARKS:
        if s.startswith(mark):
            s = s[len(mark):]
            break
    if s.endswith(")") and s.startswith("("):
        s = s[1:-1]
    s = _STRIP_RE.sub("", s)
    if not s:
        return False
    try:
        float(s)
    except ValueError:
        return False
    return True


def scan_column(values, cmp: str) -> dict | None:
    """列の値の並びを見て、数値比較から**落ちた数字に見える文字列**を数える。
       返り値: {"count": n, "samples": [...最大3件]} または None（開示不要）。"""
    if cmp not in NUMERIC_CMPS:
        return None
    stringy = [v for v in values if looks_numeric(v)]
    if not stringy:
        return None
    return {"count": len(stringy), "samples": [str(v) for v in stringy[:3]]}


def _fact(info: dict, column: str) -> str:
    """事実の一文（★ 文言の実装は 1 つ ── 経路ごとに書き直さない）。"""
    samples = "、".join(f"『{s}』" for s in info["samples"])
    return (f"列『{column}』の {info['count']} セルは、数字に見えますが"
            f"文字列として入っています（例: {samples}）")


def disclosure_inline(info: dict | None, column: str) -> str:
    """単一ブック経路（check_extract の reason は 1 本の文字列）用の 1 文。"""
    if not info:
        return ""
    return (f"（{_fact(info, column)} ── 数値の大小で比べられないため"
            "一致に数えていません。0 行は『該当なし』ではありません）")


def disclosure_lines(info: dict | None, column: str, matched_total: int) -> list:
    """フォルダ経路（複数行で見せられる）用。"""
    if not info:
        return []
    lines = [f"  ⚠ {_fact(info, column)}"]
    if matched_total == 0:
        lines.append("  → 数値の大小で比べられないため、この条件では 1 行も一致しません"
                     "（0 行は『該当なし』ではありません）")
    else:
        lines.append("  → これらの行は数値の大小で比べられないため、"
                     "一致にも不一致にも数えられていません")
    lines.append("     Excel で列を選び「区切り位置」→ 完了、または"
                 "『金額を数値に直して』で数値列にしてからお試しください")
    return lines
