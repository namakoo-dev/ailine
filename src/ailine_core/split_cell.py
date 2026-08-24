"""split_cell — 1 セルに詰まった複数値を、区切りで右の列へ割る（台帳 SPLIT_CELL・2 件）。

出所: 3203975「1 セルにまとめて記載された URL を 1 URL ごとに別セルに」/
1430969「CSV を指定フォーマットへ変換（項目分割・単位変換）」。

★ この op の ✓ が名乗れる根拠は 1 つだけ: **割った断片を同じ区切りで繋ぎ直すと元に戻る**。
「それらしく分かれた」ではなく「元と一致する」を機械が確かめる（恒真にならない検算）。
"""
from __future__ import annotations

from dataclasses import dataclass, field

# 依頼文でよく使われる区切りの言い方 → 実際の文字。★ 「改行」は \n（LibreOffice のセル内改行）。
SEPARATOR_ALIASES = {
    "改行": "\n", "\n": "\n", "newline": "\n",
    "、": "、", "，": "，", ",": ",", "カンマ": ",", "読点": "、",
    "/": "/", "／": "／", "スラッシュ": "/",
    " ": " ", "スペース": " ", "空白": " ", "半角スペース": " ",
    "・": "・", "|": "|", ";": ";", "；": "；", "セミコロン": ";",
}


def normalize_separator(raw) -> str | None:
    """依頼文/LLM から来た区切りの指定を実際の文字にする。読めなければ None。"""
    if raw is None:
        return None
    s = str(raw)
    if s in SEPARATOR_ALIASES:
        return SEPARATOR_ALIASES[s]
    stripped = s.strip()
    if stripped in SEPARATOR_ALIASES:
        return SEPARATOR_ALIASES[stripped]
    # 1〜3 文字の生の区切り（「, 」のような書き方）はそのまま使う
    if 1 <= len(s) <= 3:
        return s
    return None


def split_value(value, sep: str) -> list:
    """1 セルの中身を区切りで割る。★ 前後の空白は落とすが、**空の断片は残す**
       （「a,,b」は 3 つ ── 落とすと繋ぎ直しても元に戻らない）。"""
    if value is None:
        return []
    text = str(value)
    if not text:
        return []
    return [part.strip() for part in text.split(sep)]


def max_parts(values, sep: str) -> int:
    """必要な新しい列の数（一番多く割れる行に合わせる）。"""
    return max((len(split_value(v, sep)) for v in values), default=0)


@dataclass
class SplitCheck:
    rows_checked: int = 0
    mismatched: list = field(default_factory=list)   # (行番号, 元の値, 繋ぎ直した値)
    empty_rows: int = 0


def verify_rejoin(originals, parts_by_row, sep: str) -> SplitCheck:
    """★ 検算の本体: 各行について「割った断片を sep で繋ぎ直した文字列」が元と一致するか。

    originals:     元の列の値の並び（行順）
    parts_by_row:  割った結果の並び（行順・各行は新しい列の値のリスト・末尾の空欄を含む）
    """
    r = SplitCheck()
    for i, (orig, parts) in enumerate(zip(originals, parts_by_row), start=1):
        if orig is None or str(orig) == "":
            r.empty_rows += 1
            continue
        r.rows_checked += 1
        used = [p for p in parts if p not in (None, "")]
        rejoined = sep.join(str(p) for p in used)
        expected = sep.join(part for part in split_value(orig, sep) if part != "")
        if rejoined != expected:
            r.mismatched.append((i, str(orig), rejoined))
    return r


_SEP_LABELS = {chr(10): "改行", " ": "スペース", ",": "カンマ", "、": "、"}


def describe_separator(sep) -> str:
    """人が読む側の表記（解釈行）。改行やスペースはそのまま出すと見えないので名前で出す。"""
    return _SEP_LABELS.get(sep, str(sep))
