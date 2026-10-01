"""色の名前 → RGB（16 進）の表 ── 引数の検査・Basic の生成・事後条件の 3 つが同じ表を引く。

★ 2026-09-24 に postconditions/_shared.py から**移しただけ**（本文は 1 文字も変えていない）。
  生成（codegen）が事後条件の内輪のモジュールを import していた配置を解くため、中立な冊にした。
  _shared は同じ実体を再輸出する。
"""
from __future__ import annotations

import re

COLOR_MAP = {
    "red": "FF0000", "green": "00B050", "blue": "0000FF", "yellow": "FFFF00",
    "orange": "FFA500", "purple": "800080", "pink": "FFC0CB", "black": "000000",
    "white": "FFFFFF", "gray": "808080", "grey": "808080",
    "lightblue": "ADD8E6", "lightgreen": "90EE90", "lightyellow": "FFFFE0",
    "lightred": "FFCCCC", "lightgray": "D3D3D3", "lightgrey": "D3D3D3",
}


# ★★ 2026-10-01（依頼の項の台帳で D だった項目・背景色）: 色は「COLOR_MAP に在るか」しか
#   見ていなかった。「見出しを黄色にして」に blue が返ると青く塗り、事後条件は宣言どおり
#   青いかを確かめて ✓ を出す ── 依頼の項が欠けていた。
#   ★ 依頼文の色の語を、**COLOR_MAP の鍵**へ対応させる。行き先は必ず COLOR_MAP の鍵
#     （番人 tests/test_fill_color_follows_the_request.py が縛る）── 別の色の表を作らない。
#   ★ 「薄い／淡い／ライト」＋色は COLOR_MAP の light＋鍵 が在る時だけ（lightblue）。
#     無い組（薄い紫）や「濃い・暗い・明るい」は**近い色に寄せず**読めないとする。
#   ★ 語は列挙で増やさない（足す前に何を奪うかを測る ── sort_direction と同じ約束）。
COLOR_WORDS = {
    "red": ("赤", "レッド"),
    "blue": ("青", "ブルー"),
    "yellow": ("黄", "イエロー"),
    "green": ("緑", "グリーン"),
    "orange": ("オレンジ", "橙"),
    "purple": ("紫", "パープル"),
    "pink": ("ピンク", "桃色"),
    "black": ("黒", "ブラック"),
    "white": ("白", "ホワイト"),
    "gray": ("灰", "グレー", "グレイ"),
    "lightblue": ("水色",),
}

#: 色の字を含むが色でない語（赤字・空白…）。先に潰してから読む。
_NOT_A_COLOR = re.compile(r"赤字|黒字|空白|余白|白紙|青果")
#: 淡くする修飾（COLOR_MAP の light＋鍵 へ）と、寄せられない修飾（読めない）。
_LIGHT = r"薄い|薄|淡い|淡|ライト"
_UNMAPPABLE = r"濃い|濃|暗い|ダーク|明るい|くすんだ|鮮やかな"
#: 色の語の直後の打ち消し（「赤以外」「赤じゃなくて」）── 在れば読まない。
_NEGATED = re.compile(r"\s*(?:色|い)?\s*(?:以外|じゃな|ではな|でな|を除)")


def read_color(task: str | None, names=()) -> tuple | None:
    """依頼文から背景色を読む。戻りは (COLOR_MAP の鍵, 根拠の語)。読めなければ None。

    ★ 読めない: 色の語が無い／2 色以上（「赤を青に」）／打ち消し（「赤以外」）／
      寄せられない修飾（「濃い青」「薄い紫」）。
    ★ names（列の名前・表の値）の中の語は数えない（『青果』部の行、『白石』さん）。
    """
    text = task or ""
    if not text:
        return None
    for n in sorted({str(x) for x in names if x}, key=len, reverse=True):
        text = text.replace(n, "・")
    text = _NOT_A_COLOR.sub("・", text)
    found: dict = {}
    for key, words in COLOR_WORDS.items():
        for w in words:
            for m in re.finditer(re.escape(w), text):
                if _NEGATED.match(text, m.end()):
                    return None
                head = text[:m.start()]
                light = re.search(r"(?:" + _LIGHT + r")\s*$", head)
                if re.search(r"(?:" + _UNMAPPABLE + r")\s*$", head):
                    return None
                k = key
                if light:
                    k = key if key.startswith("light") else "light" + key
                    if k not in COLOR_MAP:
                        return None
                found.setdefault(k, (light.group(0) if light else "") + w)
    if len(found) != 1:
        return None
    key, word = next(iter(found.items()))
    return key, word
