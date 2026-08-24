# -*- coding: utf-8 -*-
"""実 LibreOffice の代わりに、生成された Basic を**読んで**真似る偽物。

★ 2026-08-24（土台固め）: 検分シートの書き手を openpyxl から Basic（LO 側）へ移した。
openpyxl の往復が xl/drawings の中の図形（描かれた角印・社判）を捨てるためで、実測で
帳票段が雛形の角印を全枚から消していた。

その結果、偽 basrun を使う検体では検分シートが作られなくなった。★ ここで assert を
緩めるのは間違い（検分シートは在るべき）── 直すのは治具の側。しかも 3 つの検体が
同じ偽物を書き写すと今日ずっと踏んでいる片配線になるので、**1 つにまとめる**。

★ 副産物: この偽物は生成された Basic を実際に解釈するので、生成側の形が壊れれば
ここが落ちる（文字列連結の書き方を変えた等）。単なるスタブより強い。
"""
from __future__ import annotations

import re

RS, US = chr(30), chr(31)

_CALL_RE = re.compile(
    r'Call WriteInspectionSheet\(oDoc,\s*"((?:[^"]|"")*)"\s*,\s*(.+?),\s*"([sn]*)"\)')


def _eval_basic_string(expr: str) -> str:
    """Basic の文字列式（"a" & Chr(31) & "b" & Chr(30) & …）を評価する。

    ★ 生の制御文字を .bas に埋めない設計なので、区切りは Chr(30)/Chr(31) の式で来る。
    """
    out = []
    for token in re.split(r"\s*&\s*", expr.strip()):
        token = token.strip()
        m = re.fullmatch(r"Chr\((\d+)\)", token)
        if m:
            out.append(chr(int(m.group(1))))
            continue
        if len(token) >= 2 and token[0] == '"' and token[-1] == '"':
            out.append(token[1:-1].replace('""', '"'))
            continue
        raise AssertionError(f"偽 LO が解釈できない Basic 断片: {token!r}")
    return "".join(out)


def apply_inspection_sheets(wb, code: str) -> int:
    """生成された Basic の WriteInspectionSheet 呼び出しを openpyxl 上で再現する。
       戻り値: 作った検分シートの枚数（0 なら呼び出しが無かった）。"""
    made = 0
    for m in _CALL_RE.finditer(code or ""):
        sheet_name = m.group(1).replace('""', '"')
        payload = _eval_basic_string(m.group(2))
        types = m.group(3)
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(title=sheet_name)
        for r, rec in enumerate(payload.split(RS), start=1):
            if rec == "":
                continue
            for c, field in enumerate(rec.split(US), start=1):
                kind = types[c - 1] if c - 1 < len(types) else "s"
                if r > 1 and kind == "n":
                    ws.cell(row=r, column=c, value=float(field) if "." in field else int(field))
                else:
                    ws.cell(row=r, column=c, value=field)
        made += 1
    return made
