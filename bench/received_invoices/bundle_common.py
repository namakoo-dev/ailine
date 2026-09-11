# -*- coding: utf-8 -*-
"""束の答えを読む共通部 ── ailine_core を import しない（採点側の独立）。"""
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path

_DATE = re.compile(r"^\s*(\d{4})\s*[/\-年.]\s*(\d{1,2})\s*[/\-月.]\s*(\d{1,2})\s*日?\s*$")


def load_bundles(path: Path) -> list:
    d = json.loads(path.read_text(encoding="utf-8"))
    return d["束一覧"] if isinstance(d, dict) else d


def norm_date(v):
    """答えと道具の出力を同じ形（'Y-M-D'）にする。令和や読めない文字はそのまま返す。"""
    if isinstance(v, (dt.date, dt.datetime)):
        return f"{v.year}-{v.month}-{v.day}"
    if isinstance(v, str):
        m = _DATE.match(v)
        if m:
            y, mo, d = (int(x) for x in m.groups())
            return f"{y}-{mo}-{d}"
    return v


_HONORIFIC = ("御中", "様")


def answer_value(field: str, v):
    """書き手は「セルの文字そのもの」を 値 に書いた（v2 の答えは 値=名前・セルの値=生 と分けていた）。
    道具は値を報告する側なので、答えを同じ土俵に下ろす: 宛先の敬称／同居ラベルを剥がす。
    ★ 剥がすのは答えの側だけ。道具の出力には触らない。"""
    if not isinstance(v, str):
        return v
    t = v.replace("　", " ").strip()
    if field == "宛先":
        for h in _HONORIFIC:
            if t.endswith(h):
                t = t[: -len(h)].strip()
    if field in ("請求日", "請求番号"):
        for sep in ("：", ":"):
            if sep in t and t.split(sep, 1)[0].strip() in ("請求日", "発行日", "請求番号", "請求書番号", "請求書No", "請求No"):
                t = t.split(sep, 1)[1].strip()
    return t
