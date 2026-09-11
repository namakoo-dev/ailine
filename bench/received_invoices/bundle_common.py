# -*- coding: utf-8 -*-
"""束の答えを読む共通部 ── ★ `ailine_core` を import しない（採点側の独立）。

★★ ここが製品の関数を呼んだ瞬間、分母は分母でなく感想になる（恒真）。
  暦の算術が製品と一致するのは**同じ暦を見ているから**であって、
  同じコードを見ているからではない ── 実装は別に持つ。
"""
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path

_DATE = re.compile(r"^\s*(\d{4})\s*[/\-年.]\s*(\d{1,2})\s*[/\-月.]\s*(\d{1,2})\s*日?\s*$")

#: 元号 → 元年の前年（元年 = +1）。★ 製品とは別に、暦から書く。
_ERA_BASE = {"明治": 1867, "大正": 1911, "昭和": 1925, "平成": 1988, "令和": 2018}
_WAREKI = re.compile(r"^\s*(明治|大正|昭和|平成|令和)\s*(\d{1,2}|元)\s*[/\-年.]"
                     r"\s*(\d{1,2})\s*[/\-月.]\s*(\d{1,2})\s*日?\s*$")


def load_bundles(path: Path) -> list:
    d = json.loads(path.read_text(encoding="utf-8"))
    return d["束一覧"] if isinstance(d, dict) else d


def norm_date(v):
    """答えと道具の出力を同じ形（'Y-M-D'）にする。読めない文字はそのまま返す。

    ★ 和暦をここに足したのは 2026-09-12 ── 器官が和暦を読めるようになった時、
      **物差しが和暦を知らないせいで製品の正解が『誤報』に見えた**（2 件）。
      予測は製品について立てていて、測定器の穴は予測に入っていなかった。
    """
    if isinstance(v, (dt.date, dt.datetime)):
        return f"{v.year}-{v.month}-{v.day}"
    if isinstance(v, str):
        m = _DATE.match(v)
        if m:
            y, mo, d = (int(x) for x in m.groups())
            return f"{y}-{mo}-{d}"
        m = _WAREKI.match(v)
        if m:
            era, y, mo, d = m.groups()
            yy = 1 if y == "元" else int(y)
            return f"{_ERA_BASE[era] + yy}-{int(mo)}-{int(d)}"
    return v


_HONORIFIC = ("御中", "様")
_GLUED_LABELS = ("請求日", "発行日", "請求番号", "請求書番号", "請求書No", "請求No")


def answer_value(field: str, v):
    """書き手は「セルの文字そのもの」を 値 に書いた（v2 の答えは 値=名前・セルの値=生 と分けていた）。
    道具は値を報告する側なので、答えを同じ土俵に下ろす: 宛先の敬称／同居ラベルを剥がす。
    ★ 剥がすのは答えの側だけ。道具の出力には触らない。"""
    if field == "請求番号" and isinstance(v, (int, float)) and not isinstance(v, bool):
        # ★ 裸の数値セル（実物 inv21 は請求番号を数値で持つ）。道具は識別子として
        #   文字で報告する ── 答えを同じ土俵に下ろす（道具の出力には触らない）。
        return str(int(v)) if float(v).is_integer() else str(v)
    if not isinstance(v, str):
        return v
    t = v.replace("　", " ").strip()
    if field == "宛先":
        for h in _HONORIFIC:
            if t.endswith(h):
                t = t[: -len(h)].strip()
    if field in ("請求日", "請求番号"):
        for sep in ("：", ":"):
            if sep in t and t.split(sep, 1)[0].strip() in _GLUED_LABELS:
                t = t.split(sep, 1)[1].strip()
    return t
