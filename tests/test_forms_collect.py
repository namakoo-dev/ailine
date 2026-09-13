# -*- coding: utf-8 -*-
"""forms_collect（帳票の一覧の並べ方）の純関数の番人。"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ailine_core import forms_collect                                   # noqa: E402
from ailine_core.field_record import Evidence, Record                   # noqa: E402


# --- 一覧から外すのは「5 項目とも無」の冊だけ（2026-09-13・§13）-------------------------
#
# ★ 陰性対照を**正確に 1 項目**で置く ── 通しの検体では「ちょうど 1 項目だけ取れた冊」を作りにくく、
#   「1 項目取れた冊まで外す」変異が素通りした（実測）。純関数で線を縛る。


def _single(field, value):
    return Record(field, (Evidence(rule="規則", value=value, at="C11", how="読み方"),))


def _none(field):
    return Record(field, (), blank_reason="見つかりません")


def test_only_a_book_with_nothing_at_all_is_left_out():
    nothing = {f: _none(f) for f in forms_collect.FIELDS}
    one = dict(nothing)
    one["請求額"] = _single("請求額", 3300)                      # ★ ちょうど 1 項目
    got = forms_collect.nothing_found([("空.xlsx", nothing), ("一つ.xlsx", one)])
    assert got == ["空.xlsx"], got
