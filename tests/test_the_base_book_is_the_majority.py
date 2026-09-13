# -*- coding: utf-8 -*-
"""scan / stack の基準は、名前順の先頭でなく**多数派**から選ぶ（2026-09-13・買い手役の初見・経理）。

★★ `10月_歓迎会出欠.xlsx`（異物 1 冊）が名前順で先頭に来ただけで、本物の請求書 20 冊が全部
  「取れなかった（欠け: 氏名, 出欠, 備考）」になった。基準は 1 行目の見出しが最も多くの冊と同じ冊。
★ 同じ雛形ばかりの束では旧版と同じ冊（名前順の先頭）── 既存の束の点数は動かない。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ailine_core.multifile import open_base_workbook            # noqa: E402


def _book(path, headers):
    wb = openpyxl.Workbook()
    wb.active.append(headers)
    wb.active.append([1] * len(headers))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()
    return path


def test_an_odd_book_that_sorts_first_is_not_the_base(tmp_path):
    odd = _book(tmp_path / "10月_歓迎会出欠.xlsx", ["氏名", "出欠", "備考"])
    real = [_book(tmp_path / f"請求_{i}.xlsx", ["品名", "数量", "金額"]) for i in range(3)]
    base, wb = open_base_workbook(sorted([odd, *real], key=lambda p: p.name))
    wb.close()
    assert base != odd, base
    assert base == min(real, key=lambda p: p.name), base       # 多数派の中では名前順


def test_a_uniform_bundle_keeps_the_first_book_as_base(tmp_path):
    """★ 陰性対照 ── 同じ雛形ばかりなら旧版と同じ（名前順の先頭）。"""
    books = [_book(tmp_path / f"{n}.xlsx", ["品名", "数量", "金額"]) for n in ("a", "b", "c")]
    base, wb = open_base_workbook(books)
    wb.close()
    assert base == books[0]


def test_an_unopenable_book_is_never_the_base(tmp_path):
    bad = tmp_path / "a_こわれた.xlsx"
    bad.write_bytes(b"text")
    good = _book(tmp_path / "b.xlsx", ["品名"])
    base, wb = open_base_workbook([bad, good])
    wb.close()
    assert base == good
