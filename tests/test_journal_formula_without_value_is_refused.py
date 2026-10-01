# -*- coding: utf-8 -*-
"""accounts: 読む列に「式のまま計算結果が無い」セルが在る仕訳は断る（2026-10-01・Namakoo 決裁 A）。

★ 値だけで読むと式のセルは空に見え、借方金額が式の行は候補から**黙って**落ちる。
  accounts は LibreOffice を起動しない設計なので、値を作らずに見つけて断る。
★ 読まない列に式が在るだけなら断らない（断りすぎない）。
"""
from __future__ import annotations

import openpyxl

from ailine_core import accounts_read


def _journal(tmp_path, amount_is_formula: bool, memo_is_formula: bool = False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["取引日", "借方勘定科目", "借方金額", "摘要", "メモ"])
    ws.append(["2026-09-01", "旅費交通費", 1200, "電車代", "a"])
    ws.append(["2026-09-02", "消耗品費", "=1000+500" if amount_is_formula else 1500,
               "文具", "=1+1" if memo_is_formula else "b"])
    p = tmp_path / "仕訳.xlsx"
    wb.save(p)   # ★ openpyxl は式の計算結果を書かない ── 「式のまま値が無い」冊そのもの
    return p


def test_formula_without_value_in_a_used_column_is_refused(tmp_path):
    book = accounts_read.read_journal(_journal(tmp_path, amount_is_formula=True))
    assert book.refused, "式のまま値の無い借方金額を空として読んだ"
    assert "C3" in book.refused and "計算結果が保存されていない" in book.refused, book.refused


def test_plain_values_are_read(tmp_path):
    book = accounts_read.read_journal(_journal(tmp_path, amount_is_formula=False))
    assert not book.refused, book.refused
    assert len(book.rows) == 2


def test_formula_in_an_unused_column_is_not_refused(tmp_path):
    book = accounts_read.read_journal(
        _journal(tmp_path, amount_is_formula=False, memo_is_formula=True))
    assert not book.refused, book.refused
