# -*- coding: utf-8 -*-
"""セル結合は「値を畳む」と宣言する（2026-10-01・実機で確定・Namakoo 決裁 B）。

★ LibreOffice は結合した範囲の左上以外の値を消す。以前は「書式だけ」と宣言していたので、
  承知（--overwrite）の後も「書式だけのはずが値が変わった」で毎回 △ に落ちていた。
★ 消すことは適用前の関所が聞く。ここで縛るのは**消え方の前提**:
  値は空になる以外に変わらない・1 枚のシートの中だけ。
"""
from __future__ import annotations

import ailine
from ailine_core import write_precondition as wp


def _snap(cells: dict) -> dict:
    return {"cells": {k: (v,) for k, v in cells.items()}}


def _check(before, after):
    return wp.PRECONDITIONS["merge_fold"](
        _snap(before), _snap(after),
        cell_ref=lambda r, c: f"{chr(64 + c)}{r}", fmt_value=repr)


def test_merge_is_declared_as_folding_values():
    assert ailine.OP_WRITE_TARGET["MERGE"].writes == (ailine.WRITE_MERGE_FOLD,)
    assert "merge_fold" in wp.PRECONDITIONS


def test_values_that_only_become_empty_keep_the_precondition():
    before = {"売上!1,1": "商品", "売上!1,2": "数量", "売上!1,3": "金額"}
    after = {"売上!1,1": "商品"}
    assert _check(before, after) is None


def test_a_value_changed_to_something_else_breaks_it():
    before = {"売上!1,1": "商品", "売上!1,2": "数量"}
    after = {"売上!1,1": "商品", "売上!1,2": "数量X"}
    msg = _check(before, after)
    assert msg and "空以外" in msg, msg


def test_a_new_value_appearing_breaks_it():
    msg = _check({"売上!1,1": "商品"}, {"売上!1,1": "商品", "売上!5,5": 99})
    assert msg and "空以外" in msg, msg


def test_values_vanishing_on_two_sheets_breaks_it():
    before = {"売上!1,2": "数量", "原価!1,2": "単価"}
    msg = _check(before, {})
    assert msg and "複数のシート" in msg, msg

