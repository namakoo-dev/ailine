# -*- coding: utf-8 -*-
"""画面に出すものを、人が読むものに絞る（2026-09-14・買い手役 3 体・4 回の指摘）。

★★ 「毎回 15〜20 行の Basic が画面を埋め、肝心の `✓` が下へ押し流される」（3 体が別々に・計 4 回）。
  ルール変換の .bas は**既定で畳む**（`--show-basic` で出す）。
★ 自由生成（語彙外・AI が直接書いた .bas）は**畳まない** ── y/N を聞く前に人が見て決める物だから。
  ここを一緒に畳むと「読まずに同意」を作る（同意の前提が消える）。

同じ日に足した「元の表の合計行と明細の和の突き合わせ」の番人もここに置く ── どちらも
「人が締めの根拠にする 1 行が画面に在るか」の話。
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core import cli_render, split_people                                  # noqa: E402
from ailine_core.split_people import SplitPlan                                    # noqa: E402

_CODE = "Sub Main\n    Dim oSheet As Object\nEnd Sub\n"


def test_the_basic_source_is_folded_by_default():
    lines = cli_render.render_basic_block("─ 生成した .bas ─", _CODE)
    assert len(lines) == 1, lines
    assert "--show-basic" in lines[0] and "変更点" in lines[0], lines
    assert "Dim oSheet" not in "\n".join(lines)


def test_show_basic_prints_the_whole_block():
    lines = cli_render.render_basic_block("─ 生成した .bas ─", _CODE, show=True)
    assert lines == cli_render.render_code_block("─ 生成した .bas ─", _CODE), lines
    assert "Dim oSheet" in "\n".join(lines)


def test_the_run_subcommand_offers_the_flag():
    import argparse

    import ailine
    sub = [ac for ac in ailine.build_parser()._actions
           if isinstance(ac, argparse._SubParsersAction)][0]
    flags = {f for ac in sub.choices["run"]._actions for f in ac.option_strings}
    assert "--show-basic" in flags, sorted(flags)


def test_the_consent_gate_still_shows_the_ai_written_code():
    """★★ 自由生成は畳まない ── 同意の前に人が見る物（畳めば「読まずに同意」を作る）。

    ★ 本体は場所で決め打ちしない（分割で実装が動いても空振りしない）。
      ★ 2026-09-23: `inspect.getsource(ailine)` は**モジュール丸ごと＝本体 1 冊**だった。
        窓は文言の在るファイルの中で切り（window_around）、数は製品全体で数える。
    """
    from _product_source import count_in_product, window_around

    text = window_around("生成した .bas（語彙外・AI が直接作成）", after=0, before=10 ** 9)
    head = text.rfind("render_")
    assert text[head:].startswith("render_code_block"), text[head:][:60]
    calls = count_in_product("render_basic_block(") - count_in_product("def render_basic_block(")
    assert calls == 3, "ルール変換の 3 経路だけを畳む"


# ── 元の表の合計行と明細の和（会計役の MISSING #3）────────────────────────

def _plan(**kw):
    base = dict(header_row=1, by_header="担当者", by_column=3, amount_header="金額", amount_column=4,
                whole_amount=94720.0)
    base.update(kw)
    return SplitPlan(**base)


def test_a_stale_total_row_is_named_with_the_difference():
    plan = _plan(total_rows=[(13, 111880.0)])
    line = split_people.total_row_check(plan)
    assert line.startswith("⚠"), line
    assert "13 行目（111,880）" in line and "94,720" in line and "差 17,160" in line, line
    assert "分けた冊は明細のとおり" in line, line


def test_a_matching_total_row_is_said_out_loud():
    """★ 合っていても言う（出ないことを信号にしない ── 締めの根拠に使う）。"""
    line = split_people.total_row_check(_plan(total_rows=[(15, 94720.0)]))
    assert line.startswith("（元の表の合計行 15 行目（94,720）と明細の和が一致）"), line


def test_no_total_row_and_no_amount_say_nothing():
    assert split_people.total_row_check(_plan(total_rows=[])) is None
    assert split_people.total_row_check(_plan(total_rows=[(13, 1.0)], amount_column=None)) is None


def test_it_does_not_claim_a_mismatch_when_the_sum_is_incomplete():
    """★ 金額が文字の行が在る回は「合わない」と言わない（分母が欠けている）。"""
    line = split_people.total_row_check(_plan(total_rows=[(13, 111880.0)], unparsed=[7]))
    assert not line.startswith("⚠") and "突き合わせていません" in line, line


def test_the_report_carries_the_line():
    lines = cli_render.render_split_report("一覧.xlsx", "配る", {
        "by": "担当者", "amount": "金額", "parts": [], "blank": [], "multi": [], "excluded": [],
        "unparsed": [], "files_written": ["山田.xlsx"], "proof": {},
        "total_row_check": "⚠ 元の表の合計行 13 行目（111,880）が…"})
    assert any(ln.startswith("⚠ 元の表の合計行") for ln in lines), lines


# ── 文字の列の並べ替え（事務職の重い 1）── 実ファイルで 4 方向 ──────────────

def _book(path, rows):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()
    return path


_HEAD = ["取引先", "金額"]


def test_a_text_sort_that_kept_every_row_is_allowed_but_not_claimed(tmp_path):
    from ailine_core.postconditions import move
    before = _book(tmp_path / "before.xlsx", [_HEAD, ["丸山工業", 100], ["青山商会", 200], ["鈴木製作所", 300]])
    after = _book(tmp_path / "after.xlsx", [_HEAD, ["青山商会", 200], ["丸山工業", 100], ["鈴木製作所", 300]])
    status, reason = move.check_sort(after, {"col": "取引先", "order": "asc"}, header_row=1,
                                     source_book=before)
    assert status == "warn", (status, reason)
    assert "行の中身は 1 行も変わっていません" in reason and "並び順そのものは確かめていません" in reason


def test_a_text_sort_that_moved_values_between_rows_fails(tmp_path):
    """★★ 一番危ない形 ── 値だけが別の行に移った（金額が他社の物になる）。"""
    from ailine_core.postconditions import move
    before = _book(tmp_path / "before.xlsx", [_HEAD, ["丸山工業", 100], ["青山商会", 200], ["鈴木製作所", 300]])
    torn = _book(tmp_path / "torn.xlsx", [_HEAD, ["青山商会", 100], ["丸山工業", 200], ["鈴木製作所", 300]])
    status, reason = move.check_sort(torn, {"col": "取引先", "order": "asc"}, header_row=1,
                                     source_book=before)
    assert status == "fail", (status, reason)


def test_a_text_column_whose_equal_values_are_scattered_fails(tmp_path):
    """★ 並べ替えが走っていない徴候 ── 同じ値が離れて現れる。"""
    from ailine_core.postconditions import move
    rows = [_HEAD, ["丸山工業", 100], ["青山商会", 200], ["丸山工業", 300]]
    same = _book(tmp_path / "same.xlsx", rows)
    status, reason = move.check_sort(same, {"col": "取引先", "order": "asc"}, header_row=1,
                                     source_book=_book(tmp_path / "b.xlsx", rows))
    assert status == "fail" and "同じ値が離れて現れます" in reason, (status, reason)


def test_a_numeric_column_is_still_verified_to_the_order(tmp_path):
    """★ 陰性対照 ── 数字の列は今までどおり順序まで確かめて pass／崩れていれば fail。"""
    from ailine_core.postconditions import move
    ok = _book(tmp_path / "ok.xlsx", [_HEAD, ["a", 100], ["b", 200], ["c", 300]])
    assert move.check_sort(ok, {"col": "金額", "order": "asc"}, header_row=1)[0] == "pass"
    ng = _book(tmp_path / "ng.xlsx", [_HEAD, ["a", 200], ["b", 100], ["c", 300]])
    assert move.check_sort(ng, {"col": "金額", "order": "asc"}, header_row=1)[0] == "fail"
