# -*- coding: utf-8 -*-
"""比較の境目の数は**依頼文**から機械が取る（2026-09-14・言い回し 120 件の盲検）。

★★ 判定者 2 人が一致した誤配 16 件のうち 4 件がこの家系:
    「残業時間が20時間超えてる人だけ教えて」 → 抽出 残業時間 = 0
    「在庫数が10個切ってるの教えて」         → 抽出 在庫数 = 0
    「在庫少ないやつ出して」                 → 抽出 在庫数 ≤ 0
    「発注しなきゃいけないもの一覧にして」   → 抽出 区分 = 発注（列に無い値・0 行が ✓ で出る）
  依頼文に数が**在る**のに LLM の `0` がそのまま通り、黙って違う答えが出ていた。

原因は 2 つ、どちらも兄弟間の片配線:
  ① 比較語の辞書が口語（超えてる／切ってる／下回る）を知らない
  ② SET_WHERE は「閾値は依頼文の数字から機械が取る」と 09-04 に直してあったのに EXTRACT には無かった
処方: 器官を 1 つ（`ailine_core/threshold.py`）にして両方が呼ぶ。数が無い／複数なら**断る**。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine   # noqa: E402
from ailine_core import threshold   # noqa: E402

HEAD = ["品番", "区分", "在庫数", "金額", "残業時間"]
ROWS = [["P-1", "資材", 3, 120000, 2.5], ["P-2", "事務用品", 45, 5000, 0],
        ["P-3", "消耗品", 0, 8000, 21.0]]


def _book(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表"
    ws.append(HEAD)
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    wb.close()
    return {"sheets": ["表"], "headers": {"表": HEAD}, "header_rows": {"表": 1}, "path": str(p)}


def _extract(meta, args, task):
    return ailine.verify_dsl_args("EXTRACT", args, meta, task=task)


# ── 器官（純関数）────────────────────────────────────────────────────────────

@pytest.mark.parametrize("task,want", [
    ("在庫数が10個切ってる", [10.0]),
    ("金額が10万円以上", [100000.0]),
    ("金額が５，０００円以上", [5000.0]),          # 全角と桁区切り
    ("金額が1,500円を超える", [1500.0]),
    ("残業が2.5時間超え", [2.5]),
    ("在庫が少ない", []),
    ("3千円以上5千円未満", [3000.0, 5000.0]),
])
def test_numbers_are_read_from_the_request(task, want):
    assert threshold.task_numbers(task) == want


def test_one_number_is_taken_and_the_llm_loses_with_disclosure():
    g = threshold.ground("残業時間が20時間超えてる行", 0, example="例")
    assert g.value == 20.0 and not g.refusal and "依頼文の数(20)" in g.warning, g


def test_no_number_is_refused_not_zero():
    """★ 「少ない」の境目を機械が 0 と決めない ── 空欄は誤値より安い。"""
    g = threshold.ground("在庫が少ない行", 0.0, example="在庫数が10未満の行を抜き出して")
    assert g.refusal and "境目の数が依頼文にありません" in g.refusal and g.value is None, g
    assert "在庫数が10未満" in g.refusal, "★ 通る書き方を添える"


def test_two_numbers_are_refused_as_ambiguous():
    g = threshold.ground("3千円以上5千円未満の行", 3000, example="例")
    assert g.refusal and "一意に決まりません" in g.refusal and "3000" in g.refusal and "5000" in g.refusal, g


# ── 比較語の口語 ─────────────────────────────────────────────────────────────

@pytest.mark.parametrize("task,want", [
    ("残業時間が20時間超えてる人", "gt"),
    ("在庫数が10個切ってるの", "lt"),
    ("在庫が発注点の5を下回る行", "lt"),
    ("金額が5000を上回る行", "gt"),
    ("金額が5000に満たない行", "lt"),
    ("原価が500以上の行を抜き出して", "gte"),      # ★ 凍結済みの検体は不変
])
def test_colloquial_comparisons_are_read(task, want):
    assert ailine.extract_cmp_from_task(task) == want


@pytest.mark.parametrize("task", [
    "品名を区切って別の列に",          # 区切って ≠ 切って（数字が近くに無い）
    "10日で締め切ってる案件",          # 締め切って ── 直前が数や数え語でない
    "予算を上回る努力をして",           # 数が無い
])
def test_fragments_of_other_words_do_not_fire(task):
    assert ailine.extract_cmp_from_task(task) is None, task


# ── EXTRACT が器官を呼ぶ（盲検の 4 件がそのまま検体）──────────────────────

def test_twenty_hours_over_becomes_gt_twenty(tmp_path):
    ok, r, _i, err = _extract(_book(tmp_path), {"col": "残業時間", "cmp": "eq", "value": "0"},
                              "残業時間が20時間超えてる行を抜き出して")
    assert ok, err
    assert (r["cmp"], r["value"]) == ("gt", 20.0), (r["cmp"], r["value"])
    assert any("依頼文の数(20)" in w for w in r.get("_warnings", [])), r.get("_warnings")


def test_under_ten_becomes_lt_ten(tmp_path):
    ok, r, _i, err = _extract(_book(tmp_path), {"col": "在庫数", "cmp": "eq", "value": "0"},
                              "在庫数が10個切ってる行を抜き出して")
    assert ok, err
    assert (r["cmp"], r["value"]) == ("lt", 10.0), (r["cmp"], r["value"])


def test_few_without_a_number_is_refused(tmp_path):
    """★ 「在庫少ないやつ出して」── 0 以下で進まない。"""
    ok, _r, _i, err = _extract(_book(tmp_path), {"col": "在庫数", "cmp": "lte", "value": 0.0},
                               "在庫数が少ない行を抜き出して")
    assert not ok and "境目の数が依頼文にありません" in err, err


def test_ten_man_yen_is_one_hundred_thousand(tmp_path):
    ok, r, _i, err = _extract(_book(tmp_path), {"col": "金額", "cmp": "gte", "value": 10},
                              "金額が10万円以上の行を抜き出して")
    assert ok, err
    assert r["value"] == 100000.0, r["value"]


def test_a_value_the_column_does_not_have_is_refused_with_the_real_ones(tmp_path):
    """★ 「発注しなきゃいけないもの」→ 区分 = 発注 ── 0 行の抽出が ✓ で出る穴。"""
    ok, _r, _i, err = _extract(_book(tmp_path), {"col": "区分", "cmp": "eq", "value": "発注"},
                               "発注しなきゃいけないものを抜き出して")
    assert not ok, "列に無い値で抽出を通した"
    assert "『発注』という値はありません" in err and "資材" in err and "消耗品" in err, err


def test_a_value_the_column_has_still_passes(tmp_path):
    """陰性対照 ── 列に在る値の等値は今までどおり。"""
    ok, r, _i, err = _extract(_book(tmp_path), {"col": "区分", "cmp": "eq", "value": "資材"},
                              "区分が資材の行を抜き出して")
    assert ok, err
    assert r["value"] == "資材"


def test_the_same_number_reading_serves_set_where_too(tmp_path):
    """★★ 片配線の番人 ── 兄弟の SET_WHERE も同じ器官で読む（万が両方で効く）。"""
    meta = _book(tmp_path)
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"cond_col": "金額", "cmp": "gte", "cond_value": 10, "col": "区分",
                      "value": "大口"}, meta,
        task="金額が10万円以上の行の区分に『大口』と入れて")
    assert ok, err
    assert r["cond_value"] == 100000.0, r["cond_value"]


def test_a_word_that_means_zero_is_a_boundary_too():
    """★★ ① の初版が**正だった 1 件を壊した**（実測で気づいた・盲検 #21）。

    「在庫数がマイナスになってる行を教えて」は数を言っていないが境目は 0 ── 判定者 2 人が
    正 と読んだ件なので、断りに落としてはいけない。★ 語は実際に打たれた物だけ（#21・#74）。
    """
    assert threshold.zero_boundary("在庫数がマイナスになってる行を教えて")
    assert not threshold.zero_boundary("在庫が少ない行を教えて")
    g = threshold.ground("在庫数がマイナスになってる行", 0.0, example="例")
    assert g.value == 0.0 and not g.refusal, g


def test_the_zero_word_survives_the_whole_path(tmp_path):
    ok, r, _i, err = _extract(_book(tmp_path), {"col": "在庫数", "cmp": "lt", "value": 0},
                              "在庫数がマイナスになってる行を教えて")
    assert ok, err
    assert (r["cmp"], r["value"]) == ("lt", 0.0), (r["cmp"], r["value"])
