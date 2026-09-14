# -*- coding: utf-8 -*-
"""引き算・割り算の向きと、件数の穴（2026-09-14・言い回し 120 件の盲検・誤配の家系③）。

★★ 判定者 2 人が一致した誤配:
    「出勤と退勤の時刻から実働時間を計算する列を作って」 → 計算列 出勤 − 退勤（符号が逆）
    「残業時間、自動で計算して」                         → 出勤 − 退勤 を残業時間へ上書き
    「担当者ごとに何件受注したか件数も出して」           → 集計 担当者 × **数量の合計**
  「AとBから」は**向きを言っていない**のに並び順をそのまま演算の順にしていた。
  集計は合計しか無いのに、「何件」を数量の合計で**代用**していた。

契約:
  - 引き算・割り算は、依頼文が向きを言っていなければ**聞き返す**（足し算に落とさない）
  - 足し算・掛け算は向きが無いので触らない（陰性対照）
  - 件数は**数えられないと言う**（語彙の穴は穴と言う・代用しない）
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine   # noqa: E402
from ailine_core import threshold   # noqa: E402

META = {"sheets": ["S"], "header_rows": {"S": 1},
        "headers": {"S": ["出勤", "退勤", "実働時間", "担当者", "金額", "数量", "単価"]}}


def _compute(args, task):
    return ailine.verify_dsl_args("COMPUTE_COLUMN", dict(args), META, task=task)


def _agg(args, task):
    return ailine.verify_dsl_args("AGGREGATE", dict(args), META, task=task)


# ── 器官（純関数）────────────────────────────────────────────────────────────

@pytest.mark.parametrize("task,want", [
    ("退勤から出勤を引いた実働時間の列を作って", ["退勤", "出勤"]),
    ("出勤を退勤から引いて実働時間に", ["退勤", "出勤"]),
    ("金額を数量で割って単価を出して", ["金額", "数量"]),
    ("退勤マイナス出勤の列を作って", ["退勤", "出勤"]),
    ("出勤と退勤の時刻から実働時間を計算する列を作って", None),   # ★ 向きを言っていない
    ("実働時間を計算して", None),                                  # 列が文に無い
])
def test_the_direction_is_read_only_when_the_request_says_it(task, want):
    div = "割" in task
    got = threshold.direction_of(task, ["金額", "数量"] if div else ["出勤", "退勤"],
                                 "/" if div else "-")
    assert got == want, got


# ── 計算列（盲検の 2 件がそのまま検体）──────────────────────────────────────

def test_from_a_and_b_is_asked_back_not_guessed():
    ok, _r, _i, err = _compute({"operands": ["出勤", "退勤"], "operator": "-", "target": "実働時間"},
                               "出勤と退勤の時刻から実働時間を計算する列を作って")
    assert not ok, "向きを推測して通した"
    assert "どちらから引くのか" in err and "退勤から出勤を引いた" in err, err


def test_an_explicit_direction_wins_over_the_llm_order():
    ok, r, _i, err = _compute({"operands": ["出勤", "退勤"], "operator": "-", "target": "実働時間"},
                              "退勤から出勤を引いた実働時間の列を作って")
    assert ok, err
    assert r["operands"] == ["退勤", "出勤"], r["operands"]
    assert any("依頼文の向き" in w for w in r.get("_warnings", [])), r.get("_warnings")


def test_division_direction_is_read():
    ok, r, _i, err = _compute({"operands": ["金額", "数量"], "operator": "/", "target": "単価"},
                              "金額を数量で割って単価を出して")
    assert ok, err
    assert r["operands"] == ["金額", "数量"]


def test_multiplication_is_not_gated():
    """★ 陰性対照 ── 掛け算に向きは無い（聞き返したら邪魔なだけ）。"""
    ok, r, _i, err = _compute({"operands": ["数量", "単価"], "operator": "*", "target": "金額"},
                              "単価に数量を掛けて金額を作って")
    assert ok, err
    assert r["operands"] == ["数量", "単価"]


def test_no_task_means_no_gate():
    """★ 依頼文が無い経路（DSL 直渡し）は触らない ── 接地する相手が無い。"""
    ok, _r, _i, err = _compute({"operands": ["出勤", "退勤"], "operator": "-"}, "")
    assert ok, err


# ── 集計の件数（語彙の穴は穴と言う）──────────────────────────────────────

def test_counting_is_refused_not_substituted():
    ok, _r, _i, err = _agg({"group_col": "担当者", "value_col": "数量"},
                           "担当者ごとに何件受注したか件数も出して")
    assert not ok, "件数を数量の合計で代用した"
    assert "件数は数えられません" in err and "合計する列" in err, err


def test_a_named_sum_column_still_passes_even_with_the_word_count():
    """★ 陰性対照 ── 合計する列を名指ししていれば通す（「合計と件数」の回）。"""
    ok, r, _i, err = _agg({"group_col": "担当者", "value_col": "金額"},
                          "担当者ごとに金額の合計と件数を出して")
    assert ok, err
    assert r["value_col"] == "金額"


def test_a_plain_aggregate_is_unchanged():
    ok, r, _i, err = _agg({"group_col": "担当者", "value_col": "金額"}, "担当者ごとに金額を集計して")
    assert ok, err
    assert (r["group_col"], r["value_col"]) == ("担当者", "金額")


def test_a_column_missing_from_the_request_is_not_this_gates_business():
    """★ 依頼文に出てこない列が混じる回は「どの列か」の食い違い ── 既存の関所（確認）に渡す。

    ★★ 実測でここを踏んだ: 凍結検体（上書きの関所・exit 7）が、向きの断り（exit 3）に
      変わっていた。関所は**自分の担当だけ**を止める（聞けば済む回を断りにしない）。
    """
    ok, r, _i, err = _compute({"operands": ["出勤", "金額"], "operator": "-", "target": "実働時間"},
                              "出勤から退勤を引いた実働時間の列を作って")
    assert ok, err
    assert r["operands"] == ["出勤", "金額"], r["operands"]


@pytest.mark.parametrize("task,operands,operator", [
    ("売上から原価を引いた列と、原価から売上を引いた列を作って", ["売上", "原価"], "-"),
    ("金額を数量で割った単価と、数量を金額で割った列", ["金額", "数量"], "/"),
])
def test_a_sentence_that_says_both_directions_is_not_decided(task, operands, operator):
    """★ 両方の向きが文に在る回は**選ばない**（1 つに決まる時だけ採る）。

    ★ 複合の依頼（両方の列がほしい）で実際に起きる形 ── ここで片方を採ると、
      頼んだ 2 つのうち 1 つが黙って別物になる。
    """
    assert threshold.direction_of(task, operands, operator) is None
