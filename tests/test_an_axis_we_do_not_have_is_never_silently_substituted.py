# -*- coding: utf-8 -*-
"""★★ 持っていない集約を、黙って持っているもので代用しない。

★ 出所（2026-09-16 の掃き出し・実機で再現）:
  「部門ごとの**平均**金額を出して」が **合計** を計算して通った。
  出力シート『集計』の見出しは『合計 - 金額』、値は青果 400（平均なら 200）、
  それで **✓ 機械検証済み** が出た。★ 間違った答えに ✓ ── いちばん重い壊れ方。

★ 器官は在った: ailine_core/op_axes.py が「軸」と「まだ扱えないもの」を宣言し、
  「？ 平均はまだ扱えません（この道具の集約関数は 合計 だけです）」の文面まで持っていた。
  ところが読まれるのは**語彙外の断り／もしかしての提案**の経路だけで、
  LLM が実在の op を返した**成功経路では一度も読まれていなかった**（片配線）。

★ 直しは op ごとに書かない ── `verify_dsl_args` の**全 op が通る 1 箇所**で判定する。
  だからこの試験も、軸の名簿から回して全 op を 1 本で縛る。
"""
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from ailine_core.op_axes import AXES, judge_axis  # noqa: E402

META = {"sheets": ["売上"],
        "headers": {"売上": ["品名", "部門", "金額", "備考", "単価", "数量"]},
        "header_rows": {"売上": 1},
        "rows": {"売上": 5}}
ARGS = {"group_col": "部門", "value_col": "金額", "col": "金額",
        "operands": ["単価", "数量"], "operator": "*", "keys": ["品名"]}

#: 事故の形 ── どれも「持っていない集約/演算を頼んでいる」。黙って別の数字を返してはならない。
ACCIDENTS = [
    ("AGGREGATE", "部門ごとの平均金額を出して", "平均"),
    ("AGGREGATE", "部門ごとの最大金額を出して", "最大"),
    ("AGGREGATE", "部門ごとに金額の最小値を出して", "最小"),
    ("PIVOT", "部門ごとの平均金額を縦横でまとめて", "平均"),
    ("APPEND_TOTAL", "金額の平均を一番下に追加して", "平均"),
    ("COMPUTE_COLUMN", "単価の累乗の列を作って", "累乗"),
]

#: 陰性対照 ── 持っているものを頼む言い方。1 件も落とさないこと。
LEGITIMATE = [
    ("AGGREGATE", "部門ごとに金額を集計して"),
    ("AGGREGATE", "部門ごとの合計金額を出して"),
    ("AGGREGATE", "最大手の取引先ごとに集計して"),   # ★ 『最大』を含むが別の語（not_when）
    ("PIVOT", "部門ごとの合計金額を縦横でまとめて"),
    ("APPEND_TOTAL", "金額の合計を一番下に追加して"),
    ("COMPUTE_COLUMN", "単価と数量を掛けた列を作って"),
    ("DEDUP", "品番が同じ行を重複として除いて"),
]


def test_the_roster_of_axes_is_not_empty():
    """★ 名簿が空でも他の試験は全部緑になる（空回りの検出）。"""
    assert len(AXES) >= 5, sorted(AXES)


def test_aggregate_and_pivot_declare_the_same_axis_as_append_total():
    """★ 兄弟間の片配線を防ぐ ── 合計しか計算しない 3 op は同じ軸・同じ穴を宣言すること。

    ★ SummaryTable も PivotSum も合計固定。片方だけ宣言すると、同じ依頼が op によって
      断られたり黙って合計になったりする（2026-09-16 に実際そうなっていた）。
    """
    base = set(AXES["APPEND_TOTAL"].lacks)
    for op in ("AGGREGATE", "PIVOT"):
        assert op in AXES, f"{op} が軸を宣言していない（合計しか計算できないのに）"
        assert AXES[op].has == AXES["APPEND_TOTAL"].has, op
        missing = sorted(base - set(AXES[op].lacks) - {"件数"})
        assert not missing, f"{op} の軸に無い（APPEND_TOTAL には在る）: {missing}"


@pytest.mark.parametrize("op, task, word", ACCIDENTS)
def test_it_refuses_instead_of_substituting(op, task, word):
    ok, _r, _i, err = ailine.verify_dsl_args(op, dict(ARGS), META, task=task,
                                              target_sheet="売上")
    assert not ok, f"持っていない『{word}』を黙って代用した: {op} / {task}"
    assert "まだ扱えません" in err, err


@pytest.mark.parametrize("op, task", LEGITIMATE)
def test_it_does_not_refuse_what_we_do_have(op, task):
    verdict, _d = judge_axis(op, task, META["headers"]["売上"])
    assert verdict == "ok", f"持っているものを断った: {op} / {task} → {verdict}"


def test_the_gate_sits_where_both_routes_pass():
    """★ 片配線の番人 ── 単発も複合計画の段も通る `verify_dsl_args` に 1 箇所だけ。"""
    import inspect
    src = inspect.getsource(ailine.verify_dsl_args)
    assert src.count("judge_axis(") == 1, "軸の関所は 1 箇所（op ごとに書き分けない）"


def test_the_refusal_says_what_we_do_have_and_how_to_ask():
    """★ 断るだけでは仕事が進まない ── 持っている方を名指しし、通る例が在ること。

    ★ 「こう頼めます」の行は**呼び出し側の断りの画面**が出す（関所は 1 行に保つ）。
      関所側でも足したら実機で 2 回出た ── 導線は 1 本。だからここでは
      「その op に、そのまま打てる例が在る」ことを別に確かめる。
    """
    ok, _r, _i, err = ailine.verify_dsl_args("AGGREGATE", dict(ARGS), META,
                                              task="部門ごとの平均金額を出して",
                                              target_sheet="売上")
    assert not ok
    assert "合計 だけです" in err, err
    assert err.count("\n") == 0, f"断り文が 1 行に収まっていない: {err!r}"
    from ailine_core.examples import render_example_line
    assert render_example_line("AGGREGATE", "集計"), "画面に出す通る例が無い"
