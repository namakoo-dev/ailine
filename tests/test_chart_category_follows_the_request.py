# グラフの横軸の列は依頼文と突き合わせる（2026-10-01・依頼の項の台帳で B だった項目）。
#
# ★ 形: LLM が返した横軸（category_col）は「実在するか」しか見ていなかった。事後条件
#   （check_chart_series）は種類と値の列しか見ないので、依頼が言っていない列を横軸にしても ✓。
#   実走行の実例:「金額の棒グラフを作って」に 横軸列:月（依頼に無い）。
#
# 契約:
#   ① 依頼文が横軸の列を名指ししていれば（「部門ごとの金額」）黙る
#   ② 実表で決まる回は黙る ── LLM も既定（先頭列）を選んだ／値の列以外に列が 1 つしか無い／
#      LLM が空で機械が先頭列を入れた（従来どおり (推定)）
#   ③ それ以外（依頼が言っていない別の列を LLM が選んだ）は ⚠（_warnings ＝ ✓→△ の材料）
#   ④ ⚠ の回も書き換えない ── 同義語（「客先ごと」→『取引先』）は照合の外で、既定へ書き換えると
#      正しい選択を壊す（読めない時に推測で書き換えない）

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import argcheck  # noqa: E402

HEAD = ["部門", "月", "担当", "金額"]


def _chart(task, cat, value="金額", head=HEAD):
    meta = {"sheets": ["売上"], "headers": {"売上": list(head)}, "header_rows": {"売上": 1}}
    args = {"value_col": value, "kind": "bar"}
    if cat is not None:
        args["category_col"] = cat
    ok, r, inferred, err = ailine.verify_dsl_args("CHART", args, meta, task=task)
    assert ok, err
    return r, inferred


def _cat_warnings(r):
    return [w for w in r.get("_warnings", []) if "横軸の列" in w]


# --- ① 名指しされていれば黙る -----------------------------------------------------------------

def test_named_category_is_silent():
    r, _i = _chart("月ごとの金額を棒グラフにして", "月")
    assert r["category_col"] == "月" and not _cat_warnings(r)


# --- ② 実表で決まる回は黙る ---------------------------------------------------------------------

def test_the_default_first_column_chosen_by_the_model_is_silent():
    r, _i = _chart("金額の棒グラフを作って", "部門")
    assert r["category_col"] == "部門" and not _cat_warnings(r)


def test_only_one_column_besides_the_value_is_silent():
    r, _i = _chart("金額の棒グラフを作って", "部門", head=["金額", "部門"])
    assert r["category_col"] == "部門" and not _cat_warnings(r)


def test_no_category_from_the_model_keeps_the_old_default():
    r, inferred = _chart("金額の棒グラフを作って", None)
    assert r["category_col"] == "部門" and "category_col" in inferred
    assert not _cat_warnings(r)


# --- ③ 依頼が言っていない別の列は ⚠ -----------------------------------------------------------

def test_unnamed_category_is_disclosed():
    """実走行の実例の形。"""
    r, _i = _chart("金額の棒グラフを作って", "月")
    ws = _cat_warnings(r)
    assert len(ws) == 1 and "『月』" in ws[0] and "『部門』" in ws[0]


# --- ④ 書き換えない ------------------------------------------------------------------------------

def test_the_guess_is_not_rewritten():
    head = ["日付", "取引先", "金額"]
    r, _i = _chart("客先ごとの売上推移を折れ線グラフで見たい", "取引先", head=head)
    assert r["category_col"] == "取引先" and _cat_warnings(r)
    assert "category_col" not in r.get("_sources", {})


# --- 変異: 照合を恒真にすると LLM の推測が通る ------------------------------------------------

def test_if_every_name_matched_the_guess_would_pass(monkeypatch):
    monkeypatch.setattr(argcheck, "name_matches_task", lambda *a, **k: True)
    r, _i = _chart("金額の棒グラフを作って", "月")
    assert r["category_col"] == "月" and not _cat_warnings(r)       # ← 直す前の挙動
