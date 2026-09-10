# -*- coding: utf-8 -*-
"""新しく作った行が、依頼が言った**名前**を持っていること。

★★ 出所（2026-09-10・出荷前の実機テストが落ちて辿り着いた・実測 4/4）:

    依頼   「5行目に丸山工業の行を作って 件数3 単価1500」
    宣言   入れる値:件数=3／単価=1500        ← 丸山工業がどこにも無い
    実物   5行目 = [None, 3, 1500, '=B5*C5']
    画面   **✓ 機械検証済み**（⚠ すら出ない）

  三項（依頼・宣言・実体）のうち宣言と実体は一致しているので事後条件は通る。
  `residue` は列名しか見ないので、値である『丸山工業』は端から対象外。

★ ここで縛るのは 2 つ:
    ① 依頼に無い値を書かない（捏造） ── values_not_grounded_in_the_request
    ② 新しい行のキー列が空なら ✓ を出さない ── detect_new_row_missing_key
  どちらも**実測で見つけた**（①は 11 回中 7 回・②は 4/4）。
"""
import pytest

import ailine
from ailine_core.intent import values_not_grounded_in_the_request


# ── ① 捏造した値を落とす ─────────────────────────────────────────
def test_invented_values_are_dropped():
    """依頼に無い値は落とす（実測: 件数1・単価1000・金額1000 を発明していた）。"""
    got = values_not_grounded_in_the_request(
        "5行目に丸山工業の行を作って",
        {"取引先": "丸山工業", "件数": 1, "単価": 1000, "金額": 1000})
    assert got == ["件数", "単価", "金額"]


def test_values_the_person_actually_asked_for_survive():
    """★ 逆向きの誤り（頼んだ値を落とす）を起こさない ── こちらの方が怖い。"""
    for task, values in [
            ("みかんの下に梨を追加して 売上2000 原価1200",
             {"商品": "梨", "売上": 2000, "原価": 1200}),
            ("5行目に丸山工業を追加して 単価は千円で",      # 漢数字
             {"取引先": "丸山工業", "単価": 1000}),
            ("5行目に丸山工業の行を作って 単価 1,500円",     # 桁区切りと単位
             {"取引先": "丸山工業", "単価": 1500}),
            ("5行目に丸山工業の行を作って 件数３ 単価１５００",  # 全角
             {"取引先": "丸山工業", "件数": 3, "単価": 1500}),
    ]:
        assert values_not_grounded_in_the_request(task, values) == [], task


def test_a_number_is_compared_as_a_number_not_as_text():
    """★ 部分文字列で照合すると `件数=1` が依頼文の `1000` に当たる（自分で踏んだ）。"""
    assert values_not_grounded_in_the_request(
        "単価1000で追加して", {"件数": 1}) == ["件数"]


# ── ② キー列が空なら ✓ を出さない ────────────────────────────────
def test_a_new_row_without_its_key_is_called_out():
    """実測 4/4 の形 ── 『丸山工業の行』に丸山工業が無いのに ✓ が出ていた。"""
    msg = ailine.detect_new_row_missing_key(
        "ADD_ROW",
        {"_headers": ["取引先", "件数", "単価", "金額"],
         "values": {"件数": 3, "単価": 1500}},
        {})
    assert msg and msg.startswith("★ 疑わしい"), msg
    assert "取引先" in msg
    # ★ ✓ を降ろす仕掛けに載っていること（★ 付きを数える既存の関所）。
    assert ailine.count_suspicious_advisories([msg]) == 1


def test_it_stays_quiet_when_the_key_is_there():
    """★ 誤爆 0 を縛る（実測: 主語が入った 6 回すべてで黙った）。"""
    assert ailine.detect_new_row_missing_key(
        "ADD_ROW",
        {"_headers": ["取引先", "件数"], "values": {"取引先": "丸山工業"}}, {}) is None


def test_other_ops_are_not_touched():
    """★ 行を作らない op に配線が漏れ出していないこと。"""
    for op in ("SORT", "DELETE_ROWS", "SET_CELL_VALUE", "EXTRACT", None):
        assert ailine.detect_new_row_missing_key(
            op, {"_headers": ["取引先"], "values": {}}, {}) is None, op


def test_an_inherited_formula_in_the_key_column_is_not_empty():
    """★ キー列が式の継承で埋まる表では鳴らない（空のままではないため）。"""
    assert ailine.detect_new_row_missing_key(
        "ADD_ROW",
        {"_headers": ["連番", "取引先"], "values": {"取引先": "丸山工業"},
         "_inherit_cols": [0]}, {}) is None


# ── ★ 番人が配線に載っていること（在っても鳴らない、を塞ぐ） ──────────
def test_the_guard_is_wired_into_the_advisory_assembly():
    """★ 器官を書いても、助言の組み立てに載っていなければ画面には出ない。

    ★ この repo が繰り返し踏んだ「在っても鳴らない」── 呼ばれていることを
      **本文の文字列**で確かめる（op ごとの配線ではなく 1 箇所であることも縛る）。
    """
    import inspect
    src = inspect.getsource(ailine._structural_advisories)
    assert src.count("detect_new_row_missing_key") == 1, (
        "助言の組み立てから呼ばれていない、または 2 箇所に配られている")



def _meta(tmp_path):
    """★ tests/test_table_basics.py の _anchor_meta と同じ形（治具を割らない）。"""
    import openpyxl
    p = tmp_path / "anchor.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "売上"
    ws.append(["商品", "売上", "原価"])
    for r in [["りんご", 1000, 600], ["みかん", 800, 400], ["ぶどう", 1200, 700]]:
        ws.append(r)
    wb.save(p)
    return {"sheets": ["売上"], "headers": {"売上": ["商品", "売上", "原価"]},
            "header_rows": {"売上": 1}, "path": str(p)}


# ── ★ 治具を実物らしくした分の被覆を、ここで明示的に埋める ────────────
def test_the_whole_pipe_refuses_when_nothing_was_asked_for(tmp_path):
    """依頼文に値が 1 つも無いのに values が来たら、**入口で断る**。

    ★ なぜここに在るか: tests/test_table_basics.py の 4 本は `task="t"` や
      「3行目に足して」という**置き場所の都合の文字列**で values を渡していた。
      それらは列の実在・並びの列名付け・None を書かない を縛る試験なので、
      2026-09-10 に依頼文を実物らしく直した（assert は 1 文字も変えていない）。
      その結果**この場面の被覆が消える**ので、ここで縛り直す。
      ── 消えたものは diff に出ない。
    """
    ok, _resolved, _inf, err = ailine.verify_dsl_args(
        "ADD_ROW", {"at": 3, "values": {"商品": "梨", "売上": 600, "原価": 300}},
        _meta(tmp_path), task="3行目に足して")   # ★ 梨も 600 も 300 も依頼文に無い
    assert not ok, "頼まれていない値を書こうとしたのに通した"
    assert "頼まれていない値は書きません" in err, err


def test_the_grounded_part_survives_even_when_the_rest_is_invented(tmp_path):
    """★ 一部だけ接地している時は、接地した分だけ残る（全部落とさない）。"""
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "ADD_ROW", {"at": 3, "values": {"商品": "梨", "売上": 600, "原価": 300}},
        _meta(tmp_path), task="3行目に梨を足して")   # ★ 梨だけが依頼文に在る
    assert ok, err
    assert resolved["values"] == {"商品": "梨"}, resolved["values"]
    assert "売上" in resolved["_dropped_label"] and "原価" in resolved["_dropped_label"]
