# #17「同じ依頼の 2 回目が別結果」── 設計 DESIGN-20260824-history-suggest.md の実装。
#
# 実測した事故（盲検の使い勝手レビュー）: 同じブックに一字一句同じ依頼を 2 回投げると、
# 1 回目は EXIT=0 で通り、2 回目は「頼める操作の一覧に照合できませんでした」EXIT=3。
# ★ 原因は LLM のサンプリングの揺れ。直す対象は揺れそのものではない
#   （temperature を 0 にしても別の入力で揺れる）。直すのは「揺れたときに人が困る」方。
#
# 契約（実装前に凍結した誤爆の条件）:
#   ① 一致は **依頼文の完全一致**のみ（部分一致で別の依頼に前回の op を当てない）
#   ② **成功した run のみ**を材料にする（失敗した run の op を勧めない）
#   ③ 材料が無ければ黙る（初回・別 PC で挙動が変わらない）
#   ④ PLAN（複数段）は 1 op に決まらないので勧めない
#   ⑤ 渡された並び順に依存しない（read_history は新しい順・ファイルは古い順）
#   ⑥ 履歴に op が**機械の値として**残る（表示テキストの解析で復元しない）
#   ⑦ 出す時は必ず根拠（日付）を言う ── 黙って前回の op を実行しない

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _h(task, ok, op, ts, **extra):
    return dict({"task": task, "ok": ok, "op": op, "ts": ts}, **extra)


# --- ①〜⑤ 引き当ての規則 ---------------------------------------------------------

def test_recalls_the_op_that_worked():
    h = [_h("請求書を作って", True, "SORT", "2026-08-20T10:00:00+00:00")]
    assert ailine.op_that_worked_before("請求書を作って", h) == ("SORT", "2026-08-20")


def test_ignores_failed_runs():
    """② 失敗した run の op を勧めない。"""
    h = [_h("x", False, "DEDUP", "2026-08-23T10:00:00+00:00")]
    assert ailine.op_that_worked_before("x", h) == (None, None)


def test_requires_exact_task_match():
    """① 部分一致で拾わない（別の依頼に前回の op を当てる事故を作らない）。"""
    h = [_h("金額で並べ替えて", True, "SORT", "2026-08-20T10:00:00+00:00")]
    assert ailine.op_that_worked_before("金額で並べ替えて、太字にして", h) == (None, None)
    assert ailine.op_that_worked_before("並べ替えて", h) == (None, None)
    # 前後の空白だけは吸収する（同じ依頼と見なしてよい範囲）
    assert ailine.op_that_worked_before("  金額で並べ替えて  ", h)[0] == "SORT"


def test_silent_without_material():
    """③ 履歴が無ければ黙る（初回・別 PC で挙動が変わらない）。"""
    assert ailine.op_that_worked_before("x", []) == (None, None)
    assert ailine.op_that_worked_before("x", None) == (None, None)
    assert ailine.op_that_worked_before("", [_h("", True, "SORT", "2026-08-20T00:00:00+00:00")]) \
        == (None, None)


def test_plan_is_never_suggested():
    """④ PLAN は 1 op に決まらない。"""
    h = [_h("x", True, "PLAN", "2026-08-23T10:00:00+00:00")]
    assert ailine.op_that_worked_before("x", h) == (None, None)


def test_picks_the_newest_regardless_of_input_order():
    """⑤ 並び順に依存しない ── 初版は reversed() で**最古の成功**を拾っていた
       （自分で用意した古い順の検体で試したので気づけなかった）。"""
    h = [_h("x", True, "SORT", "2026-08-20T10:00:00+00:00"),
         _h("x", True, "DEDUP", "2026-08-23T10:00:00+00:00")]
    assert ailine.op_that_worked_before("x", h) == ("DEDUP", "2026-08-23")
    assert ailine.op_that_worked_before("x", list(reversed(h))) == ("DEDUP", "2026-08-23")


def test_broken_entries_do_not_crash():
    h = ["not a dict", None, {"task": "x"}, _h("x", True, "SORT", "2026-08-20T00:00:00+00:00")]
    assert ailine.op_that_worked_before("x", h)[0] == "SORT"


# --- ⑥ 履歴に op が機械の値として残る -----------------------------------------------

def test_history_entry_carries_the_op_as_data():
    entry = ailine.build_history_entry(
        {"ok": True, "op": "SORT", "path": "dsl"}, Path("b.xlsx"), "並べ替えて",
        "qwen2.5-coder:7b", failure_kind="")
    assert entry["op"] == "SORT", "履歴に op が残らない（解釈行の解析に逆戻りする）"


def test_history_entry_carries_plan_steps():
    entry = ailine.build_history_entry(
        {"ok": True, "op": "PLAN", "ops": ["SORT", "BOLD"], "path": "plan"},
        Path("b.xlsx"), "並べ替えて太字に", "m", failure_kind="")
    assert entry["ops"] == ["SORT", "BOLD"]


def test_all_result_sites_record_the_op():
    """★ 同じ result dict を 4 箇所が組んでいる（片配線の温床）── 全部に載っていること。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    dsl_sites = src.count('"path": "dsl", "command": confirm.line')
    with_op = src.count('"op": op,')
    assert dsl_sites == 3, f"dsl の組み立てが 3 箇所でない（{dsl_sites}）── 検体の前提が古い"
    assert with_op >= 3, f"op を載せていない組み立てが残っている（{with_op}/3）"
    assert '"op": "PLAN"' in src, "plan 経路に op が無い"
