# -*- coding: utf-8 -*-
"""合格率の採り方が、決めた線どおりであること（2026-09-20）。

★★ 出所（Namakoo 決裁）:「一回で合格はかえって疑わしさが残った状態だ。ただ全勝しても
  N+1 回目の動作保証は出来ないから、**合格率で判定しよう**」。
  理由は本人の使い方 ──「俺の性格だと完成後に結局複数回テストすると思うんだよね」。

★★ この試験が守るのは 3 つ:
  ① **率を当てる先は「到達」だけ** ── 嘘の到達（✓ が付いたのに中身が違う）には
     率を当てない。1 度でも出たら不合格（取り返しが非対称なものに平均を当てない）。
  ② **「100%」と書かない** ── 30/30 で言えるのは「真の成功率 90% 以上（片側95%）」まで。
     観測していないことを主張しない、というこの repo の一番古い線そのもの。
  ③ **N は 30**（Namakoo 決裁）── 30/30 が「9 割は通る」と言い切れる最小の回数。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

import blind_session_core as core  # noqa: E402

DOC = REPO / "docs" / "FROZEN-20260917-盲検の合格線.md"


@pytest.mark.parametrize("k, n, want", [
    (30, 30, 90), (20, 20, 86), (10, 10, 74), (9, 10, 61), (19, 20, 78), (29, 29, 90),
])
def test_the_lower_bound_matches_what_we_told_namakoo(k, n, want):
    """★★ 下限の数が、Namakoo に見せた表と一致すること。

    ★ この表を根拠に N=30 を選んでもらった ── 数が動いたら、決裁の前提が動く。
    ★ 外の道具に頼らず自前で解いているので、ここが唯一の校正点。
    """
    got = round(core.lower_bound(k, n) * 100)
    assert got == want, f"{k}/{n} の下限が {got}%（表は {want}%）"


def test_a_perfect_run_is_never_reported_as_a_hundred_percent():
    """★★ 全勝を「100%」と書かないこと ── 観測していないことを主張しない。"""
    line = core.rate_line(30, 30)
    assert "30/30" in line, line
    assert "100%" not in line, f"★ 全勝を 100% と書いている: {line}"
    assert "以上" in line, f"★ 下限であることが書かれていない: {line}"


def test_the_rate_never_softens_a_silent_wrong_answer():
    """★★ 嘘の到達の影には率を当てない ── 1 件でも在れば不合格。

    ★ ここが緩むと「10 回に 1 回だけ静かに間違う道具」が、9 割通ることを根拠に合格する。
    """
    ok = [{"task": "A", "runs": 30, "reached": 30, "landed": ["並べ替え"],
           "two_answers": False, "screens": []} for _ in range(9)]
    bad = {"task": "B", "runs": 30, "reached": 30, "landed": ["並べ替え", "行削除"],
           "two_answers": True, "screens": []}
    v = core.verdicts(ok + [bad])
    assert v["failed_by_two_answers"], "★ 嘘の到達の影を率で薄めている"
    assert v["two_answers"] == ["B"], v
    # ★ 逆側: 影が無ければ、率が低くてもここでは落とさない（率は記録であって閾値でない）
    low = [{"task": "C", "runs": 30, "reached": 3, "landed": ["並べ替え"],
            "two_answers": False, "screens": []}]
    assert not core.verdicts(low)["failed_by_two_answers"]


def test_two_answers_is_derived_from_what_actually_landed():
    """★ 「2 つの答えで通った」は、**通った回の着地先**から導くこと（人が書かない）。

    ★★ 検体の作り方（2026-09-20・変異試験が初版の弱さを指した）: 初版は通った回を
      2 つの op に分けていたので、**断った回を混ぜる変異でも `two_answers` が True のまま**
      緑を通した。いまは逆に作る ── **通った回は 1 つの op だけ**、断った回だけが別の語を
      持つ。混ぜた瞬間に「2 つの答え」になるので、そこで赤くなる。
    """
    calls = []

    def fake(book, task, home):
        calls.append(1)
        if len(calls) % 3 == 0:                      # ★ 断った回
            return {"rc": 3, "landed": "★断りの語", "tail": "？ 断り"}
        return {"rc": 0, "landed": "並べ替え", "tail": "✓"}   # ★ 通った回は 1 つだけ

    req = {"task": "T", "book": "dummy.xlsx", "book_name": "x.xlsx"}
    import shutil
    real_copy = shutil.copy2
    shutil.copy2 = lambda *a, **k: None
    try:
        got = core.replay_one(REPO / "tests", req, 6, run_fn=fake)
    finally:
        shutil.copy2 = real_copy
    assert got["reached"] == 4, got
    assert got["landed"] == ["並べ替え"], f"★ 断った回を着地先に混ぜている: {got['landed']}"
    assert got["two_answers"] is False, got


def test_the_decision_is_written_in_the_frozen_document():
    """★★ 決裁が文書に在ること ── 道具と文書がずれたら、どちらかが嘘になる。

    ★ 数（30）は**文書と道具の両方**に在るので、両方向で突き合わせる。
    """
    doc = DOC.read_bytes().decode("utf-8")
    assert "合格率で判定しよう" in doc, "★ 決裁の言葉が文書に無い"
    assert f"**{core.DEFAULT_RUNS} 回**" in doc, (
        f"★ 文書と道具で回数がずれている（道具は {core.DEFAULT_RUNS}）")
    assert "率を当てない ── 1 度でも出たら不合格" in doc, "★ 嘘の到達の扱いが文書に無い"
    assert "90%" in doc, "★ 30/30 で言える下限が文書に無い"


def test_the_tool_does_not_change_the_blind_protocol():
    """★ 買い手に見せるものを増やしていないこと（盲検の作法は動かさない）。

    ★ 道具が触るのは「環境の用意」と「終わった後の記録」だけ。
    """
    src = (REPO / "tests" / "blind_session_core.py").read_bytes().decode("utf-8")
    assert "買い手に渡すのは README だけ" in src, "★ 作法が道具に書かれていない"
    for forbidden in ("docs/", "tests/", "bench/"):
        assert f"買い手に{forbidden}" not in src
