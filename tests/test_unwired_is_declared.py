# -*- coding: utf-8 -*-
"""未到達（在るのに誰も呼んでいないもの）が、理由なしに増えないこと（2026-09-16）。

★ なぜ在るか（Namakoo「構築したグラフから逆に未配線などは見つけられたりはしない？」）:
  `scripts/deps_graph.py` は**在る配線**を描く。同じ材料で**在るべきなのに無い配線**も
  数えられる。初回の走査で本物を 1 件掴んだ ── `attributes.render_no_evidence`（候補 0 件の
  ときの断り文）が三兄弟のうち 1 つだけ配線されていなかった。

★★ 大前提: **到達不能＝宝ではない**。別のプロジェクトで 833 行の未配線モジュールを
  「宝」と呼び、実際は断念済みだった事故がある。だからこの番人は
  **消させない・直させない**。やらせるのは「**分類して理由を書く**」ことだけ。

★ 台帳の作り方は、同じ日に作った被覆の台帳（tests/op_completeness_register.json）と揃えた ──
  **見つけた瞬間に書かせる**（決めた時に書く方式にすると、次に増えた分が黙って漏れる）。
"""
import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
TOOL = REPO / "scripts" / "unwired.py"
REGISTER = Path(__file__).resolve().parent / "unwired_register.json"
KINDS = {"by_design", "leftover", "unwired"}
BUCKETS = ("orphan_modules", "functions_named_nowhere", "functions_only_tests_name",
           "rosters_no_one_reads", "basic_arms_unreached")


def _survey() -> dict:
    r = subprocess.run([sys.executable, str(TOOL), "--json"], cwd=str(REPO),
                        capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr[-800:]
    return json.loads(r.stdout)


def _found() -> set:
    got = _survey()
    return {x for b in BUCKETS for x in got[b]}


def _declared() -> dict:
    return json.loads(REGISTER.read_bytes().decode("utf-8"))["declared"]


def test_the_survey_finds_something():
    """★ 空回りの検出 ── 走査が壊れて 0 件になったら、下の試験は全部素通りする。"""
    assert len(_found()) >= 10, "未到達の走査が何も見つけていない（道具が壊れている疑い）"


def test_every_unreached_thing_is_declared():
    """★ 新しい未到達が現れたら、分類と理由を書かせる（それまでは赤）。"""
    missing = sorted(_found() - set(_declared()))
    assert not missing, (
        f"未到達なのに tests/unwired_register.json に宣言が無い: {missing}。"
        " kind（by_design / leftover / unwired）と why（調べた上での理由。"
        "分からなければ『未調査』）を書くこと ── ★ 消す判断はここではしない")


def test_no_declaration_outlives_its_subject():
    """★ 腐り防止 ── 配線された（または消えた）のに宣言が残っていたら赤。"""
    stale = sorted(set(_declared()) - _found())
    assert not stale, (
        f"宣言に在るが、もう未到達ではない: {stale}（配線されたか消えたはず）。"
        " tests/unwired_register.json から消すこと")


def test_declarations_have_a_known_kind_and_a_reason():
    bad = [k for k, v in _declared().items()
           if v.get("kind") not in KINDS or not str(v.get("why", "")).strip()]
    assert not bad, f"kind が未知、または why が空の宣言: {sorted(bad)}（許される kind: {sorted(KINDS)}）"


def test_the_tool_warns_that_unreachable_is_not_treasure():
    """★★ 道具の口上そのものを縛る ── ここが消えると、次に読む者が『宝を見つけた』と読む。

    ★ 2026-09-16 時点で `unwired` は 1 件だけ。数が少ないうちほど「全部直そう」に
      傾きやすいので、警告は道具の側に据え置く。

    ★★ 2026-09-16: 初版はソースの中に語が在るかだけを見ていた。ところが語は 2 か所に在り
      （docstring と画面へ出す行）、変異で docstring だけ消しても**緑のまま**だった ──
      番人が**綴り**を見て**出力**を見ていなかった。同じ日に軸の番人で踏んだのと同じ形。
      ★ だから道具を**実際に走らせて**、人の目に届く行を確かめる。
    """
    r = subprocess.run([sys.executable, str(TOOL)], cwd=str(REPO),
                        capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr[-400:]
    assert "到達不能＝宝ではない" in r.stdout, (
        "道具が『到達不能＝宝ではない』を画面に出していない ── "
        "次に読む者が『宝を見つけた』と読む")
    reg = json.loads(REGISTER.read_bytes().decode("utf-8"))
    assert "宝ではない" in reg["_meta"].get("warning", ""), "台帳から同じ断りが消えた"


def test_the_tool_states_that_it_counts_names_not_calls():
    """★ 限界の明記 ── 『名前の出現』で数えている以上、0 件は『呼ばれていない』ではない。"""
    src = TOOL.read_bytes().decode("utf-8")
    assert "名前の出現" in src, "数え方の限界が道具から消えた（動的な呼び出しは追えない）"


def test_the_real_unwired_ones_are_few_and_named():
    """★ `unwired`（本物の候補）は**名指しで**持つ ── 数だけ見て安心しない。

    ★ 増えること自体は悪くない（見つかった証拠）。悪いのは、増えたのに
      理由が書かれないこと。それは上の試験が捕まえる。
    """
    real = {k: v for k, v in _declared().items() if v["kind"] == "unwired"}
    assert real, "本物の未配線が 0 件 ── 走査か分類のどちらかを疑うこと"
    for k, v in real.items():
        assert len(v["why"]) >= 40, f"{k} の理由が短すぎる（何を確かめたかを書く）"
