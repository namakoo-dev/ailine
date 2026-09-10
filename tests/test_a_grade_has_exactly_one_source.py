# -*- coding: utf-8 -*-
"""区分（確 / 単 / 割 / 無）は 1 箇所からしか作れない ── **実装より先に置く番人**。

★★ なぜ実装の前に置くか（2026-09-11）:

  次に作る「N 冊の帳票から項目を集める」機能は、値と一緒に**区分**を出す:

      確   独立した根拠が 2 つ以上あり、一致した
      単   根拠が 1 つだけ（値は出すが、裏が取れていない）
      割   根拠が 2 つ以上あり、食い違った（値を出さず、両側の数字を見せる）
      無   根拠が無い（取れなかった。理由を名指しする）

  この区分は 4 つの出口（画面・出力ブックのセル・検分シート・--json）に現れる。
  ★ 出口ごとに文字列を組み立てると、**片方だけ直る事故**の足場になる ──
    この repo が何度も踏んだ形（docs/開発手法.md §13）。

  ★ 処方は「両方直す」でなく「**1 関数に畳んで呼び出し側に判断を持たせない**」。
    番人も「1 本の試験で全経路を縛る」形にする。

★ 独立に設計を書いた 3 人が、全員この番人を「実装と同時に置くもの」に挙げた。
  そのうち 1 人は「**先に書け（この時点では赤くてよい）**」と書いた。
  ここでは**赤くならない形**で先に置ける ── 下の test_one 参照。

★ いま（実装前）の状態: 単独リテラルは src/ 全体で 0 件。
  だからこの番人は今日から緑で、実装が 2 箇所目を作った瞬間に赤くなる。
"""
from __future__ import annotations

import ast
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src"

#: 区分の語。★ ここに足すときは、下の 2 つの試験の意味が変わることを承知の上で。
GRADES = {"確", "単", "割", "無"}

#: 区分を作ってよい唯一の場所（実装が入ったらここになる）。
#: ★ まだ存在しない。存在しない間は「0 箇所」で緑（下の試験が両向きに縛る）。
#: ★ 走査は repo 相対で返す（src/ を含む）。ここも同じ形にする ──
#:   形が違うと「正しく書いたのに通らない」番人になる（変異試験の陽性対照で捕まえた）。
GRADE_HOME = "src/ailine_core/field_record.py"


def _literal_sites() -> list[tuple[str, int, str]]:
    """src/ 全体で、区分の語**そのもの**が文字列リテラルとして現れる場所。

    ★ 部分一致では数えない（「確認」「単価」「割合」「無効」に当たる）。
      **値が区分の語と完全一致する Constant だけ**を数える。
    """
    out: list[tuple[str, int, str]] = []
    for p in sorted(SRC.rglob("*.py")):
        try:
            tree = ast.parse(p.read_bytes().decode("utf-8"))
        except SyntaxError:  # pragma: no cover ── 壊れた .py は他の番人の仕事
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.Constant) and isinstance(node.value, str) \
                    and node.value in GRADES:
                rel = p.relative_to(REPO).as_posix()
                out.append((rel, node.lineno, node.value))
    return out


def test_one_module_may_build_a_grade() -> None:
    """① 区分の語を組み立てるのは 1 モジュールだけ。

    ★ 呼び出し側（画面・出力・検分・json）が自分で "確" を書けないようにする。
      書けてしまうと、導出の規則を直しても出口の 1 つが古い判断のまま残る。
    """
    sites = _literal_sites()
    modules = sorted({m for m, _, _ in sites})
    assert len(modules) <= 1, (
        f"区分の語を作っている場所が {len(modules)} 箇所ある: {modules}\n"
        f"★ 区分は {GRADE_HOME} の導出関数だけが作る。呼び出し側は受け取って表示するだけ。\n"
        f"  出現: {sites[:12]}")
    if modules:
        assert modules[0] == GRADE_HOME, (
            f"区分を作っているのが {modules[0]} ── {GRADE_HOME} のはず")


def test_the_home_is_named_and_not_yet_there_or_actually_there() -> None:
    """② ★ 「0 箇所だから緑」が居座らないようにする（負の被覆）。

    ★ この番人は実装前から置ける ── だが「まだ無いから緑」は**守っていない**のと同じ。
      だから在否を明示的に測り、**在るなら区分を作っていること**まで確かめる。
      （skip で流すと、実装が入った日に誰も気づかない）
    """
    home = REPO / GRADE_HOME
    sites = _literal_sites()
    if not home.exists():
        # 実装前: 区分を作っている場所は 0 でなければならない
        assert not sites, (
            f"{GRADE_HOME} がまだ無いのに、区分の語を作っている場所がある: {sites[:8]}\n"
            "★ 区分の家より先に区分が生えている。置き場所を決めてから作ること")
        return
    # 実装後: 家が在るなら、家が区分を作っていること（空の家を残さない）
    in_home = [s for s in sites if s[0] == GRADE_HOME]
    assert in_home, (
        f"{GRADE_HOME} は在るのに、そこで区分の語を作っていない\n"
        "★ 家だけ在って中身が別の場所に移った可能性がある（片配線）")


def test_the_grades_are_spelled_the_same_everywhere_in_this_guard() -> None:
    """③ この番人自身の語が、設計文書と食い違っていないこと。

    ★ 番人が古い語を見張ると、静かに何も守らなくなる（「在っても鳴らない」の一形）。
    """
    doc = REPO / "docs" / "DESIGN-20260910-フォルダから語のマップを作る.md"
    if not doc.exists():
        pytest.skip("設計文書が無い（測れない回は skip と書く）")
    text = doc.read_text(encoding="utf-8")
    missing = [g for g in sorted(GRADES) if g not in text]
    assert not missing, (
        f"この番人が見張っている語 {missing} が設計文書に出てこない ── "
        "語が変わったか、番人が古い")
