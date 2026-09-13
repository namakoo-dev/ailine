# -*- coding: utf-8 -*-
"""後からの独立検算（分けた冊・科目の候補・帳票の一覧）は、1 つの口で話す（2026-09-13）。

★★ なぜ要るか: 3 経路が `5 if result.get("mismatch") else 0` を**書き写していた**。
  片配線の足場そのもの（`docs/開発手法.md` §13）── 実際 B8 の直し（0 件照合で ✓ を出さない）
  は、番号の側を 3 箇所直さなければ効かない形だった。
★ 処方は「3 箇所直す」でなく「1 つの関数に畳んで、呼び出し側に判断を持たせない」。
  ここは**その配線を機械で縛る** ── 報告を出す口の数と、番号を決める口の数を突き合わせる。
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core.cli_render import render_independent_verify_report          # noqa: E402

AILINE_PY = REPO / "src" / "ailine" / "__init__.py"


def _calls(name: str) -> int:
    tree = ast.parse(AILINE_PY.read_text(encoding="utf-8"))
    return sum(1 for n in ast.walk(tree)
               if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id == name)


def test_every_report_of_this_kind_takes_its_exit_code_from_one_place():
    reports = _calls("render_independent_verify_report")
    decisions = _calls("_independent_verify_exit")
    assert reports >= 3, f"この器を使う経路が {reports} しかない（分母が痩せている）"
    assert decisions >= reports, (
        f"報告 {reports} 経路に対して番号を決める口が {decisions} しかない ── "
        "どこかが `5 if mismatch else 0` を書き写している（片配線）")


def test_no_route_writes_the_verdict_by_hand():
    """★ 変異の裏返し ── 手書きの `5 if ... else 0` が 1 つでも戻ったら赤。"""
    text = AILINE_PY.read_text(encoding="utf-8")
    # ★ 先に「探す場所が空でない」ことを確かめる（`not in` は空なら必ず通る）。
    assert text.count("_independent_verify_exit(result)") >= 3, (
        "畳んだ関数を通す経路が 3 本未満 ── 分母が痩せている（この試験は無意味になる）")
    assert '5 if result.get("mismatch") else 0' not in text, (
        "番号を手書きしている行がある ── `_independent_verify_exit` を通すこと")


def test_a_vacuous_verification_never_prints_the_check_mark():
    """★★ 0 件照合で合格を名乗らない（B8）── 器の側で縛る（3 つの検算器すべてに効く）。"""
    facts = {"一覧の行": 2, "含有を確かめた値": 0}
    lines = render_independent_verify_report(
        "帳票の一覧", "一覧.xlsx", "受領",
        {"breaks": [], "facts": facts, "mismatch": False,
         "vacuous": "含有を確かめた値が 0 件です"})
    body = "\n".join(lines)
    assert "✓" not in body, body
    assert "合格でも不合格でもありません" in body, body


def test_a_real_verification_still_prints_the_check_mark():
    """★ 陰性対照 ── 測って破れが無かった回は今までどおり ✓。"""
    lines = render_independent_verify_report(
        "帳票の一覧", "一覧.xlsx", "受領",
        {"breaks": [], "facts": {"含有を確かめた値": 270}, "mismatch": False})
    assert any(ln.startswith("✓") for ln in lines), lines
