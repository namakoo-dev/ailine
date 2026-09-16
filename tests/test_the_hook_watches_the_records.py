# -*- coding: utf-8 -*-
"""commit の瞬間の番人が、記録のずれを見ていること（2026-09-16）。

★ 出所（Namakoo「自動じゃなくてもフックで気付ける？ 実際の構成と図やグラフが
  一致してないとそれを追うのが不可能になる」）:
  記録の番人（行数・試験の本数・依存の図）は**元から在った**。居場所が全件の中、
  つまり **push の 30 分後**に鳴っていた。同じ日に 2 回それで push が止まっている。
  ★ 検出は失敗していない。**検出が遅い**ことが損だった ── 直すのはループの長さ
  （改行コードの番人を commit の瞬間へ出した 2026-09-15 と同じ処方）。

★★ この試験が在る理由: フックは**壊れても静かに通る**（`exit 0` を 1 行足せば黙る）。
  全件テストの中からはフックの中身が見えないので、ここで中身を縛る ──
  「在っても鳴らない」の予防。実弾（わざとずらして commit を試す）は
  tests/test_hooks_fire.py 側の作法に倣い、ここでは**中身の形**を見る。
"""
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
HOOK = REPO / "scripts" / "pre-commit.sh"
INSTALL = REPO / "scripts" / "install-hooks.sh"
TOOL = REPO / "scripts" / "refresh_records.py"


def _hook() -> str:
    return HOOK.read_bytes().decode("utf-8")


def test_the_hook_runs_the_record_check():
    """★ フックが記録の道具を呼んでいること。"""
    src = _hook()
    assert "refresh_records.py" in src, (
        "commit の瞬間に記録のずれを見ていない ── ずれは push の 30 分後まで気づけません")


def test_the_hook_does_not_repair_silently():
    """★★ 直さない（気づかせるだけ）。

    ★ commit の最中にファイルを書き換えると、staged と working tree がずれる ──
      「commit したものと違うものが記録される」という、いちばん追いにくい壊れ方になる。

    ★ 初版は「直す: … --write」という**案内文**を呼び出しと数えて自分で赤くなった。
      見るのは**実際に走らせている行**だけ（echo / printf の中は画面へ出す文字列）。
    """
    calls = [ln for ln in _hook().splitlines()
             if "refresh_records.py" in ln
             and not re.match(r"\s*(#|echo|printf)", ln)]
    assert calls, "記録の道具を走らせている行が見つからない"
    for c in calls:
        assert "--write" not in c, (
            f"フックが記録を書き換えている: {c.strip()} ── 気づかせるだけにすること")


def test_the_hook_scopes_the_check_by_what_changed():
    """★ 費用で外されないこと ── 変わったファイルで絞る。

    ★ 全部見ると毎 commit 10 秒（行数 0.1 / 図 3.4 / 試験の本数 6.8）。
      重い番人は「いずれ外される」── 外された番人は在っても鳴らない。
    """
    src = _hook()
    assert "--parts" in src, "見る範囲を絞っていない（毎 commit 10 秒は外される）"
    assert "git diff --cached --name-only" in src, "何が変わったかを見ていない"
    for part in ("lines", "tests", "graph"):
        assert part in src, f"絞り方に {part} が無い"


def test_the_hook_tells_you_how_to_fix_it():
    """★ 止めるだけでは仕事が進まない ── 直す一行を画面に出すこと。"""
    src = _hook()
    assert "refresh_records.py --write" in src, "直し方を出していない"
    assert "--no-verify" in src, "意図的に通す道を書いていない"


def test_the_hook_says_what_it_cannot_fix():
    """★★ 測定の記録には触れないことを、その場で言うこと。

    ★ ここが無いと、赤を見た人が `--write` を打って「直らない」と混乱する。
      実機を回して人が書く、という線を画面で伝える。
    """
    src = _hook()
    assert "battery_recorded.json" in src, "測定の記録の直し方を案内していない"
    assert "実機" in src


def test_the_installer_installs_this_hook():
    """★ 入れる道具が、この番人を入れること（入っていなければ在るだけ）。"""
    src = INSTALL.read_bytes().decode("utf-8")
    assert "pre-commit" in src, "install-hooks.sh が pre-commit を入れていない"


def test_the_check_tool_exits_nonzero_when_stale():
    """★ 空回りの検出 ── 道具がずれを非ゼロで返さなければ、フックは永久に緑。"""
    budget = REPO / "tests" / "ailine_py_line_budget.txt"
    before = budget.read_bytes()
    budget.write_bytes(b"1\n")
    try:
        r = subprocess.run([sys.executable, str(TOOL), "--parts", "lines"],
                            cwd=str(REPO), capture_output=True, text=True,
                            encoding="utf-8", errors="replace")
        assert r.returncode != 0, "ずれているのに 0 を返している（フックが鳴らない）"
    finally:
        budget.write_bytes(before)
