# 初日・素の環境（2026-08-24 の盲検査定・最も痛い指摘の第 1 位）。
#
# ★ 実測: openpyxl が入っていない環境で `ailine` を起動すると
#   `NameError: name 'exit_environment' is not defined` という生の traceback が出ていた。
#   import ガード（64 行目）が、まだ定義されていない関数を呼んでいたため。
#   ── 「足りないものを名指しする」ための `ailine doctor` にすら到達できない。
#   ★ 15 日間、監査の波も盲検レビューも pre-push の番人も、これを一度も捕まえなかった。
#     開発機には openpyxl が**常に在る**から（居るから見えない）。
#
# 契約:
#   ① openpyxl が無い環境でも、生の traceback でなく 1 行の案内と exit 9 で落ちる
#   ② 案内には「何が要るか」と「どう入れるか」が両方入っている
#   ③ 依存を呼ぶ前に、その案内を出す関数が**定義済み**であること（構造で縛る）

import ast
import os
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src" / "ailine" / "__init__.py"


def _run_without(module_name: str, *argv):
    """指定モジュールを import できない子プロセスで ailine を起動する。"""
    blocker = (
        "import sys\n"
        "class Block:\n"
        "    def find_spec(self, name, path=None, target=None):\n"
        f"        if name.split('.')[0] == {module_name!r}:\n"
        "            raise ImportError('blocked for the test')\n"
        "        return None\n"
        "sys.meta_path.insert(0, Block())\n"
        "import runpy\n"
        "runpy.run_module('ailine', run_name='__main__')\n"
    )
    env = dict(os.environ)
    env["PYTHONPATH"] = str(REPO / "src")
    return subprocess.run([sys.executable, "-c", blocker, *argv],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=env, timeout=180)


def test_missing_openpyxl_gives_guidance_not_a_traceback():
    r = _run_without("openpyxl", "doctor")
    combined = (r.stdout or "") + (r.stderr or "")
    assert "Traceback" not in combined, f"生の traceback が出た: {combined[-500:]}"
    assert "NameError" not in combined, combined[-500:]
    assert r.returncode == 9, f"環境エラーの exit 9 でない: {r.returncode} / {combined[-300:]}"


def test_the_guidance_says_what_and_how():
    r = _run_without("openpyxl", "doctor")
    combined = (r.stdout or "") + (r.stderr or "")
    assert "openpyxl" in combined, combined
    assert "pip install" in combined, f"入れ方が書いていない: {combined}"


def test_the_bail_out_helper_is_defined_before_it_is_used():
    """③ 構造で縛る ── 依存の import ガードより前に exit_environment が定義済みであること。

    ★ 実行時のテスト（上の 2 本）だけだと、将来また import ガードを増やしたときに
      同じ順序事故が起きうる。ソースの構造そのものを見る。
    """
    tree = ast.parse(SRC.read_text(encoding="utf-8"))
    def_line = next(n.lineno for n in tree.body
                     if isinstance(n, ast.FunctionDef) and n.name == "exit_environment")
    use_lines = [n.lineno for n in ast.walk(tree)
                  if isinstance(n, ast.Call) and isinstance(n.func, ast.Name)
                  and n.func.id == "exit_environment"]
    module_level_uses = [ln for ln in use_lines if ln < def_line]
    assert not module_level_uses, (
        f"exit_environment を定義（{def_line} 行目）より前で呼んでいる: {module_level_uses} 行目 "
        "── その依存が無い環境では NameError の生 traceback になる")
