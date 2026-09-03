# 依存の図が実体と一致していること（2026-09-03）。
#
# ★★ なぜ在るか: この repo は「人が書いた数は必ず古くなる」を何度も踏んでいる ──
#   README の行数（17,671 のまま残った）・試験数（3,085 のまま）・翻訳精度
#   （98.1% と書いて実測は 94.2% だった）。★ **図はもっと古くなりやすい**（見た目が
#   それらしいので、ずれていても気づかない）。だから手で描かず、実体から生成して
#   ここで一致を守る。
#
# 契約:
#   ① docs/依存関係.md が `scripts/deps_graph.py` の出力と 1 バイトも違わない
#   ② 層の向きの違反（ailine_core → ailine）がゼロ
#      ★ core が本体を知らないから、本体だけを差し替えられる。ここが破れると
#        「部品を取り出して別の入口から使う」ができなくなる
#   ③ ★ 図が「見えないもの」を明記していること
#      （辞書経由の呼び出しは辺に出ない ── 読む人が「全部見えている」と誤解しないため）

import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DOC = REPO / "docs" / "依存関係.md"
GEN = REPO / "scripts" / "deps_graph.py"


def _generated() -> str:
    r = subprocess.run([sys.executable, str(GEN)], cwd=str(REPO),
                        capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr[-800:]
    return r.stdout.replace("\r\n", "\n")


def test_the_document_matches_the_machine():
    """① 図が実体と一致していること。

    ★ 直し方: `python scripts/deps_graph.py --write` で作り直し、
      **git diff を読んでから** commit する（増えた辺・消えた辺が意図どおりか）。
    """
    assert DOC.exists(), f"図が無い: {DOC}（scripts/deps_graph.py --write で生成）"
    have = DOC.read_bytes().decode("utf-8").replace("\r\n", "\n")
    want = _generated()
    assert have.strip() == want.strip(), (
        "依存の図が実体とずれている ── `python scripts/deps_graph.py --write` で"
        "作り直し、git diff で増減を確かめてから commit すること")


def test_the_core_does_not_import_the_main_module():
    """② 層の向き ── ailine_core は本体を知らない。

    ★ ここが破れると「部品を取り出して別の入口から使う」ができなくなる。
      分割の意味そのものが失われるので、図の一致とは別に単独で見る。
    """
    import ast
    src = REPO / "src"
    bad = []
    for p in sorted((src / "ailine_core").rglob("*.py")):
        tree = ast.parse(p.read_bytes().decode("utf-8"))
        for n in ast.walk(tree):
            if isinstance(n, ast.ImportFrom) and n.module == "ailine":
                bad.append(f"{p.name}:{n.lineno}")
            elif isinstance(n, ast.Import):
                bad += [f"{p.name}:{n.lineno}" for a in n.names if a.name == "ailine"]
    assert not bad, f"ailine_core が本体を import している: {bad}"


def test_the_document_states_what_it_cannot_show():
    """③ ★ 図が「見えないもの」を明記していること。

    ★ 静的な import グラフは**辞書経由の呼び出しを追えない**。ailine は
      POSTCONDITIONS 辞書で op → 事後条件を引くので、その 28 本は図の上で
      「本体から呼ばれていない」ように見える。
    ★ この注記が消えると、図が「全部見えている」という嘘をつく ──
      この repo でいちばん避けたい形（出ないことは信号でない）。
    """
    t = DOC.read_bytes().decode("utf-8")
    assert "POSTCONDITIONS" in t, "辞書経由が見えないことの注記が消えた"
    assert "辺が無い" in t and "影響が無い" in t, (
        "「辺が無い＝影響が無い ではない」の注記が消えた")
