"""bench の道具は本物の ~/.ailine に書かない（2026-09-24）。

★★ なぜ在るか（実害）: 効果の行列（bench/basic_ops_matrix.py）を pytest の外で直に流したら、本物の
  ~/.ailine/history.jsonl に 246 行を書いた。試験は conftest が切り離すが、bench の道具は誰も
  切り離していなかった（製品を動かす道具 19 本すべて）。入口を bench/_bench_home.py の 1 つにした。
★ 名簿は手で書かない: 「製品を動かす道具」は本文から導く（`-m ailine` を呼ぶ／ailine を import する）。
★ 縛るのは**位置**まで: isolate() は ailine を import する前（import の時に保存先が決まる）。
"""
import ast
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
BENCH = REPO / "bench"
RUNS = re.compile(r'"-m",\s*"ailine"|\bimport ailine\b|\bfrom ailine import\b')

#: ★ 理由つきの免除（30 字以上・句点あり）
EXEMPT = {
    "verify_golden_probe.py":
        "pytest の plugin として conftest より先に読み込まれる。ここで AILINE_HOME を触ると、"
        "conftest の本物のホームの検算が一時フォルダの方を見張ってしまう。pytest の中は conftest が切り離す。",
}


def _tools():
    return [p for p in sorted(BENCH.glob("*.py"))
            if not p.name.startswith("_") and RUNS.search(p.read_bytes().decode("utf-8"))]


def _isolates_first(src: str) -> bool:
    """モジュールの最上位で、ailine を import する文より前に `_bench_home.isolate()` を呼んでいるか。"""
    for n in ast.parse(src).body:
        if (isinstance(n, ast.Expr) and isinstance(n.value, ast.Call)
                and ast.unparse(n.value.func) == "_bench_home.isolate"):
            return True
        if isinstance(n, (ast.Import, ast.ImportFrom)):
            names = [a.name for a in n.names] if isinstance(n, ast.Import) else [n.module or ""]
            if any(x == "ailine" or x.startswith("ailine.") for x in names):
                return False
    return False


def test_every_tool_that_runs_the_product_isolates_first():
    tools = _tools()
    assert len(tools) >= 10, f"製品を動かす道具が {len(tools)} 本しか見つからない ── 名簿の導き方が壊れている疑い"
    bad = [p.name for p in tools if p.name not in EXEMPT
           and not _isolates_first(p.read_bytes().decode("utf-8"))]
    assert not bad, ("本物の ~/.ailine に書きうる bench の道具（先頭で _bench_home.isolate() を呼んでいない）:\n  "
                     + "\n  ".join(bad))


def test_the_exemptions_are_current_and_explained():
    names = {p.name for p in _tools()}
    stale = sorted(set(EXEMPT) - names)
    assert not stale, f"もう製品を動かさない道具が免除に残っている: {stale}"
    for k, why in EXEMPT.items():
        assert len(why) >= 30 and "。" in why, k


def test_the_detector_catches_a_tool_that_forgot():
    """★ 変異: isolate を呼ばない形・ailine の後に呼ぶ形を拾う。正しい形は拾わない。"""
    assert not _isolates_first("import ailine\n")
    assert not _isolates_first("import ailine\nimport _bench_home\n_bench_home.isolate()\n")
    assert _isolates_first('"""d"""\nimport _bench_home\n_bench_home.isolate()\nimport ailine\n')


def test_isolate_really_moves_the_home():
    """★ 陽性対照: 素の子プロセスで isolate() の後に ailine を import すると、保存先が本物のホームの外。"""
    code = ("import os, sys; os.environ.pop('AILINE_HOME', None); sys.path[:0] = [sys.argv[1], sys.argv[2]]\n"
            "import _bench_home; d = _bench_home.isolate()\n"
            "import ailine; from pathlib import Path\n"
            "real = Path.home() / '.ailine'\n"
            "p = Path(ailine.HISTORY_FILE)\n"
            "print('OUT' if real not in p.parents and Path(d) in p.parents else 'IN', p)\n")
    env = {k: v for k, v in __import__("os").environ.items() if k != "AILINE_HOME"}
    r = subprocess.run([sys.executable, "-c", code, str(BENCH), str(REPO / "src")],
                       capture_output=True, text=True, encoding="utf-8", env=env, timeout=120)
    assert r.returncode == 0, r.stderr[-800:]
    assert r.stdout.startswith("OUT"), f"切り離されていない: {r.stdout}"
