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


def test_no_relative_imports_inside_the_core():
    """④ ★ `ailine_core` の中で相対 import を使わない。

    ★★ なぜ番人が要るか（2026-09-11 に実測で踏んだ）:
      新しいモジュールを `from .field_record import ...` と相対で書いたところ、
      **図のモジュール数だけが増えて辺が増えなかった**（64→67 で辺は 84 のまま）。
      `scripts/deps_graph.py` は `ImportFrom.module` が "ailine" で始まる辺だけを
      拾うので、相対 import は `module="field_record"` になり**辺として消える**。
      → 図は「一致している」と言いながら、**依存を 2 本隠していた**。

    ★ これは「番人が在っても、その事故の形では鳴らない」の一例。
      図の一致（①）は通ってしまうので、書き方そのものをここで縛る。
      ついでに repo 全体の作法（絶対 import）とも揃う。
    """
    import ast
    bad = []
    for p in sorted((REPO / "src" / "ailine_core").rglob("*.py")):
        tree = ast.parse(p.read_bytes().decode("utf-8"))
        for n in ast.walk(tree):
            if isinstance(n, ast.ImportFrom) and n.level and n.level > 0:
                bad.append(f"{p.relative_to(REPO)}:{n.lineno} "
                           f"（{'.' * n.level}{n.module or ''}）")
    assert not bad, (
        "ailine_core で相対 import を使っている ── 依存が図から消えます。"
        f"絶対 import（from ailine_core.x import y）に直してください: {bad}")


# ---------------------------------------------------------------------------
# ⑤ ★★ import では見えない 2 層が、図に**在る**こと（2026-09-16）
#
# ★ なぜ足したか: この図はもともと「辞書経由の呼び出しと Basic 側は見えない」と
#   **自分で断っていた**。そして 2026-09-16 に出たバグ 3 件は、全部その 2 つに住んでいた ──
#     ・名簿 `_OPS_THAT_SKIP_NON_DATA_ROWS` が AGGREGATE を落とし、合計行が実際に消えた
#     ・`PROJECTIONS` が DEDUP を落とし、式が値に落ちることを誰にも告げなかった
#     ・`AXES` が AGGREGATE/PIVOT を落とし、平均を頼むと黙って合計が返った
#   ★ 「見えない」と断っただけでは、そこは永久に暗いまま。描けるものは描く。
# ★ 消えても図は「一致」で通ってしまう（節ごと無くなれば生成側も一致する）ので、
#   **節が在ること**と**中身が空でないこと**を別に縛る。
# ---------------------------------------------------------------------------

def test_the_layers_that_imports_cannot_see_are_drawn():
    """⑤-1 層②（op → 生成関数 → Basic の腕 → 事後条件）と層③（op の名簿）が在ること。"""
    t = DOC.read_bytes().decode("utf-8")
    assert "## 層② op の配線" in t, "op の配線の節が消えた（辞書経由の層がまた暗くなる）"
    assert "## 層③ op の名簿" in t, "op の名簿の節が消えた"
    assert "Basic の腕" in t, "Basic 側への配線が図から消えた"


def test_every_op_appears_in_the_wiring_table():
    """⑤-2 空回りの検出 ── 全 op が配線の表に出ていること。

    ★ 表が空でも、または 1 行でも、①の一致は通る。だから**分母**を別に縛る。
    """
    import sys as _sys
    _sys.path.insert(0, str(REPO / "src"))
    import ailine
    t = DOC.read_bytes().decode("utf-8")
    body = t.split("## 層② op の配線")[1].split("## 層③")[0]
    missing = [op for op in ailine.OP_SCHEMA if f"| `{op}` |" not in body]
    assert not missing, f"配線の表に出ていない op: {sorted(missing)}"


def test_the_roster_table_lists_every_op_and_at_least_one_roster():
    """⑤-3 名簿の表も全 op を持つこと（名簿が 0 本なら表として無意味）。"""
    import sys as _sys
    _sys.path.insert(0, str(REPO / "src"))
    _sys.path.insert(0, str(REPO / "tests"))
    import ailine
    from test_op_completeness import discover_op_rosters
    n = len(ailine.OP_SCHEMA)
    partial = [k for k, v in discover_op_rosters().items() if len(v["ops"]) < n]
    assert partial, "部分の名簿が 1 本も見つからない（走査が壊れている）"
    t = DOC.read_bytes().decode("utf-8")
    body = t.split("## 層③ op の名簿")[1]
    for k in partial:
        assert k.split(":")[-1] in body, f"名簿が表に出ていない: {k}"
    missing = [op for op in ailine.OP_SCHEMA if f"| `{op}` |" not in body]
    assert not missing, f"名簿の表に出ていない op: {sorted(missing)}"


def test_the_document_still_says_what_it_cannot_show():
    """⑤-4 ★ 描いた分だけ注記を縮めたが、**残りの限界**は必ず書いてあること。

    ★ 「全部見えている」と読ませないための行。描けば描くほど、この行が要る。
    """
    t = DOC.read_bytes().decode("utf-8")
    assert "それでもなお出ないもの" in t, (
        "描いた後の限界（実行時の条件分岐・Basic の腕どうしの呼び出し）が書かれていない")
