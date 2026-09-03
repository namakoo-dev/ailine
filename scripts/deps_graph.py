# -*- coding: utf-8 -*-
"""モジュール依存を Mermaid の図にする（実体から生成する・手で描かない）。

★ なぜ在るか（2026-09-03・Namakoo「全体の依存関係を読みだせるようにしたい。
  修正や変更の際に依存関係の破れや影響範囲を特定できるようにしておきたい」）:
  分割で ailine_core が 46 モジュールになった。どれがどれに依存しているかを
  読める形にしないと、次に何かを動かすとき「何に触るか」が人の記憶頼りになる。

★ 手で描かない: この repo は「人が書いた数は必ず古くなる」を何度も踏んでいる
  （README の行数・試験数・翻訳精度）。図も同じなので**実体から生成**し、
  tests/test_dependency_graph_is_current.py が「図と実体が一致すること」を守る。

★★ この図の限界（先に書く。読む人が「全部見えている」と誤解しないため）:
  ・**import だけを見る**。辞書や getattr 経由の呼び出しは辺として現れない。
    ailine は POSTCONDITIONS 辞書で op → 事後条件を引くので、その 28 本は
    「本体から呼ばれていない」ように見える（実際は毎回呼ばれている）
  ・Basic 側（helpers/*.bas）への依存も見えない
  ・★ **「辺が無い＝影響が無い」ではない。「import では繋がっていない」まで。**

    python scripts/deps_graph.py            # Mermaid を標準出力へ
    python scripts/deps_graph.py --write    # docs/依存関係.md を書き換える
"""
from __future__ import annotations

import argparse
import ast
import sys
from collections import Counter, defaultdict
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src"
DOC = REPO / "docs" / "依存関係.md"


def modules() -> dict:
    out = {"ailine": SRC / "ailine" / "__init__.py"}
    for p in sorted((SRC / "ailine_core").rglob("*.py")):
        rel = p.relative_to(SRC).with_suffix("")
        name = ".".join(rel.parts)
        if name.endswith(".__init__"):
            name = name[: -len(".__init__")]
        out[name] = p
    return out


def edges() -> tuple:
    mods = modules()
    imports = defaultdict(set)
    for name, path in mods.items():
        tree = ast.parse(path.read_bytes().decode("utf-8"))
        for n in ast.walk(tree):
            if isinstance(n, ast.ImportFrom) and n.module and n.module.startswith("ailine"):
                if n.module in mods and n.module != name:
                    imports[name].add(n.module)
            elif isinstance(n, ast.Import):
                for a in n.names:
                    if a.name in mods and a.name != name:
                        imports[name].add(a.name)
    return mods, imports


def _id(name: str) -> str:
    return name.replace(".", "_")


def _short(name: str) -> str:
    return name.replace("ailine_core.", "")


def _ranked(counter):
    """被依存の多い順。★ **同数のときは名前順**で並べる。

    ★ Counter.most_common は同数の並びを保証しない ── 実行ごとに順序が変わると、
      生成した図が毎回ちがう文字列になり、**一致を守る番人が永久に赤くなる**。
      2026-09-03 に実際そうなった（4 の 3 モジュールが入れ替わった）。
    """
    return sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))


def render() -> str:
    mods, imports = edges()
    n_edge = sum(len(v) for v in imports.values())
    indeg = Counter(t for ts in imports.values() for t in ts)
    lines = ["```mermaid", "graph LR"]
    groups = defaultdict(list)
    for name in sorted(mods):
        if name == "ailine":
            groups["本体"].append(name)
        elif name.startswith("ailine_core.postconditions"):
            groups["事後条件（op ごと）"].append(name)
        else:
            groups["ailine_core（部品）"].append(name)
    for gi, (g, names) in enumerate(groups.items()):
        lines.append(f"  subgraph G{gi}[{g}]")
        for name in names:
            label = _short(name)
            deg = indeg.get(name, 0)
            if deg >= 5:
                label = f"{label}<br/>被依存 {deg}"
            lines.append(f"    {_id(name)}[{label}]")
        lines.append("  end")
    for src_ in sorted(imports):
        for dst in sorted(imports[src_]):
            lines.append(f"  {_id(src_)} --> {_id(dst)}")
    lines.append("```")
    header = [
        "# 依存関係（自動生成）",
        "",
        "★ **この図は `scripts/deps_graph.py` が実体から生成する。手で書き換えない。**",
        "  `tests/test_dependency_graph_is_current.py` が図と実体の一致を守っている。",
        "",
        f"- モジュール **{len(mods)}** / import の辺 **{n_edge}**",
        f"- 層の向きの違反（`ailine_core` → `ailine`）: "
        f"**{sum(1 for m, ts in imports.items() if m != 'ailine' and 'ailine' in ts)} 件**"
        "（core は本体を知らない ＝ 本体だけを差し替えられる）",
        "- 被依存が多い順: "
        + " / ".join(f"`{_short(m)}`({n})" for m, n in _ranked(indeg)[:5]),
        "",
        "★★ **この図に出ないもの**（読む人が「全部見えている」と誤解しないため）:",
        "",
        "- **辞書や getattr 経由の呼び出し**。`POSTCONDITIONS` 辞書で op → 事後条件を",
        "  引くので、事後条件 28 本は「本体から呼ばれていない」ように見える（実際は毎回呼ばれる）",
        "- Basic 側（`helpers/*.bas`）への依存",
        "- ★ **「辺が無い＝影響が無い」ではない。「import では繋がっていない」まで。**",
        "",
    ]
    return "\n".join(header + lines) + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--write", action="store_true", help="docs/依存関係.md を書き換える")
    a = ap.parse_args(argv)
    text = render()
    if a.write:
        DOC.parent.mkdir(exist_ok=True)
        DOC.write_bytes(text.encode("utf-8"))
        print(f"書いた: {DOC}（{len(text.splitlines())} 行）")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
