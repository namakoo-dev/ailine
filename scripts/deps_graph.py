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
                # ★★ 2026-09-16: `from ailine_core import subject` は部品名が **names 側**に
                #   来るので、辺がパッケージ止まりになる。実測すると **105 本**がこの形で
                #   隠れており、部品名で書かれた辺 85 本より多かった ──
                #   図は配線の半分も描いていなかった（`ailine_core` の被依存 23 はその影）。
                #   ★ 名前が実在の部品なら、そこへ辺を張る。
                for a in n.names:
                    sub = f"{n.module}.{a.name}"
                    if sub in mods and sub != name:
                        imports[name].add(sub)
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
        "★★ 2026-09-16: ここに挙げた 2 つ（辞書経由・Basic 側）に**バグが 3 件住んでいた**。",
        "  『見えない』と断って済ませるのをやめ、**下に層②・層③として描いた**。",
        "  それでもなお出ないもの: 実行時の条件分岐（どの経路を通るかは依頼文で変わる）と、",
        "  Basic の腕どうしの呼び出し関係。",
        "",
    ]
    return "\n".join(header + lines + _wiring_section()) + "\n"


def _op_wiring() -> dict:
    """op → (生成関数, 事後条件, 呼ぶ Basic の腕) を**実体から**引く。

    ★ なぜ要るか: この図はもともと import だけを見ていて、自分で
      「辞書経由の呼び出しは辺に出ない／Basic 側は見えない」と断っていた。
      2026-09-16 に出たバグ 3 件は**全部そこ**に住んでいた ──
        ・名簿 `_OPS_THAT_SKIP_NON_DATA_ROWS` が AGGREGATE を落として合計行が消えた
        ・`PROJECTIONS` が DEDUP を落として「失うもの」を告げなかった
        ・`AXES` が AGGREGATE/PIVOT を落として平均が黙って合計になった
      断ったなら描く。描けない理由が無いものを「見えない」で済ませない。
    """
    import inspect
    import re as _re
    sys.path.insert(0, str(SRC))
    import ailine  # noqa: E402

    bas = (SRC / "ailine" / "helpers" / "AiLineHelpers.bas").read_text(
        encoding="utf-8", errors="replace")
    arms = sorted(set(_re.findall(r"^\s*(?:Sub|Function)\s+(\w+)", bas, _re.M)))

    def _src(fn):
        try:
            return inspect.getsource(fn)
        except (OSError, TypeError):
            return ""

    rows = []
    for op in sorted(ailine.OP_SCHEMA):
        fn = ailine.CODEGEN_BY_OP.get(op)
        pc = ailine.POSTCONDITIONS.get(op)
        body = _src(fn)
        called = sorted({m for m in _re.findall(r"Call\s+(\w+)\s*\(", body) if m in arms})
        rows.append({
            "op": op,
            "codegen": getattr(fn, "__name__", "(無し)"),
            "post": getattr(pc, "__name__", "★ 辞書に載らない"),
            "arms": called,
        })

    # ★ 腕に届いているかは **Python 全体**から数える ── 生成関数だけを分母にすると
    #   `wrap()` が足す MoveColumnTo や、別経路の WriteInspectionSheet を
    #   「死んでいる」と誤って名指しする（試作で実際に踏んだ）。
    whole = "\n".join(p.read_text(encoding="utf-8", errors="replace")
                       for p in [SRC / "ailine" / "__init__.py"]
                       + sorted((SRC / "ailine_core").rglob("*.py")))
    unreached = [a for a in arms
                 if not _re.search(r"\b" + a + r"\b", whole)
                 and not _re.search(r"\b" + a + r"\s*\(",
                                    _re.sub(r"^\s*(?:Sub|Function)\s+" + a + r"\b.*$",
                                            "", bas, flags=_re.M))]
    return {"rows": rows, "arms": arms, "unreached": sorted(unreached)}


def _op_rosters() -> dict:
    """op 名を並べた名簿（dict/set/frozenset/tuple/list）を実体から集める。

    ★ 走査の本体は tests/test_op_completeness.py の `discover_op_rosters`。
      **数え方を 2 つ持たない** ── 番人と図が別々に数えると、片方だけ直る。
    """
    sys.path.insert(0, str(REPO / "tests"))
    from test_op_completeness import discover_op_rosters  # noqa: E402
    sys.path.insert(0, str(SRC))
    import ailine  # noqa: E402
    n = len(ailine.OP_SCHEMA)
    found = discover_op_rosters()
    partial = {k: v["ops"] for k, v in found.items() if len(v["ops"]) < n}
    return {"all": found, "partial": partial, "n": n,
            "ops": sorted(ailine.OP_SCHEMA)}


def _wiring_section() -> list:
    w = _op_wiring()
    r = _op_rosters()
    short = {k: k.split(":")[-1] for k in r["partial"]}
    out = [
        "",
        "---",
        "",
        "## 層② op の配線（import では見えない）",
        "",
        "★ 上の図は **import しか見ていない**。op の実処理は「辞書で引いて呼ぶ」ので、",
        "  `op → 生成関数 → Basic の腕 → 事後条件` の道筋は 1 本も辺に出ない。",
        "  2026-09-16 のバグ 3 件は**全部この層**に住んでいた。だからここに並べる。",
        "",
        f"- Basic の腕 **{len(w['arms'])}** 本 / Python のどこからも名指しされない腕: "
        f"**{len(w['unreached'])}** 本"
        + ("（" + " / ".join("`" + a + "`" for a in w["unreached"]) + "）"
           if w["unreached"] else ""),
        "",
        "| op | 生成関数 | Basic の腕 | 事後条件 |",
        "|---|---|---|---|",
    ]
    for row in w["rows"]:
        arms = " / ".join("`" + a + "`" for a in row["arms"]) or "（Call を使わず直に書く）"
        out.append(f"| `{row['op']}` | `{row['codegen']}` | {arms} | `{row['post']}` |")

    # 同じ腕を分け合う op ── 片方だけ直すと片配線になる組
    share = {}
    for row in w["rows"]:
        for a in row["arms"]:
            share.setdefault(a, []).append(row["op"])
    shared = {a: v for a, v in share.items() if len(v) > 1}
    out += ["",
            "★ **同じ腕を分け合う op**（片方だけ直すと、もう片方に同じ穴が残る）:",
            ""]
    out += [f"- `{a}` ← " + " / ".join("`" + o + "`" for o in sorted(v))
            for a, v in sorted(shared.items())] or ["- （無し）"]

    out += [
        "",
        "---",
        "",
        "## 層③ op の名簿（どの op がどの扱いを受けるか）",
        "",
        "★ ailine の判断の多くは「この op はこの名簿に居るか」で決まる。名簿は",
        "  コードの中の dict / set / tuple なので、これも import の辺には出ない。",
        "★ **全 op が居る名簿は載せない**（どの op も同じなので読む意味が無い）。",
        f"  ここに出るのは**部分**の名簿 **{len(r['partial'])}** 本だけ。",
        "★ 名簿が部分であること自体は多くの場合正しい。理由と解除条件は",
        "  `tests/op_completeness_register.json` の `op_roster_coverage` に書いてある。",
        "",
        "| op | " + " | ".join(short[k] for k in sorted(r["partial"])) + " |",
        "|---|" + "---|" * len(r["partial"]),
    ]
    for op in r["ops"]:
        cells = ["○" if op in r["partial"][k] else "·" for k in sorted(r["partial"])]
        out.append(f"| `{op}` | " + " | ".join(cells) + " |")
    out.append("")
    return out


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
