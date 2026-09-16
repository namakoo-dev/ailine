# -*- coding: utf-8 -*-
"""図を**逆から**読む ── 在るのに誰も呼んでいないものを数える（2026-09-16）。

★ なぜ在るか（Namakoo「構築したグラフから逆に未配線などは見つけられたりはしない？」）:
  `scripts/deps_graph.py` は**在る配線**を描く。同じ材料で**在るべきなのに無い配線**も
  数えられる。実際、初回の走査で「候補 0 件のときの断り文
  （`attributes.render_no_evidence`）が三兄弟のうち 1 つだけ配線されていない」を掴んだ。

★★ 大前提（ここを外すと害になる）: **到達不能＝宝ではない**。
  別プロジェクトで 833 行の未配線モジュールを「宝」と呼び、実際は断念済みだった事故がある。
  ここが出すのは**候補**であって判定ではない。分類は人がやる。
  ★ だから「消す」機能は付けない。数えて名指しするところで止める。

★ 仕分けの結果は `tests/unwired_register.json` に宣言する。
  `tests/test_unwired_is_declared.py` が「新しい未到達が理由なしに増えたら赤」にする
  （2026-09-16 の被覆台帳と同じ作り ── 決めた時に書く方式にすると、次に増えた分が漏れる）。

    python scripts/unwired.py            # 一覧を出す
    python scripts/unwired.py --json     # 機械が読む形（番人が使う）

★ この道具の限界（先に書く）:
  ・**名前の出現**で数える。文字列から動的に引く呼び出し（getattr・辞書の値）は追えない
  ・だから「0 件」は「呼ばれていない」ではなく「**名前がどこにも書かれていない**」まで
  ・初版は `from ailine_core import intent` の形を数え落として、孤児を 34 本と誤報した
    （実測で気づいた ── 数が多すぎたから。少なく出ていたら見逃していた）
"""
from __future__ import annotations

import argparse
import ast
import json
import re
import sys
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src"


def _modname(p: Path) -> str:
    rel = p.relative_to(SRC).with_suffix("")
    n = ".".join(rel.parts)
    return n[: -len(".__init__")] if n.endswith(".__init__") else n


def _sources() -> tuple:
    prod = [SRC / "ailine" / "__init__.py"] + sorted((SRC / "ailine_core").rglob("*.py"))
    around = (sorted((REPO / "tests").rglob("*.py")) + sorted((REPO / "bench").rglob("*.py"))
              + sorted((REPO / "scripts").rglob("*.py")))
    return prod, around


_WORD = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


def _word_counts(text: str) -> Counter:
    """その文章に現れる語を**一度だけ**数える。

    ★ 速さのために在る: 名前ごとに全文へ正規表現を掛けると 1 回 109 秒かかり、
      番人にできなかった（全件が 5 分半なので 1 本で 3 分の 1 を食う）。
      語を 1 回数えて辞書で引く形にすると同じ答えが秒で出る。
    ★ 数えるのは**識別子だけ**（日本語のコメントや文字列の中の語は拾わない）。
    """
    return Counter(_WORD.findall(text))


def survey() -> dict:
    prod_files, around_files = _sources()
    prod = {p: p.read_text(encoding="utf-8", errors="replace") for p in prod_files}
    around = "\n".join(p.read_text(encoding="utf-8", errors="replace") for p in around_files)
    all_prod = "\n".join(prod.values())

    # --- ① 製品の中で誰も import しないモジュール ---
    imported: set = set()
    for text in prod.values():
        for n in ast.walk(ast.parse(text)):
            if isinstance(n, ast.ImportFrom) and (n.module or "").startswith("ailine"):
                imported.add(n.module)
                # ★ `from ailine_core import intent` は**部品名が names 側**に来る。
                #   初版はここを数え落として孤児 34 本と誤報した。
                for a in n.names:
                    imported.add(n.module + "." + a.name)
            elif isinstance(n, ast.Import):
                for a in n.names:
                    if a.name.startswith("ailine"):
                        imported.add(a.name)
    orphan_modules = sorted(_modname(p) for p in prod_files
                            if _modname(p) not in imported and _modname(p) != "ailine")

    # --- ②⑥ 関数の到達 ---
    defs: dict = {}
    for p, text in prod.items():
        for n in ast.walk(ast.parse(text)):
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)) and not n.name.startswith("__"):
                defs.setdefault(n.name, (_modname(p), n.lineno))
    # ★ 定義そのものの 1 回を引く ── 同名が複数モジュールに在る場合はその数だけ引く。
    def_count: Counter = Counter()
    for text in prod.values():
        for n in ast.walk(ast.parse(text)):
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
                def_count[n.name] += 1
    prod_words = _word_counts(all_prod)
    around_words = _word_counts(around)
    nowhere, test_only = [], []
    for name, where in defs.items():
        in_prod = prod_words[name] - def_count[name]
        in_around = around_words[name]
        if in_prod <= 0 and in_around == 0:
            nowhere.append(f"{where[0]}:{name}")
        elif in_prod <= 0:
            test_only.append(f"{where[0]}:{name}")

    # --- ③ 誰も読まない op の名簿 ---
    sys.path.insert(0, str(REPO / "tests"))
    from test_op_completeness import discover_op_rosters  # noqa: E402
    unread_rosters = []
    for key in sorted(discover_op_rosters()):
        var = key.split(":")[-1]
        body = re.sub(r"^\s*" + re.escape(var) + r"\s*=.*$", "", all_prod, flags=re.M)
        if not re.search(r"\b" + re.escape(var) + r"\b", body):
            unread_rosters.append(key)

    # --- ⑤ 製品からも他の腕からも届かない Basic の腕 ---
    bas = (SRC / "ailine" / "helpers" / "AiLineHelpers.bas").read_text(
        encoding="utf-8", errors="replace")
    arms = sorted(set(re.findall(r"^\s*(?:Sub|Function)\s+(\w+)", bas, re.M)))
    unreached_arms = []
    for a in arms:
        inner = re.sub(r"^\s*(?:Sub|Function)\s+" + a + r"\b.*$", "", bas, flags=re.M)
        if not re.search(r"\b" + a + r"\b", all_prod) and not re.search(
                r"\b" + a + r"\s*\(", inner):
            unreached_arms.append(a)

    return {
        "orphan_modules": orphan_modules,
        "functions_named_nowhere": sorted(nowhere),
        "functions_only_tests_name": sorted(test_only),
        "rosters_no_one_reads": unread_rosters,
        "basic_arms_unreached": sorted(unreached_arms),
    }


KINDS = ("orphan_modules", "functions_named_nowhere", "functions_only_tests_name",
         "rosters_no_one_reads", "basic_arms_unreached")
LABEL = {
    "orphan_modules": "製品の中で誰も import しないモジュール",
    "functions_named_nowhere": "製品からも試験からも名指しされない関数",
    "functions_only_tests_name": "試験からしか名指しされない関数",
    "rosters_no_one_reads": "宣言されているのに製品が読まない op の名簿",
    "basic_arms_unreached": "製品からも他の腕からも届かない Basic の腕",
}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--json", action="store_true", help="機械が読む形で出す")
    a = ap.parse_args(argv)
    got = survey()
    if a.json:
        print(json.dumps(got, ensure_ascii=False, indent=2))
        return 0
    print("★ ここに出るのは**候補**。到達不能＝宝ではない ── 分類は人がする。")
    print("★ 数えているのは『名前の出現』であって呼び出しではない（動的な呼び出しは追えない）。\n")
    for k in KINDS:
        v = got[k]
        print(f"■ {LABEL[k]}: {len(v)}")
        for x in v:
            print("   ", x)
        print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
