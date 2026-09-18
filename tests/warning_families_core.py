# -*- coding: utf-8 -*-
"""⚠ の家系の盤 ── 警告を出す場所を**実装から導いて**台帳と突き合わせる（2026-09-18）。

★★ なぜ要るか（Namakoo「三項全ては正しく成り立っているか？ 漏れはない？」）:
  この repo には台帳が並んでいる ── 事後条件・読み直し・断り・配線・記法。
  ところが **⚠ を出す家系だけ台帳が無かった**。個別の試験は在るが
  「家系がいくつ在るか」を持つ物が無く、**増えても減っても誰も気づかない**。
  ★ 棚の処方（片配線の系譜）:「検分は『⚠ を出す全家系の列挙』から入る」。

★★ 置いた瞬間に手勘定が 1 つ直った: 人は `warning_count += 1` を grep して **6** と数えたが、
  実装は **7** だった ── `warning_count += count_suspicious_advisories([msg])` の形を
  落としていた（出力の忠実性の警告）。★ 名簿は手書きせず**宣言から導く**。

★★ この盤が持つ第 2 の列（Namakoo「漏れは塞ぐけど実際に欠陥とは限らないのか」）:
  どの家系も三項（依頼・宣言・実体）のうち 2 項しか持っていない。だが**それが欠陥とは
  限らない** ── 棚の処方は「運べない項があるなら主張の範囲を狭める」で、どれも ✓ を
  名乗らず ⚠ だけを出している（narrowed）。
  ★ そこで盲点ごとに **measured（見えないと実測した）/ assumed（そう思っているだけ）**
    を持たせる。assumed は「まだ調べていない」であって「安全」ではない ──
    配線盤が 2026-09-17 に学んだ「未調査を調べ終えた色と同じにしない」と同じ線。

★ なぜ tests/ に在るか: 素の環境の番人（scripts/_ci_parity_blocker.py）は
  requirements-dev.txt に無い import を全部止めるので、scripts/ 同士の import が
  弾かれる（3 度踏んだ）。中身はここ、scripts/warning_families.py は薄い入口。
"""
from __future__ import annotations

import argparse
import ast
import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

from _product_source import product_files  # noqa: E402

REGISTER = REPO / "tests" / "warning_register.json"

#: 家系の名前として採らない呼び出し（言語の道具・材料を作る側）。
_NOISE = {"getattr", "str", "get", "list", "set", "dict", "int", "len", "sorted",
           "_op_match_pool", "OP_LABELS", "print", "format", "join"}


def _calls_in(node) -> list:
    out = []
    for n in ast.walk(node):
        if isinstance(n, ast.Call):
            f = n.func
            name = f.attr if isinstance(f, ast.Attribute) else getattr(f, "id", "")
            if name and name not in _NOISE:
                out.append(name)
    return out


class _Finder(ast.NodeVisitor):
    """`warning_count` を増やす場所と、その場所が**何を見て**増やしたかを集める。

    ★ 素朴に「囲みの if の test」を読むだけでは足りない ── `_qv = f(...)` のように
      **一度変数へ置いてから** `if _qv:` と書く形が 3 つあり、名前が空になる
      （初版がそうなり、7 件中 3 件が無名で出た）。同じ関数の中で代入まで辿る。
    """

    def __init__(self) -> None:
        self.fn_stack: list = []
        self.block_stack: list = []
        self.hits: list = []

    def visit_FunctionDef(self, node):          # noqa: N802
        self.fn_stack.append(node)
        self.generic_visit(node)
        self.fn_stack.pop()

    def generic_visit(self, node):
        if isinstance(node, (ast.If, ast.For, ast.While)):
            self.block_stack.append(node)
            super().generic_visit(node)
            self.block_stack.pop()
        else:
            super().generic_visit(node)

    def visit_AugAssign(self, node):            # noqa: N802
        t = node.target
        if not (isinstance(t, ast.Name) and t.id == "warning_count"):
            return
        names = _calls_in(node.value)           # ★ 右辺自身も見る（+= f(...) の形）
        owner = self.block_stack[-1] if self.block_stack else None
        if isinstance(owner, ast.If):
            names += _calls_in(owner.test)
            names += self._resolve(owner.test)
        elif isinstance(owner, (ast.For, ast.While)):
            names += _calls_in(getattr(owner, "iter", None) or owner.test)
        # ★★ 並べ替えない ── ast.walk は外側から来るので、**判定をしている関数**が先頭に
        #   立つ。初版は sorted() していて、材料を作る内側（task_quotes_a_value）を
        #   家系の名前に選んでいた（判定は unaccounted_quoted_value の方）。
        self.hits.append({"line": node.lineno,
                           "sources": list(dict.fromkeys(names))})

    def _resolve(self, test) -> list:
        """`if _qv:` の `_qv` を、同じ関数の中の代入まで辿って呼び出し名にする。"""
        if not isinstance(test, ast.Name) or not self.fn_stack:
            return []
        want, out = test.id, []
        for n in ast.walk(self.fn_stack[-1]):
            if isinstance(n, ast.Assign):
                for tgt in n.targets:
                    if isinstance(tgt, ast.Name) and tgt.id == want:
                        out += _calls_in(n.value)
            elif isinstance(n, ast.NamedExpr) and isinstance(n.target, ast.Name) \
                    and n.target.id == want:
                out += _calls_in(n.value)
        return out


def derive() -> dict:
    """実装から家系を導く（名簿を読まない）。source 名 → 行番号。"""
    out = {}
    for p in product_files():
        try:
            tree = ast.parse(p.read_bytes().decode("utf-8"))
        except SyntaxError:
            continue
        f = _Finder()
        f.visit(tree)
        for h in f.hits:
            key = h["sources"][0] if h["sources"] else f"?line{h['line']}"
            out.setdefault(key, []).append(f"{p.name}:{h['line']}")
    return out


def load_register() -> dict:
    return json.loads(REGISTER.read_bytes().decode("utf-8"))


def survey() -> dict:
    reg = load_register()
    got = derive()
    fams = []
    for key, sites in sorted(got.items()):
        rec = reg["families"].get(key)
        fams.append({"key": key, "sites": sites, "declared": rec is not None,
                      "label": (rec or {}).get("label", ""),
                      "terms": (rec or {}).get("terms", {}),
                      "blind_spots": (rec or {}).get("blind_spots", [])})
    return {
        "commit": subprocess.run(["git", "log", "-1", "--format=%h %ad %s", "--date=short"],
                                  cwd=str(REPO), capture_output=True, text=True,
                                  encoding="utf-8").stdout.strip(),
        "families": fams,
        "declared_only": sorted(set(reg["families"]) - set(got)),
    }


def mismatches(data: dict) -> list:
    out = []
    for f in data["families"]:
        if not f["declared"]:
            out.append({"key": f["key"], "kind": "実装に在るのに台帳に無い",
                         "where": ", ".join(f["sites"])})
    for key in data["declared_only"]:
        out.append({"key": key, "kind": "台帳に在るのに実装から消えた", "where": ""})
    return out


TERM_MARK = {True: "○", False: "★ 無い", "partial": "△"}


def render(data: dict) -> str:
    lines = [f"⚠ の家系の盤  {data['commit']}", ""]
    lines.append(f"{'家系':<34} 依頼 宣言 実体   盲点(measured/assumed)")
    lines.append("-" * 92)
    n_assumed = 0
    for f in data["families"]:
        t = f["terms"] or {}
        m = [TERM_MARK.get(t.get(k), "?") for k in ("request", "declaration", "artifact")]
        meas = sum(1 for b in f["blind_spots"] if b.get("status") == "measured")
        assu = sum(1 for b in f["blind_spots"] if b.get("status") == "assumed")
        n_assumed += assu
        flag = "" if f["declared"] else "   ★ 台帳に無い"
        lines.append(f"{(f['label'] or f['key'])[:34]:<34} "
                      f"{m[0]:<4} {m[1]:<4} {m[2]:<5} {meas} / {assu}{flag}")
    lines.append("")
    lines.append(f"家系: {len(data['families'])}   "
                  f"★ まだ調べていない盲点（assumed）: {n_assumed}")
    bad = mismatches(data)
    lines.append(f"台帳と実装の食い違い: {len(bad)}")
    for b in bad:
        lines.append(f"    ★ {b['key']}  {b['kind']}  {b['where']}")
    lines.append("")
    for f in data["families"]:
        if not f["blind_spots"]:
            continue
        lines.append(f"■ {f['label'] or f['key']}")
        for b in f["blind_spots"]:
            mark = "測った" if b.get("status") == "measured" else "★ 未調査"
            lines.append(f"    [{mark}] {b.get('what','')}")
            if b.get("specimen"):
                lines.append(f"              検体: {b['specimen']}")
            if b.get("note"):
                lines.append(f"              {b['note']}")
        lines.append("")
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="⚠ の家系の盤（実装から導く × 台帳）")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    data = survey()
    print(json.dumps(data, ensure_ascii=False, indent=1) if a.json else render(data))
    return 1 if mismatches(data) else 0


if __name__ == "__main__":
    raise SystemExit(main())
