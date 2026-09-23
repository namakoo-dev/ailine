# -*- coding: utf-8 -*-
"""断りは**ガード兼フロンティア候補**として台帳に載せる（2026-09-17）。

★★ なぜ在るか（Namakoo:「後に断りから到達にできるものもあるかもしれないからね」）:
  断りの一覧は UX の点検ではなく**在庫**。棚の線（DEAD_END の二読法）と同じで、
  「なぜ断るか」だけでなく「**何が変われば到達になるか**」を書かないと、
  次に読む人が「できない」と読んで捨てる。

★ 分母は機械で出す ── src/ の print/say に渡る文字列で『？』から始まるもの。
  分類（4 つ）と発火条件は人が書く（機械に決まらない）。盤と同じ作法。

★★ 断りには**家系が 2 つ**ある（2026-09-17 に数えた）:
    ① 画面に『？』を出す経路そのものの断り ── 23 箇所。**ここが台帳の対象**
    ② 検算が返す断り文（op ごと）        ── 149 箇所。画面には「止めた理由: …」の
       1 行として出る。**出口が 1 箇所に畳まれている**ので片配線の危険が構造的に無い。
  だから台帳は①だけで作り、②は「出口が 1 つのまま」を別に縛る（下の試験）。

★ 鍵は func#N（その関数の中で上から N 番目）。途中に断りを差し込むと番号がずれて赤くなる
  ── それは**正しい挙動**（差し込んだ人に分類を書かせる）。
"""
import ast
import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src"
REGISTER = Path(__file__).resolve().parent / "refusal_register.json"

WHY_VALUES = {"vocab", "precondition", "ambiguous", "dangerous"}
OUT_FUNCS = {"print", "say"}
ESCAPE_FUNCS = ("render_retry_options", "render_choices")


def _text_of(node):
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):
        return "".join(v.value if isinstance(v, ast.Constant) and isinstance(v.value, str)
                       else "{…}" for v in node.values)
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
        return _text_of(node.left) + _text_of(node.right)
    return ""


def survey() -> dict:
    """実装から断りを数える ── 鍵は func#N（関数内の出現順）。"""
    out = {}
    from _product_source import src_files
    # ★ 2026-09-23: 手で並べず芯から引く（再帰する ── 初版の `glob("*.py")` は
    #   ailine_core/postconditions/ を見ていなかった）。視野は元と同じ src の下。
    for path in src_files():
        src = path.read_bytes().decode("utf-8")
        tree = ast.parse(src)
        spans = sorted(((n.lineno, getattr(n, "end_lineno", n.lineno), n.name, n)
                        for n in ast.walk(tree)
                        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))),
                       key=lambda s: (s[1] - s[0]))
        found = []
        for node in ast.walk(tree):
            if not (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                    and node.func.id in OUT_FUNCS and node.args):
                continue
            t = _text_of(node.args[0])
            if not t.lstrip().startswith("？"):
                continue
            fn, fnode = "", None
            for a, b, name, nd in spans:
                if a <= node.lineno <= b:
                    fn, fnode = name, nd
                    break
            body = ast.get_source_segment(src, fnode) or "" if fnode else ""
            found.append((fn, node.lineno, re.sub(r"\s+", " ", t),
                          any(e in body for e in ESCAPE_FUNCS),
                          ("例:" in t) or ("例：" in t)))
        seen = {}
        for fn, lineno, t, esc, ex in sorted(found, key=lambda f: f[1]):
            seen[fn] = seen.get(fn, 0) + 1
            out[f"{fn}#{seen[fn]}"] = {"text": t, "escape": esc, "example": ex}
    return out


def _register() -> dict:
    return json.loads(REGISTER.read_bytes().decode("utf-8"))


def test_the_denominator_is_not_empty():
    """★ 分母を先に確かめる ── 空なら下の検査は全部素通りする。"""
    got = survey()
    assert len(got) >= 15, f"断りが少なすぎる: {len(got)}（数え方が壊れている疑い）"


def test_every_refusal_is_in_the_register():
    """★★ 断りを 1 つ足したら、分類と発火条件を書かせる。"""
    got, want = survey(), _register()["refusals"]
    assert set(got) == set(want), (
        f"台帳に無い断り: {sorted(set(got) - set(want))} / "
        f"実体の無い台帳の項目: {sorted(set(want) - set(got))}")


def test_every_entry_is_classified_with_an_unlock():
    """★★ 「なぜ断るか」と「何が変われば到達になるか」の両方。★ 未調査 と書くのは可。"""
    want = _register()["refusals"]
    assert want, "★ 台帳が空"
    for key, e in want.items():
        assert e.get("why") in WHY_VALUES, f"{key}: why が {WHY_VALUES} のどれでもない: {e.get('why')}"
        assert str(e.get("unlock", "")).strip(), (
            f"{key}: unlock（何が変われば到達になるか）が空 ── "
            "分からないなら『★ 未調査』と書くこと。空欄は『無い』と読まれる")


def test_the_measured_fields_match_what_the_machine_sees():
    """★★ escape / example は**機械が見た値**。人が書き換えていないこと。

    ★ ここが縛られていないと、逃げ道を落とした変更が台帳には反映されず、
      台帳だけが「逃げ道あり」と言い続ける（宣言と実体のずれ＝今日ずっと潰した形）。
    """
    got, want = survey(), _register()["refusals"]
    for key in sorted(set(got) & set(want)):
        for field in ("escape", "example"):
            assert bool(want[key].get(field)) == bool(got[key][field]), (
                f"{key}: 台帳の {field}={want[key].get(field)} だが実体は {got[key][field]}")


def test_the_dangerous_ones_do_not_claim_an_unlock():
    """★ 危険だから断るものに『到達できる』と書かない（在庫と混ぜない）。"""
    for key, e in _register()["refusals"].items():
        if e.get("why") == "dangerous":
            assert "無い" in str(e.get("unlock", "")), (
                f"{key}: 危険側の断りに発火条件が書かれている ── 在庫に混ぜない")


def test_the_second_family_still_has_one_exit():
    """★★ もう 1 つの家系（検算が返す断り文・149 箇所）は**出口が 1 つ**のまま。

    ★ こちらを台帳にしないのは、出口が畳まれていて片配線の危険が構造的に無いから。
      その前提が崩れたら（出口が 2 つ以上になったら）、台帳の設計を見直す合図。
    """
    from _product_source import count_in_product
    assert count_in_product('f"  止めた理由: {reason}"') == 1, (
        "検算の断りの出口が 1 箇所でなくなった ── 台帳の分母の考え方を見直すこと")
