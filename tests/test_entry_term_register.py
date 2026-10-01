# -*- coding: utf-8 -*-
"""入口の項の台帳（tests/entry_term_register.json）を宣言と等号で縛る（2026-10-01）。

★★ なぜ在るか: 依頼の項の台帳（tests/request_term_register.json）は `run` の 31 op だけを
  数えていた。盲検 7 体目の請求と入金の突き合わせ（2 冊の照合）で、依頼が名指しした
  『税込金額』が式のままで値が無く、残った『税抜金額』が**断りなく**選ばれた ──
  照合の道は op ではないので、地図に載っていなかった。★ 載っていない所に同じ形が在った。
  run 以外の入口（照合・フォルダ抽出・split・verify・accounts…）を同じ作法で数える。

★ 別の台帳にした理由: 分母の導き方が違う（op の宣言でなく argparse の登録簿と入口の関数）。
  1 冊に混ぜると等号の片側が 2 種類の導出の和になり、どちらかが壊れても和で隠れる。

★ 分母（入口ごとの項目）は宣言から導く:
  ① 入口 ＝ argparse に登録されたサブコマンド（登録簿と等号）
  ② 項目 ＝ その入口のオプション（`--dest`）
           ∪ 入口の関数から 3 段までに届く「決める器官」が決める項目（_meta.deciders）
             ・`worksheets[0]` / `sheetnames[0]` → シート
             ・`load_workbook(data_only=True)` → 式の値
             ・resolve_columns → ColumnResolution の列の項（dataclass の宣言から）
             ・read_journal → 仕訳の列の役割（accounts_core.COLUMN_ALIASES の宣言から）
             ・translate_task → 入口の関数が読む LLM の引数の鍵（args の添字・get）
  ★ `run` の 1 冊の道は request_term_register.json が持つ ── ここで数える関数は
    _meta.route_functions で名指しした 2 冊の照合とフォルダ抽出（関数の実在は番人が見る）。
  ★ 限界（正直に）: 「決める器官」の名簿そのものは手で持つ。新しい種類の器官が生えても
    ここは鳴らない ── 名簿に足すまで見えない。鳴るのは**既知の器官が新しい入口に届いた時**と
    **新しいオプションが生えた時**。
"""
import argparse
import ast
import dataclasses
import inspect
import json
import sys
import textwrap
import types
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))
import ailine  # noqa: E402
from ailine_core import accounts_core, match  # noqa: E402
from test_request_term_register import _defined, _keys_touched  # noqa: E402

REGISTER = REPO / "tests" / "entry_term_register.json"
CLASSES = ("A", "B", "C", "D")
DEPTH = 3


def _load(path=REGISTER) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


# --- 分母を宣言から導く ---------------------------------------------------------------

def _subparsers() -> dict:
    out = {}
    for ac in ailine.build_parser()._actions:
        if isinstance(ac, argparse._SubParsersAction):
            out.update(ac.choices)
    return out


def _resolution_items() -> set:
    """ColumnResolution の宣言から、照合が決める列の項（判定や在庫の欄は除く）。"""
    return {f.name for f in dataclasses.fields(match.ColumnResolution)
            if f.name not in ("ok", "unresolved", "unusable")}


def _journal_items() -> set:
    return {f"列:{role}" for role in accounts_core.COLUMN_ALIASES}


def _callee(mod, node):
    """呼び出しの先が ailine / ailine_core の関数なら、その関数。"""
    f = node.func
    if isinstance(f, ast.Name):
        obj = getattr(mod, f.id, None)
    elif isinstance(f, ast.Attribute) and isinstance(f.value, ast.Name):
        base = getattr(mod, f.value.id, None)
        obj = getattr(base, f.attr, None) if isinstance(base, types.ModuleType) and \
            base.__name__.startswith("ailine_core") else None
    else:
        obj = None
    if isinstance(obj, types.FunctionType) and (obj.__module__ or "").startswith("ailine"):
        return obj
    return None


def _call_name(node) -> str:
    f = node.func
    return f.id if isinstance(f, ast.Name) else (f.attr if isinstance(f, ast.Attribute) else "")


def reached_items(root, deciders: dict, depth: int = DEPTH) -> set:
    """入口の関数 root から depth 段までに届く「決める器官」が決める項目。"""
    out: set = set()
    seen: set = set()

    def walk(fn, d):
        key = (fn.__module__, fn.__qualname__)
        if key in seen:
            return
        seen.add(key)
        mod = sys.modules[fn.__module__]
        try:
            tree = ast.parse(textwrap.dedent(inspect.getsource(fn)))
        except (OSError, TypeError):
            return
        for node in ast.walk(tree):
            if (isinstance(node, ast.Subscript) and isinstance(node.value, ast.Attribute)
                    and node.value.attr in ("worksheets", "sheetnames")
                    and isinstance(node.slice, ast.Constant) and node.slice.value == 0):
                out.add("シート")
            if not isinstance(node, ast.Call):
                continue
            name = _call_name(node)
            if name == "load_workbook" and any(
                    k.arg == "data_only" and isinstance(k.value, ast.Constant) and k.value.value
                    for k in node.keywords):
                out.add("式の値")
            if name == "resolve_columns":
                out.update(_resolution_items())
            if name == "read_journal":
                out.update(_journal_items())
            if name == "translate_task":
                out.update(f"LLM:{k}" for k in _keys_touched(root, "args", "EXTRACT"))
            out.update(deciders.get(name, ()))
            callee = _callee(mod, node)
            if callee is not None and d > 0 and not callee.__name__.startswith("cmd_"):
                walk(callee, d - 1)

    walk(root, depth)
    return out


def declared_items(reg=None) -> set:
    reg = reg or _load()
    deciders = reg["_meta"]["deciders"]
    override = reg["_meta"]["route_functions"]
    items = set()
    for route, sp in _subparsers().items():
        opts = {f"--{a.dest}" for a in sp._actions if a.option_strings and a.dest != "help"}
        # ★ 1 つの入口に道が複数ある時（run: 2冊の照合／フォルダ抽出）は、道ごとに項目を分ける
        #   （同じ「シート」でも道ごとに区分が違いうる ── 混ぜると片方の在庫が隠れる）。
        if route in override:
            roots = [(f"{label}:", getattr(ailine, name))
                     for label, name in override[route]["functions"].items()]
        else:
            roots = [("", sp.get_default("func"))]
        implicit = set()
        for prefix, root in roots:
            implicit |= {prefix + i for i in reached_items(root, deciders)}
        items |= {(route, i) for i in opts | implicit}
    return items


def _register_items(reg) -> set:
    return {(e["route"], e["item"]) for e in reg["items"]}


def _counts(reg) -> dict:
    out = {c: 0 for c in CLASSES}
    for e in reg["items"]:
        out[e["class"]] += 1
    return out


def _cap_breaches(reg) -> list:
    n, caps = _counts(reg), reg["_meta"]["caps"]
    bad = []
    for c in ("B", "D"):
        if n[c] > caps[c]:
            bad.append(f"{c} が {n[c]} 件（上限 {caps[c]}）── 依頼と突き合わせない項目が増えた")
        elif n[c] < caps[c]:
            bad.append(f"{c} が {n[c]} 件に減った ── 上限（{caps[c]}）を {n[c]} に下げてください")
    return bad


# --- ① 入口と項目は宣言から・台帳と等号 ---------------------------------------------------

def test_every_route_is_in_the_register():
    """入口の集合（登録簿）と、台帳が名を挙げた入口の集合が等しい（項目の無い入口も載せる）。"""
    reg = _load()
    registered = set(_subparsers())
    listed = set(reg["routes"])
    assert registered == listed, (
        f"台帳に無い入口: {sorted(registered - listed)} / 登録簿に無い入口: {sorted(listed - registered)}")


def test_every_declared_item_is_classified_and_nothing_else():
    reg = _load()
    declared, written = declared_items(reg), _register_items(reg)
    unclassified, stale = sorted(declared - written), sorted(written - declared)
    assert not unclassified and not stale, (
        "台帳と宣言がずれている:\n"
        + "".join(f"  台帳に無い（区分を書いてください）: {r} {k}\n" for r, k in unclassified)
        + "".join(f"  宣言に無い（台帳から消してください）: {r} {k}\n" for r, k in stale))


def test_routes_without_items_say_why():
    """項目が 0 の入口は、なぜ 0 かを書く（空欄は『見ていない』と区別できない）。"""
    reg = _load()
    have = {r for r, _k in _register_items(reg)}
    silent = [r for r, v in reg["routes"].items() if r not in have and not str(v).strip()]
    assert not silent, f"項目が無いのに理由が空: {silent}"


def test_the_denominator_is_not_empty():
    """★ 陽性対照: 導出が壊れて空になれば、等号は台帳も空で恒真になりうる。"""
    d = declared_items()
    assert len(d) >= 80, len(d)
    for must in (("run", "2冊の照合:amount_a"), ("run", "2冊の照合:シート"),
                 ("run", "フォルダ抽出:LLM:value"),
                 ("split", "式の値"), ("split", "--amount"), ("accounts", "列:借方勘定科目"),
                 ("verify", "式の値")):
        assert must in d, must


def test_a_new_option_turns_this_red(monkeypatch):
    """変異: 入口に新しいオプションが生えたら未分類として出る。"""
    orig = ailine.build_parser

    def mutant():
        ap = orig()
        for ac in ap._actions:
            if isinstance(ac, argparse._SubParsersAction):
                ac.choices["split"].add_argument("--架空の列")
        return ap

    monkeypatch.setattr(ailine, "build_parser", mutant)
    assert ("split", "--架空の列") in declared_items() - _register_items(_load())


def test_a_known_organ_reaching_a_new_route_turns_this_red(monkeypatch):
    """変異: 既知の「決める器官」が新しく入口に届いたら（ここでは照合の列の解決）未分類として出る。"""
    def cmd_export_pdf(a):
        return match.resolve_columns("", [], [], [], [])

    monkeypatch.setattr(ailine, "cmd_export_pdf", cmd_export_pdf)
    orig = ailine.build_parser

    def mutant():
        ap = orig()
        for ac in ap._actions:
            if isinstance(ac, argparse._SubParsersAction):
                ac.choices["export-pdf"].set_defaults(func=cmd_export_pdf)
        return ap

    monkeypatch.setattr(ailine, "build_parser", mutant)
    assert ("export-pdf", "amount_a") in declared_items() - _register_items(_load())


# --- ② B と D は縮める在庫 -----------------------------------------------------------------

def test_the_unmatched_stock_only_shrinks():
    reg = _load()
    bad = [f"{e['route']} {e['item']}: 区分『{e['class']}』は A/B/C/D のどれでもない"
           for e in reg["items"] if e["class"] not in CLASSES]
    assert not bad, "\n".join(bad)
    assert not _cap_breaches(reg), "\n".join(_cap_breaches(reg))


def test_the_stock_guard_rings_both_ways():
    reg = _load()
    grow = json.loads(json.dumps(reg))
    next(e for e in grow["items"] if e["class"] == "A")["class"] = "D"
    assert any("増えた" in b for b in _cap_breaches(grow))
    shrink = json.loads(json.dumps(reg))
    stock = next((c for c in ("D", "B") if any(e["class"] == c for e in shrink["items"])), None)
    if stock is not None:
        next(e for e in shrink["items"] if e["class"] == stock)["class"] = "A"
    else:
        shrink["_meta"]["caps"]["B"] += 1
    assert any("下げて" in b for b in _cap_breaches(shrink))


def test_the_history_ends_at_the_current_counts():
    reg = _load()
    assert reg["_meta"]["history"][-1]["counts"] == {c: n for c, n in _counts(reg).items() if n}


def test_the_stock_says_what_is_left():
    """B と D の項目は、なぜ残したか（stock）を書く ── 在庫は理由つきで持つ。"""
    reg = _load()
    silent = [f"{e['route']} {e['item']}" for e in reg["items"]
              if e["class"] in ("B", "D") and not str(e.get("stock", "")).strip()]
    assert not silent, "在庫の理由が空: " + "、".join(silent)


# --- ③ 根拠が実在する ----------------------------------------------------------------------

def test_every_basis_points_at_a_function_that_exists():
    reg = _load()
    rotten, empty = [], []
    for e in reg["items"]:
        if not e.get("basis"):
            empty.append(f"{e['route']} {e['item']}")
        for b in e.get("basis") or ():
            path, _, func = b.rpartition(":")
            if not _defined(path, func):
                rotten.append(f"{e['route']} {e['item']}: {b}")
    for r, v in reg["_meta"]["route_functions"].items():
        for fn in v["functions"].values():
            if not callable(getattr(ailine, fn, None)):
                rotten.append(f"route_functions[{r}]: {fn}")
    assert not empty, "根拠が空: " + "、".join(empty)
    assert not rotten, "根拠の関数が見つからない:\n  " + "\n  ".join(rotten)
