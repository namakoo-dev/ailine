# -*- coding: utf-8 -*-
"""依頼の項の台帳（tests/request_term_register.json）を宣言と等号で縛る（2026-10-01）。

★★ なぜ在るか: 「嘘の ✓」は、事後条件が**宣言↔実体**しか見ない所で生まれる。
  2026-09-30 の盲検 7 体目は、並べ替えの向きを LLM だけが決めていた
  （「新しい順」を昇順と読み、何も動かず ✓）。直したのは 1 項目だったが、
  同じ形の項目が**どこに・何個**残っているかを誰も持っていなかった。
  ★ 地図（読むだけの子が作った 66 項目）を台帳にし、ここで 3 つを縛る:
    ① (op, 項目) の集合は**宣言から導く**（手書きの名簿を分母にしない）── 台帳と等号。
       新しい op や項目が増えたら、台帳に区分を書くまで赤。
    ② B（実表とだけ）と D（何とも突き合わせない）は**縮める在庫**。上限を超えたら赤、
       下回ったら上限を下げるまで赤（下がる向きにしか動かない）。
    ③ 根拠（ファイル:関数）が実在すること（腐った根拠を残さない）。

★ 分母の導き方（地図と同じ 66 になることを確かめてから入れた）:
    OP_SCHEMA ∪ OP_SUBJECT_SLOTS ∪ _CONFIRM_FIELDS（_ で始まらない鍵）
    ∪ CODEGEN_BY_OP・POSTCONDITIONS・verify_dsl_args の振り分け先が読み書きする resolved の鍵
  ★ 共有される検査関数（_verify_bold・_verify_add_row）の中の**他の op の枝**
    （`if op == "FILL_COLOR":` 等）は数えない ── 読まないと BOLD.color・ADD_ROW.count という
    存在しない項目が生える（実測）。
"""
import ast
import functools
import inspect
import json
import sys
import textwrap
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import argcheck, codegen  # noqa: E402

REGISTER = REPO / "tests" / "request_term_register.json"
CLASSES = ("A", "B", "C", "D")


# --- 分母を宣言から導く ---------------------------------------------------------------

def _op_test(test, op):
    """if の条件が op の道か。True=この op の道 / False=他の op の道 / None=op と無関係。"""
    if isinstance(test, ast.BoolOp) and isinstance(test.op, ast.And):
        return False if any(_op_test(v, op) is False for v in test.values) else None
    if not (isinstance(test, ast.Compare) and isinstance(test.left, ast.Name)
            and test.left.id == "op" and len(test.ops) == 1):
        return None
    c, o = test.comparators[0], test.ops[0]
    if isinstance(c, ast.Constant):
        if isinstance(o, ast.Eq):
            return c.value == op
        if isinstance(o, ast.NotEq):
            return c.value != op
    if isinstance(c, (ast.Tuple, ast.List, ast.Set)):
        vals = {e.value for e in c.elts if isinstance(e, ast.Constant)}
        if isinstance(o, ast.In):
            return op in vals
        if isinstance(o, ast.NotIn):
            return op not in vals
    return None


def _keys_touched(fn, pname: str, op: str, mod=None, depth: int = 1) -> set:
    """fn の中で引数 pname（resolved / args）の鍵として読み書きされる文字列。

    ★ 他の op の枝は辿らない。★ 同じモジュールの関数へ `pname=pname` で丸ごと渡す委譲は
      1 段だけ辿る（_codegen_dedup_delete → _codegen_delete_rows）。"""
    if fn is None:
        return set()
    tree = ast.parse(textwrap.dedent(inspect.getsource(fn)))
    out: set = set()

    def visit(node):
        if isinstance(node, ast.If):
            r = _op_test(node.test, op)
            if r is not None:
                for s in (node.body if r else node.orelse):
                    visit(s)
                return
        if (isinstance(node, ast.Subscript) and isinstance(node.value, ast.Name)
                and node.value.id == pname and isinstance(node.slice, ast.Constant)):
            out.add(node.slice.value)
        if isinstance(node, ast.Call):
            f = node.func
            if (isinstance(f, ast.Attribute) and f.attr == "get" and isinstance(f.value, ast.Name)
                    and f.value.id == pname and node.args and isinstance(node.args[0], ast.Constant)):
                out.add(node.args[0].value)
            if depth and mod is not None and isinstance(f, ast.Name):
                kw = [k.arg for k in node.keywords
                      if isinstance(k.value, ast.Name) and k.value.id == pname]
                callee = getattr(mod, f.id, None)
                if kw and callable(callee) and getattr(callee, "__module__", "") == mod.__name__:
                    out.update(_keys_touched(callee, kw[0], op, mod, depth - 1))
        for ch in ast.iter_child_nodes(node):
            visit(ch)

    visit(tree)
    return {k for k in out if isinstance(k, str) and not k.startswith("_")}


def _verify_dispatch() -> dict:
    """verify_dsl_args の `if op == "X": _verify_x(...)` から op → 検査関数名。"""
    tree = ast.parse(textwrap.dedent(inspect.getsource(ailine.verify_dsl_args)))
    found: dict = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.If):
            continue
        for op in ailine.OP_SCHEMA:
            if _op_test(node.test, op) is not True:
                continue
            for st in node.body:
                for c in ast.walk(st):
                    if (isinstance(c, ast.Call) and isinstance(c.func, ast.Name)
                            and c.func.id.startswith("_verify_")):
                        found.setdefault(op, c.func.id)
    return found


def declared_items() -> set:
    """宣言から導いた (op, 項目) の集合。"""
    dispatch = _verify_dispatch()
    items = set()
    for op, slots in ailine.OP_SCHEMA.items():
        keys = set(slots)
        keys |= {k for k, _kind in ailine.OP_SUBJECT_SLOTS.get(op, ())}
        keys |= {f[1] for f in ailine._CONFIRM_FIELDS.get(op, ()) if not f[1].startswith("_")}
        keys |= _keys_touched(codegen.CODEGEN_BY_OP.get(op), "resolved_args", op, codegen)
        keys |= _keys_touched(ailine.POSTCONDITIONS.get(op), "args", op)
        vname = dispatch.get(op)
        vfn = (getattr(argcheck, vname, None) or getattr(ailine, vname, None)) if vname else None
        keys |= _keys_touched(vfn, "resolved", op)
        items |= {(op, k) for k in keys}
    return items


def _load(path=REGISTER) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def _register_items(reg) -> set:
    return {(e["op"], e["item"]) for e in reg["items"]}


def _counts(reg) -> dict:
    out = {c: 0 for c in CLASSES}
    for e in reg["items"]:
        out[e["class"]] += 1
    return out


def _cap_breaches(reg) -> list:
    """上限と数の食い違い（上限超え＝在庫が増えた／下回り＝上限を下げ忘れ）。"""
    n, caps = _counts(reg), reg["_meta"]["caps"]
    bad = []
    for c in ("B", "D"):
        if n[c] > caps[c]:
            bad.append(f"{c} が {n[c]} 件（上限 {caps[c]}）── 依頼と突き合わせない項目が増えた")
        elif n[c] < caps[c]:
            bad.append(f"{c} が {n[c]} 件に減った ── 上限（{caps[c]}）を {n[c]} に下げてください")
    return bad


@functools.lru_cache(maxsize=None)
def _functions_in(relpath: str) -> frozenset:
    p = REPO / relpath
    if not p.is_file():
        return frozenset()
    tree = ast.parse(p.read_text(encoding="utf-8"))
    return frozenset(n.name for n in ast.walk(tree)
                     if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)))


def _defined(relpath: str, func: str) -> bool:
    return func in _functions_in(relpath)


# --- ① 分母は宣言から・台帳と等号 --------------------------------------------------------

def test_every_declared_item_is_classified_and_nothing_else():
    reg = _load()
    declared, written = declared_items(), _register_items(reg)
    unclassified, stale = sorted(declared - written), sorted(written - declared)
    assert not unclassified and not stale, (
        "台帳と宣言がずれている:\n"
        + "".join(f"  台帳に無い（区分を書いてください）: {op}.{k}\n" for op, k in unclassified)
        + "".join(f"  宣言に無い（台帳から消してください）: {op}.{k}\n" for op, k in stale))


def test_the_denominator_is_not_empty():
    """★ 陽性対照: 導出が壊れて空集合になれば、等号は台帳も空で恒真になりうる。"""
    d = declared_items()
    assert len(d) >= 60
    assert ("SORT", "order") in d and ("DELETE_ROWS", "count") in d
    # 他の op の枝を数えていない（数えると生える幻の項目）
    assert ("BOLD", "color") not in d and ("ADD_ROW", "count") not in d


def test_a_new_item_in_the_schema_turns_this_red(monkeypatch):
    """変異: OP_SCHEMA に架空の項目・架空の op を足したつもりの検体で、未分類として出ること。"""
    schema = dict(ailine.OP_SCHEMA)
    schema["SORT"] = tuple(schema["SORT"]) + ("架空の項目",)
    schema["架空の操作"] = ("何か",)
    monkeypatch.setattr(ailine, "OP_SCHEMA", schema)
    missing = declared_items() - _register_items(_load())
    assert ("SORT", "架空の項目") in missing
    assert ("架空の操作", "何か") in missing


def test_a_new_key_read_by_codegen_turns_this_red(monkeypatch):
    """変異: 生成部が新しい鍵を読み始めた（＝文書に効く項目が増えた）ら未分類として出ること。"""
    def _codegen_sort_mutant(*, op, resolved_args, book_meta, use_formula, headers, first_sheet,
                             header_row, hr0, wrap):
        return resolved_args.get("架空の鍵")

    table = dict(codegen.CODEGEN_BY_OP)
    table["SORT"] = _codegen_sort_mutant
    monkeypatch.setattr(codegen, "CODEGEN_BY_OP", table)
    assert ("SORT", "架空の鍵") in declared_items() - _register_items(_load())


# --- ② B と D は縮める在庫 -----------------------------------------------------------------

def test_the_unmatched_stock_only_shrinks():
    reg = _load()
    bad = [f"{e['op']}.{e['item']}: 区分『{e['class']}』は A/B/C/D のどれでもない"
           for e in reg["items"] if e["class"] not in CLASSES]
    assert not bad, "\n".join(bad)
    assert not _cap_breaches(reg), "\n".join(_cap_breaches(reg))


def test_the_stock_guard_rings_both_ways():
    """変異: D を 1 つ増やした台帳は赤・在庫（D か B）を 1 つ減らして上限を据え置いた台帳も赤。

    ★ D が 0 になった（2026-10-01）後も鳴ることを確かめるため、減らす側は在庫の残っている
      区分を選ぶ（D が空なら B）。
    ★ B も 0 になった（2026-10-01・同日 2 度目）後は減らせる在庫が無い ── 上限を 1 つ上げた
      台帳（＝数が上限を下回った形）で「下げて」が鳴ることを確かめる。"""
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
    # B を増やした台帳も赤（上限 0 の B が 1 件でも生えたら鳴る）
    grow_b = json.loads(json.dumps(reg))
    next(e for e in grow_b["items"] if e["class"] == "A")["class"] = "B"
    assert any("増えた" in b for b in _cap_breaches(grow_b))


def test_the_history_ends_at_the_current_counts():
    """台帳の履歴の最後の行が、今の数と一致する（履歴を書かずに区分だけ動かさない）。"""
    reg = _load()
    assert reg["_meta"]["history"][-1]["counts"] == {
        c: n for c, n in _counts(reg).items() if n}


# --- ③ 根拠が実在する ----------------------------------------------------------------------

def test_every_basis_points_at_a_function_that_exists():
    reg = _load()
    rotten, empty = [], []
    for e in reg["items"]:
        if not e.get("basis"):
            empty.append(f"{e['op']}.{e['item']}")
        for b in e.get("basis") or ():
            path, _, func = b.rpartition(":")
            if not _defined(path, func):
                rotten.append(f"{e['op']}.{e['item']}: {b}")
    assert not empty, "根拠が空: " + "、".join(empty)
    assert not rotten, "根拠の関数が見つからない:\n  " + "\n  ".join(rotten)


def test_the_basis_check_catches_a_rotten_name():
    """変異: 存在しない関数名は見つからない（番人が何でも通す恒真になっていない）。"""
    assert _defined("src/ailine_core/argcheck.py", "_verify_sort")
    assert not _defined("src/ailine_core/argcheck.py", "_verify_sort_架空")
    assert not _defined("src/ailine_core/存在しない.py", "_verify_sort")
