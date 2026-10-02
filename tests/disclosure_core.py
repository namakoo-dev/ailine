# -*- coding: utf-8 -*-
"""開示の盤 ── 「黙って残す・黙って外す・黙って置き換える・画面だけに言う」在り処を、書かれた形から導く（2026-10-02）。

★★ なぜ要るか（盲検の欠陥の形 4: 開示の欠落）: 同じ形が 9 件出て、毎回**見つけた 1 つ**を直していた。
  関所で止まった run が `.out` を黙って残す家系は 3 度（ae85f5f・d1e2b99）・⚠ が画面だけで冊に残らない
  （a70b5d3）・2 冊照合が history に残らない（4a61518）。記録に書かれた根は
  「出口の一覧を宣言から導かず、見つけた出口を 1 つずつ塞いでいる」。
  ★ だから在り処を**走査で数え**、台帳（tests/disclosure_register.json）と等号で縛る。
    新しい出口・⚠・外す所・置き換える所を書いたら、**分類するまで赤**。開示なしは減る向きだけ。

★ 4 つの軸（brief の (a)〜(d)）。**何を在り処と数えるか**は下の宣言（口の名簿・上限の名簿）から導く:

  (a) 終わり方   cmd_* から届く関数の出口（`return <数>` / `return EXIT_*` / `sys.exit` / `SystemExit`）。
        no_out    その出口より前（行の順）に、出力を作る呼び出しが無い ── 開示すべき残り物がまだ無い
        routed    出力を作った後で、履歴の口（`_finish_run` 系・`_record_*`・`record_made_book`）を通る
        unrouted  出力を作った後で、履歴の口を通らない  ← **開示なし**（台帳に理由を書くまで赤）
      ★ 履歴に残す＝「次の run/人が、これが道具の作った物と分かる」。画面での名指しは named で別に数える。
  (b) ⚠ の行き先   `⚠` を含む文を書く所。
        carried       リスト・返り値・`_say` に載る（冊と履歴の口へ運ばれうる）
        print_no_out  `print` 直書き。ただしその前に出力を作っていない（断り・前提の破れ）── 運ぶ冊が無い
        print_after_out `print` 直書きで、出力を作った後  ← **開示なし**（画面だけ）
  (c) 黙って外す   上限の名簿（走査で導く: `MAX|LIMIT|CAP` を含む名前の数値定数）を使う所。
        said      同じ関数が、外した/届かなかったことを言う語（上限・切り詰・先頭・見ていない…）を書く
        silent    言わない  ← **開示なし**
  (d) 置き換える   既存の物を置き換える書き込み（`shutil.copy2`・`os.replace`・`.save`・LO/PDF 書き出し 等）。
        gated     同じ関数が、書く前に関所（`exists()`・`overwrite`・`_own_*_status`・`refuse_if_*`）を通る
        ungated   通らない  ← **開示なし**
      ★ ungated には「一時ディレクトリへの書き込み（置き換える相手が無い）」も含まれる ── 台帳に理由で書く。

★★ これは**構文の近似**であって、実行の追跡ではない ── 「行の順で前」「同じ関数の中」で数える。
  取りこぼしは「台帳に載らない」だけで嘘を増やさないが、**効かない所は陽性対照で名指しする**
  （直す前の版で鳴り、直した版で鳴らないこと ── tests/test_disclosure_is_observed.py）。
★ 走査は AST（docstring とコメントは文言でない・連結した文は 1 文として読む）。
★ なぜ tests/ に在るか: scripts/ から tests/ の中身を import すると素の環境が弾く（3 度踏んだ）。
"""
from __future__ import annotations

import ast
import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

from _product_source import product_files  # noqa: E402

REGISTER = REPO / "tests" / "disclosure_register.json"

_JA = re.compile(r"[぀-ヿ一-鿿]")

# ---------------------------------------------------------------------------------------------
# 宣言（口の名簿）── ここだけが「どの呼び出しを何と見るか」を決める。足したら台帳が動く。
# ---------------------------------------------------------------------------------------------

#: 出力（人の目に残る物）を作る呼び出し。attr 名または関数名。
OUT_MAKERS = frozenset({
    "copy2", "copyfile", "move", "save", "write_bytes",
    "basrun_apply", "apply_dsl_step", "atomic_replace_inplace",
    "write_quarantined_xlsx", "_write_csv_output", "_soffice_to_pdf",
})
#: `os.replace(...)` は文字列の `.replace` と区別する（基底が os の時だけ）。
_OS_REPLACE = ("os", "replace")

#: 出力（既に作られた物）を引数で受け取る関数の引数名。これを持つ関数は「作った後」にいる。
OUT_PARAMS = frozenset({"out_book", "out_path", "apply_target"})

#: 履歴に残す口（これを通れば、次の run / 人が「道具が作った物」と分かる）。
HISTORY_MOUTHS = frozenset({
    "_finish_run", "_finish_failed_apply", "_finish_gated",
    "_record_history", "_record_side_command_history", "record_made_book",
    "append_history", "_record_csv_conversion_history", "_record_csv_export_history",
})

#: 画面で残したものを名指しする口。
NAMING_MOUTHS = frozenset({"_untouched_original_line", "_untouched_claim"})

#: 置き換えの関所（書く前に通る）。
GATE_NAMES = frozenset({
    "_export_out_path", "_export_csv_out_path", "_own_csv_output_status", "_refuse_output_conflict", "_refuse_edited_output",
    "refuse_if_output_is_someone_elses", "_confirm_overwrite_or_gate", "own_output_mark",
    "_own_extract_output_status", "_csv_output_edited_since", "_is_our_scratch_output",
})
_GATE_WORDS = re.compile(r"overwrite|exists")

#: 外したこと・届かなかったことを言う語（同じ関数の文に載っていれば said）。
SAID_WORDS = re.compile(
    r"上限|切り詰|切れ|打ち切|先頭|以降|確認していない|見ていない|届かな|読めませんでした|読み飛ばし|"
    r"除いて|除外|除き|含め|行目より下|超え|多すぎ|削っ|落と|省略|詰め|ほか")

_CAP_NAME = re.compile(r"(MAX|LIMIT|CAP)")


# ---------------------------------------------------------------------------------------------
# 走査の共通部品
# ---------------------------------------------------------------------------------------------

def _docstring_ids(tree) -> set:
    ids = set()
    for n in ast.walk(tree):
        if isinstance(n, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            body = n.body
            if (body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant)
                    and isinstance(body[0].value.value, str)):
                ids.add(id(body[0].value))
    return ids


def _render(n):
    """文字列を作る式を 1 つの文に描く（式の部分は `{…}`・描けなければ None）。"""
    if isinstance(n, ast.Constant) and isinstance(n.value, str):
        return n.value
    if isinstance(n, ast.JoinedStr):
        out = ""
        for v in n.values:
            if isinstance(v, ast.Constant):
                out += str(v.value)
            else:
                out += "{" + ast.unparse(v.value)[:30] + "}"
        return out
    if isinstance(n, ast.BinOp) and isinstance(n.op, ast.Add):
        a, b = _render(n.left), _render(n.right)
        if a is not None and b is not None:
            return a + b
    return None


def _call_name(c: ast.Call) -> str:
    f = c.func
    return f.attr if isinstance(f, ast.Attribute) else getattr(f, "id", "")


def _is_out_maker(c: ast.Call) -> bool:
    name = _call_name(c)
    if name in OUT_MAKERS:
        return True
    if name == "open" and isinstance(c.func, ast.Name):      # open(path, "wb") ── 書く開き方だけ
        mode = c.args[1] if len(c.args) > 1 else next((k.value for k in c.keywords if k.arg == "mode"), None)
        if isinstance(mode, ast.Constant) and isinstance(mode.value, str) and mode.value[:1] in "wax":
            return True
    f = c.func
    return (isinstance(f, ast.Attribute) and f.attr == _OS_REPLACE[1]
            and isinstance(f.value, ast.Name) and f.value.id == _OS_REPLACE[0])


class _Func:
    """1 つの関数の「自分の本体」（入れ子の def・lambda・class の中は含めない）。"""

    def __init__(self, file: str, node):
        self.file, self.node, self.name = file, node, node.name
        self.nodes: list = []
        stack = list(node.body)
        while stack:
            c = stack.pop()
            self.nodes.append(c)
            # lambda の中の呼び出しは**この関数の呼び出し**（`under_run_lock(lambda: _cmd_run_body(a))`）。
            if isinstance(c, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                continue
            stack.extend(ast.iter_child_nodes(c))
        self.calls = [c for c in self.nodes if isinstance(c, ast.Call)]
        self.call_names = {_call_name(c) for c in self.calls}
        self.make_lines = sorted(c.lineno for c in self.calls if _is_out_maker(c))
        self.first_make = self.make_lines[0] if self.make_lines else None
        a = node.args
        self.takes_out = any(p.arg in OUT_PARAMS for p in a.posonlyargs + a.args + a.kwonlyargs)

    def made_before(self, line: int) -> bool:
        # 出力を引数で受け取る関数（`_finish_apply(a, book, out_book, …)`）は、呼び出し側が既に作っている
        return self.takes_out or (self.first_make is not None and self.first_make < line)


def _parse(files):
    funcs: list = []
    trees: dict = {}
    for f in files:
        try:
            rel = str(Path(f).resolve().relative_to(REPO)).replace(chr(92), "/")
        except ValueError:                # repo の外（変異試験の一時ファイル）── 名前だけで持つ
            rel = Path(f).name
        tree = ast.parse(Path(f).read_bytes().decode("utf-8"))
        trees[rel] = tree
        for n in ast.walk(tree):
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
                funcs.append(_Func(rel, n))
    return funcs, trees


def _records_closure(funcs) -> set:
    """履歴の口を（直接・間接に）呼ぶ関数の名前。名前だけの呼び出しグラフ（衝突は保守側＝通す）。"""
    records = set(HISTORY_MOUTHS)
    changed = True
    while changed:
        changed = False
        for fn in funcs:
            if fn.name in records:
                continue
            if fn.call_names & records:
                records.add(fn.name)
                changed = True
    return records


def entry_names(trees) -> set:
    """サブコマンドの入口 ── 宣言から導く（`<parser>.set_defaults(func=X)` の X）。名簿を手で持たない。"""
    out = set()
    for tree in trees.values():
        for n in ast.walk(tree):
            if isinstance(n, ast.Call) and _call_name(n) == "set_defaults":
                for k in n.keywords:
                    if k.arg == "func" and isinstance(k.value, ast.Name):
                        out.add(k.value.id)
    return out


def _cmd_reach(funcs, entries=None) -> set:
    """サブコマンドの入口から名前で届く関数（出口を数える範囲）。"""
    by_name: dict = {}
    for fn in funcs:
        by_name.setdefault(fn.name, []).append(fn)
    reach = set(entries) if entries else {fn.name for fn in funcs if fn.name.startswith("cmd_")}
    frontier = list(reach)
    while frontier:
        name = frontier.pop()
        for fn in by_name.get(name, []):
            for callee in fn.call_names:
                if callee in by_name and callee not in reach:
                    reach.add(callee)
                    frontier.append(callee)
    return reach


# ---------------------------------------------------------------------------------------------
# (a) 終わり方
# ---------------------------------------------------------------------------------------------

def _int_exit(v) -> str | None:
    if isinstance(v, ast.Constant) and isinstance(v.value, int) and not isinstance(v.value, bool):
        return str(v.value)
    if isinstance(v, ast.Name) and v.id.startswith("EXIT_"):
        return v.id
    return None


def _assigned_from(fn: _Func, name: str, records: set) -> bool:
    """`name = <履歴の口を通る呼び出し>` が在るか（`_rc = report_postcondition(...)` ... `return _rc`）。"""
    for n in fn.nodes:
        if isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id == name for t in n.targets):
            v = n.value
            if isinstance(v, ast.Call) and _call_name(v) in records:
                return True
    return False


_TERMINATORS = (ast.Return, ast.Raise, ast.Continue, ast.Break)
_SUBLISTS = ("body", "orelse", "finalbody")


def _mouth_in(stmt, records: set) -> bool:
    """この文を**通り抜けた先**で、履歴の口が呼ばれているか。
    ★ 「その出口で終わる枝」の中の呼び出しは数えない（`if dry: _finish_run(); return 0` の後ろで
      `return 3` しても、それは別の枝）── 行の順だけで見ると、片方の枝が全部を routed にしてしまう。"""
    if isinstance(stmt, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
        return False
    compound = False
    for field in _SUBLISTS:
        sub = getattr(stmt, field, None)
        if isinstance(sub, list) and sub and isinstance(sub[0], ast.stmt):
            compound = True
            if isinstance(sub[-1], _TERMINATORS):
                continue                      # この枝は出口で終わる ── 通り抜けない
            if any(_mouth_in(s, records) for s in sub):
                return True
    if isinstance(stmt, ast.Try):
        for h in stmt.handlers:
            compound = True
            if h.body and isinstance(h.body[-1], _TERMINATORS):
                continue
            if any(_mouth_in(s, records) for s in h.body):
                return True
    if compound:
        for field in ("test", "iter", "items", "subject"):
            part = getattr(stmt, field, None)
            parts = part if isinstance(part, list) else ([part] if part is not None else [])
            for pnode in parts:
                if any(isinstance(c, ast.Call) and _call_name(c) in records for c in ast.walk(pnode)):
                    return True
        return False
    return any(isinstance(c, ast.Call) and _call_name(c) in records for c in ast.walk(stmt))


def _routed_before(fn: _Func, node, records: set) -> bool:
    """出口 node に至る道の**手前**（同じ文の並び・外側の文の並びの、先に在る文）で履歴の口を通ったか。"""
    parent = {}
    stack = [fn.node]
    while stack:
        p = stack.pop()
        for c in ast.iter_child_nodes(p):
            parent[id(c)] = p
            if not isinstance(c, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda, ast.ClassDef)):
                stack.append(c)
    child = node
    while id(child) in parent:
        anc = parent[id(child)]
        # `if _answer_before_asking(...): return 3` ── 条件の呼び出しが後始末を済ませて True を返す形
        test = getattr(anc, "test", None)
        if test is not None and any(isinstance(c, ast.Call) and _call_name(c) in records
                                    for c in ast.walk(test)):
            return True
        lists = [getattr(anc, f, None) for f in _SUBLISTS]
        if isinstance(anc, ast.Try):
            lists += [h.body for h in anc.handlers]
            # `except` の中の出口は、`try` の本体が（途中まで）走った後 ── 本体で通した口も数える
            if child in anc.handlers and any(_mouth_in(s, records) for s in anc.body):
                return True
        for lst in lists:
            if isinstance(lst, list) and child in lst:
                for s in lst[:lst.index(child)]:
                    if _mouth_in(s, records):
                        return True
        child = anc
        if anc is fn.node:
            break
    return False


_EXPR_EXIT = (ast.Attribute, ast.Name, ast.Subscript, ast.IfExp, ast.BoolOp)


def scan_exits(funcs, reach, records) -> list:
    out = []
    for fn in funcs:
        if fn.name not in reach:
            continue
        # 終了コードを返す関数か（数の出口を 1 つでも持つ）── 持つなら `return e.exit_code` のような式も出口
        exit_func = any(isinstance(n, ast.Return) and n.value is not None and _int_exit(n.value) is not None
                        for n in fn.nodes)
        for n in fn.nodes:
            kind = None
            routed = False
            line = getattr(n, "lineno", 0)
            if isinstance(n, ast.Return) and n.value is not None:
                v = n.value
                code = _int_exit(v)
                if code is not None:
                    kind = "int"
                elif isinstance(v, ast.Call) and _call_name(v) in records and (
                        _call_name(v).startswith("_finish") or _call_name(v).startswith("_refuse")):
                    kind, routed = "mouth", True
                elif isinstance(v, ast.Name) and _assigned_from(fn, v.id, records):
                    kind, routed = "carried", True
                elif exit_func and isinstance(v, _EXPR_EXIT) and not (
                        isinstance(v, ast.Constant) or (isinstance(v, ast.Name) and v.id in ("True", "False", "None"))):
                    # ★ `return e.exit_code` ── 数の出口の一覧に出ない出口（ae85f5f の 3 件目はこれだった）
                    kind = "expr"
            elif isinstance(n, ast.Call) and _call_name(n) == "exit" and isinstance(n.func, ast.Attribute):
                kind = "sys_exit"
            elif isinstance(n, ast.Raise) and n.exc is not None and "SystemExit" in ast.unparse(n.exc):
                kind = "sys_exit"
            if kind is None:
                continue
            made = fn.made_before(line)
            if kind in ("int", "sys_exit", "expr"):
                # 手前の文で履歴の口を通っていれば、この出口はその後始末の後
                routed = _routed_before(fn, n, records)
            named = any(c.lineno < line and _call_name(c) in NAMING_MOUTHS for c in fn.calls)
            if not made:
                cls = "no_out"
            elif routed:
                cls = "routed"
            else:
                cls = "unrouted"
            out.append({"axis": "a", "file": fn.file, "func": fn.name, "class": cls,
                        "kind": kind, "line": line, "named": named})
    return out


_CATCH_ALL = {"BaseException", "KeyboardInterrupt"}


def _makes_out_closure(funcs) -> set:
    """出力を作る呼び出しを（直接・間接に）含む関数の名前。"""
    makes = {fn.name for fn in funcs if fn.make_lines}
    changed = True
    while changed:
        changed = False
        for fn in funcs:
            if fn.name not in makes and fn.call_names & makes:
                makes.add(fn.name)
                changed = True
    return makes


def _guards_exceptions(fn: _Func, records: set) -> bool:
    """`except BaseException/KeyboardInterrupt/裸 except` が、履歴の口を通ってから投げ直しているか。"""
    for n in fn.nodes:
        if not isinstance(n, ast.Try):
            continue
        for h in n.handlers:
            t = h.type
            names = ({t.id} if isinstance(t, ast.Name) else
                     {e.id for e in t.elts if isinstance(e, ast.Name)} if isinstance(t, ast.Tuple) else set())
            if t is not None and not (names & _CATCH_ALL):
                continue
            calls = {_call_name(c) for s in h.body for c in ast.walk(s) if isinstance(c, ast.Call)}
            if calls & records and any(isinstance(s, ast.Raise) for s in h.body):
                return True
    return False


def scan_raises(funcs, entries, records) -> list:
    """例外（Ctrl-C・想定外のエラー）で出る道 ── 入口ごとに、出力を作りうるのに、例外の出口が
    履歴の口を通らないか。★ `return` の出口だけ数えると、例外の出口は**数えられない**（盲検の形 4 の 4 件目）。
    入口から届く関数のどれか（出力を作る側）が、投げ直す前に後始末を通せば exc_guarded。"""
    by_name: dict = {}
    for fn in funcs:
        by_name.setdefault(fn.name, []).append(fn)
    makes = _makes_out_closure(funcs)
    out = []
    for e in sorted(entries):
        if e not in makes:
            continue
        seen, frontier, guarded = {e}, [e], False
        while frontier and not guarded:
            name = frontier.pop()
            for fn in by_name.get(name, []):
                if fn.name in makes and _guards_exceptions(fn, records):
                    guarded = True
                    break
                for callee in fn.call_names:
                    if callee in by_name and callee not in seen:
                        seen.add(callee)
                        frontier.append(callee)
        fn0 = by_name[e][0]
        out.append({"axis": "a", "file": fn0.file, "func": e,
                    "class": "exc_guarded" if guarded else "exc_unguarded",
                    "kind": "raise", "line": fn0.node.lineno})
    return out


# ---------------------------------------------------------------------------------------------
# (b) ⚠ の行き先
# ---------------------------------------------------------------------------------------------

def scan_warnings(funcs, trees) -> list:
    out = []
    by_file_funcs: dict = {}
    for fn in funcs:
        by_file_funcs.setdefault(fn.file, []).append(fn)
    for rel, tree in trees.items():
        doc = _docstring_ids(tree)
        parent = {}
        for p in ast.walk(tree):
            for c in ast.iter_child_nodes(p):
                parent[id(c)] = p
        owner = {}
        for fn in by_file_funcs.get(rel, []):
            for n in fn.nodes:
                owner.setdefault(id(n), fn)
        for n in ast.walk(tree):
            if not isinstance(n, (ast.Constant, ast.JoinedStr, ast.BinOp)) or id(n) in doc:
                continue
            p = parent.get(id(n))
            if p is not None and isinstance(p, (ast.BinOp, ast.JoinedStr)) and _render(p) is not None:
                continue
            text = _render(n)
            if text is None or "⚠" not in text or not _JA.search(text):
                continue
            # 最も近い呼び出し（この文を引数に持つ）
            q, call = n, None
            while id(q) in parent:
                q = parent[id(q)]
                if isinstance(q, ast.Call):
                    call = q
                    break
                if isinstance(q, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda)):
                    break
            fn = owner.get(id(n))
            fname = fn.name if fn else "<module>"
            if call is not None and _call_name(call) == "print":
                made = bool(fn and fn.made_before(n.lineno))
                cls = "print_after_out" if made else "print_no_out"
            else:
                cls = "carried"
            out.append({"axis": "b", "file": rel, "func": fname, "class": cls,
                        "line": n.lineno, "text": text[:60]})
    return out


# ---------------------------------------------------------------------------------------------
# (c) 黙って外す
# ---------------------------------------------------------------------------------------------

def cap_names(trees) -> set:
    """上限の名簿 ── モジュール直下の `…MAX…/…LIMIT…/…CAP…` という名前の数値定数（走査で導く）。"""
    names = set()
    for tree in trees.values():
        for n in tree.body:
            if isinstance(n, ast.Assign) and isinstance(n.value, ast.Constant) \
                    and isinstance(n.value.value, (int, float)) and not isinstance(n.value.value, bool):
                for t in n.targets:
                    if isinstance(t, ast.Name) and _CAP_NAME.search(t.id):
                        names.add(t.id)
    return names


#: 「この行/値は読まない」と決める `if <述語>: continue|break` の述語に現れる語（合計行・空行・重複・既読…）。
_SKIP_TEST = re.compile(r"total|blank|empty|exclud|skip|dup|seen|stop|merged|hidden|unread|drop|ignore", re.I)


def _said_in(fn: _Func) -> bool:
    for n in fn.nodes:
        if isinstance(n, (ast.Constant, ast.JoinedStr)):
            t = _render(n)
            if t and _JA.search(t) and SAID_WORDS.search(t):
                return True
    return False


def scan_caps(funcs, trees) -> list:
    caps = cap_names(trees)
    out = []
    # 読み飛ばし（合計行の除外・空行で止まる・重複の畳み）── `if <述語>: continue|break`
    for fn in funcs:
        said = None
        for n in fn.nodes:
            if (isinstance(n, ast.If) and n.body and len(n.body) <= 2
                    and isinstance(n.body[-1], (ast.Continue, ast.Break))
                    and _SKIP_TEST.search(ast.unparse(n.test))):
                if said is None:
                    said = _said_in(fn)
                out.append({"axis": "c", "file": fn.file, "func": fn.name,
                            "class": "said" if said else "silent", "cap": "skip",
                            "line": n.lineno})
    for fn in funcs:
        used = set()
        for n in fn.nodes:
            if isinstance(n, ast.Name) and n.id in caps:
                used.add(n.id)
            elif isinstance(n, ast.Attribute) and n.attr in caps:
                used.add(n.attr)
        if not used:
            continue
        said = False
        for n in fn.nodes:
            if isinstance(n, (ast.Constant, ast.JoinedStr)):
                t = _render(n)
                if t and _JA.search(t) and SAID_WORDS.search(t):
                    said = True
                    break
        for cap in sorted(used):
            out.append({"axis": "c", "file": fn.file, "func": fn.name,
                        "class": "said" if said else "silent", "cap": cap, "line": fn.node.lineno})
    return out


# ---------------------------------------------------------------------------------------------
# (d) 置き換える
# ---------------------------------------------------------------------------------------------

def scan_writes(funcs) -> list:
    out = []
    for fn in funcs:
        if not fn.make_lines:
            continue
        has_gate = bool(fn.call_names & GATE_NAMES)
        for n in fn.nodes:
            if isinstance(n, ast.Attribute) and _GATE_WORDS.search(n.attr):
                has_gate = True
            elif isinstance(n, ast.Name) and _GATE_WORDS.search(n.id):
                has_gate = True
            elif isinstance(n, ast.keyword) and n.arg and _GATE_WORDS.search(n.arg):
                has_gate = True
        for c in fn.calls:
            if not _is_out_maker(c):
                continue
            out.append({"axis": "d", "file": fn.file, "func": fn.name,
                        "class": "gated" if has_gate else "ungated",
                        "maker": _call_name(c), "line": c.lineno})
    return out


# ---------------------------------------------------------------------------------------------
# 全体
# ---------------------------------------------------------------------------------------------

def scan(files=None) -> list:
    """4 軸の在り処を [{axis, file, func, class, line, …}] で返す（製品コード全体）。"""
    funcs, trees = _parse(files if files is not None else product_files())
    records = _records_closure(funcs)
    entries = entry_names(trees)
    reach = _cmd_reach(funcs, entries)
    return (scan_exits(funcs, reach, records) + scan_raises(funcs, entries, records)
            + scan_warnings(funcs, trees) + scan_caps(funcs, trees) + scan_writes(funcs))


#: 開示なしの class（台帳に理由を書くまで赤）。
UNDISCLOSED = {"a": ("unrouted", "exc_unguarded"), "b": ("print_after_out",), "c": ("silent",),
               "d": ("ungated",)}


def tally(found: list, only_undisclosed: bool = True) -> dict:
    """{(axis, file, func, class): 件数}。★ 行番号は鍵に入れない（行が動くだけで赤くしない）。"""
    out: dict = {}
    for x in found:
        if only_undisclosed and x["class"] not in UNDISCLOSED[x["axis"]]:
            continue
        k = (x["axis"], x["file"], x["func"], x["class"])
        out[k] = out.get(k, 0) + 1
    return out


def summary(found: list) -> dict:
    """{axis: {class: 件数}}（報告用）。"""
    out: dict = {}
    for x in found:
        out.setdefault(x["axis"], {})
        out[x["axis"]][x["class"]] = out[x["axis"]].get(x["class"], 0) + 1
    return out


def load_register(path: Path = REGISTER) -> dict:
    data = json.loads(Path(path).read_bytes().decode("utf-8"))
    return {(e["axis"], e["file"], e["func"], e["class"]): e for e in data["undisclosed"]}


def compare(found_tally: dict, register: dict) -> dict:
    """台帳と実装の差。★ 等号で縛る ── 足した在り処は理由を書くまで赤・減らしたら台帳も減らす
    （台帳に残った項目も赤＝死んだ免除）・**件数が増える向きは常に赤**（減る向きだけが前進）。"""
    unlisted = {k: n for k, n in found_tally.items() if k not in register}
    wrong_count = {k: (found_tally[k], register[k]["count"]) for k in found_tally
                   if k in register and found_tally[k] != register[k]["count"]}
    stale = [k for k in register if k not in found_tally]
    return {"unlisted": unlisted, "wrong_count": wrong_count, "stale": stale}


if __name__ == "__main__":
    f = scan()
    print(json.dumps(summary(f), ensure_ascii=False, indent=1))
    for k, n in sorted(tally(f).items()):
        print(n, *k)
