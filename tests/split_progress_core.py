# -*- coding: utf-8 -*-
"""単一ファイルを割る作業の**分母を出す**芯（測定器。番人の数えもここを通る）。

★ 入口は `scripts/split_progress.py`（薄い）。中身をここ（tests/）に置くのは、番人が
  scripts/ を import すると素の環境で弾かれるため（repo の道具は tests/ に置く）。

★★ 2026-09-23 に測り方を直した（設計レビューで第 1 版の地図が崩れた件）:
  初版は「I/O を呼ぶ名前」だけで純かどうかを決め、**モジュール定数を見ていなかった**。
  `load_vocab` は `VOCAB_FILE`（~/.ailine の下の Path）を直に読むのに「移せる」側に入り、
  そのまま ailine_core へ移すと**試験の切り離しが黙って外れ、本物のホームに書くのに緑**になる。
  → 次の 3 つを「本体に残る理由」として数える:
    ① I/O を呼ぶ（IMPURE の名前）
    ② **試験が差し替える名前**を読む（setattr の的 ＋ ホームの下の Path ── 名簿は試験と本体から導く）
    ③ 上の①②に当たる本体の関数を**呼ぶ**（推移的に。初版は依存を並べるだけで辿っていなかった）
★ まだ見ていないもの（粗さ）: 動的な呼び出し（getattr・globals）・ailine_core 側の関数が
  内部で I/O をすること（import 名は「連れて行けるもの」として数えるだけ）。
"""
from __future__ import annotations

import ast
from pathlib import Path

import sys as _sys

_sys.path.insert(0, str(Path(__file__).resolve().parent))
from _product_source import MAIN_FILE  # noqa: E402

REPO = Path(__file__).resolve().parent.parent
MAIN = MAIN_FILE          # ★ 本体というファイルそのものを測る（場所を書くのは芯の 1 か所）
CORE = REPO / "src" / "ailine_core"

#: 触っていたら「純ロジックではない」と見なす名前（呼び出し名・属性名で照合）
IMPURE = {
    "open", "print", "input", "exit", "Path", "subprocess", "load_workbook",
    "Workbook", "run", "copy2", "move", "unlink", "mkdir", "rmtree",
    "write_text", "write_bytes", "read_text", "read_bytes", "urlopen",
    "system", "popen", "sleep", "chat_json",
}


def _called_names(fn: ast.AST) -> set:
    out = set()
    for x in ast.walk(fn):
        if isinstance(x, ast.Call):
            f = x.func
            if isinstance(f, ast.Name):
                out.add(f.id)
            elif isinstance(f, ast.Attribute):
                out.add(f.attr)
        elif isinstance(x, ast.Attribute):
            out.add(x.attr)
    return out


def _loaded_names(fn: ast.AST) -> set:
    return {x.id for x in ast.walk(fn) if isinstance(x, ast.Name)}


def module_table(tree: ast.Module) -> tuple:
    """トップレベルの名前を {名前: 種類}（def / const / import）と {名前: ノード} で。"""
    kinds, nodes = {}, {}
    for n in tree.body:
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            kinds[n.name], nodes[n.name] = "def", n
        elif isinstance(n, (ast.Assign, ast.AnnAssign, ast.AugAssign)):
            ts = n.targets if isinstance(n, ast.Assign) else [n.target]
            for t in ts:
                for x in ast.walk(t):
                    if isinstance(x, ast.Name):
                        kinds[x.id], nodes[x.id] = "const", n
        elif isinstance(n, (ast.Import, ast.ImportFrom)):
            for a in n.names:
                nm = (a.asname or a.name).split(".")[0]
                kinds[nm], nodes[nm] = "import", n
    return kinds, nodes


def patched_names(tests_dir: Path = REPO / "tests") -> set:
    """試験が `ailine` に差し替える名前（setattr の的・直接代入）。★ 名簿は手で書かない。"""
    out = set()
    for p in sorted(tests_dir.rglob("*.py")):
        try:
            tree = ast.parse(p.read_bytes().decode("utf-8"))
        except SyntaxError:
            continue
        for n in ast.walk(tree):
            if isinstance(n, ast.Call) and ast.unparse(n.func).endswith("setattr") and len(n.args) >= 2:
                a0, a1 = n.args[0], n.args[1]
                if isinstance(a0, ast.Name) and a0.id == "ailine" and isinstance(a1, ast.Constant):
                    out.add(str(a1.value))
                elif (isinstance(a0, ast.Constant) and isinstance(a0.value, str)
                      and a0.value.startswith("ailine.") and a0.value.count(".") == 1):
                    out.add(a0.value.split(".", 1)[1])
            elif isinstance(n, ast.Assign):
                for t in n.targets:
                    if (isinstance(t, ast.Attribute) and isinstance(t.value, ast.Name)
                            and t.value.id == "ailine"):
                        out.add(t.attr)
    return out


def home_names() -> set:
    """ホーム（~/.ailine）の下の Path を持つ本体の大域名。★ 本体を import して導く。"""
    import sys
    sys.path.insert(0, str(REPO / "src"))
    sys.path.insert(0, str(REPO / "tests"))
    import ailine  # noqa: E402
    from _home_isolation import home_bound_paths  # noqa: E402
    return set(home_bound_paths(ailine))


def survey(patched: set | None = None, home: set | None = None) -> dict:
    """本体のトップレベル関数を 1 つずつ、残る理由と連れて行くものつきで。"""
    src = MAIN.read_bytes().decode("utf-8")
    tree = ast.parse(src)
    kinds, nodes = module_table(tree)
    patched = patched_names() if patched is None else patched
    home = home_names() if home is None else home
    held = patched | home
    funcs = {k: v for k, v in nodes.items()
             if kinds[k] == "def" and isinstance(v, (ast.FunctionDef, ast.AsyncFunctionDef))}

    info = {}
    for name, fn in funcs.items():
        loaded = _loaded_names(fn)
        info[name] = {
            "name": name,
            "lines": fn.end_lineno - fn.lineno + 1,
            "io": sorted(_called_names(fn) & IMPURE),
            "held": sorted((loaded & held) - {name}),
            "calls": sorted({x for x in loaded if kinds.get(x) == "def" and x != name}),
            "consts": sorted({x for x in loaded if kinds.get(x) == "const"}),
            "imports": sorted({x for x in loaded if kinds.get(x) == "import"}),
            "patched_itself": name in held,
        }
    # ③ 推移: 残る関数を呼ぶ関数も残る（不動点まで）
    stays = {k for k, r in info.items() if r["io"] or r["held"] or r["patched_itself"]}
    changed = True
    while changed:
        changed = False
        for k, r in info.items():
            if k not in stays and any(c in stays for c in r["calls"]):
                stays.add(k)
                changed = True
    for k, r in info.items():
        r["pure"] = k not in stays
    return {"src": src, "funcs": info, "kinds": kinds}


def pure_logic(result: dict) -> tuple:
    """(行数, 関数の数) ── 本体に残っている純ロジック。"""
    pure = [r for r in result["funcs"].values() if r["pure"]]
    return sum(r["lines"] for r in pure), len(pure)


def core_lines() -> tuple:
    """ailine_core の (行数, モジュール数)。★ 再帰して数える（postconditions/ 等の下も）。"""
    files = sorted(CORE.rglob("*.py"))
    return sum(len(p.read_bytes().decode("utf-8").splitlines()) for p in files), len(files)
