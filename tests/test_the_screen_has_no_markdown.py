"""画面に出る文字列に markdown の強調（**）を書かない（2026-09-22）。

★★ 事故: 盲検の買い手役が 2 度（09-15 ⑧ / 09-20 5 体目）「画面に `**` が出ている」と挙げた。
  ★ 端末は markdown を描かない。**GUI も描かない** ── gui/index.html は textContent と
    <pre> で出しているので、`**` は記号としてそのまま見える（2026-09-22 に実測）。
    つまり強調として効いている場所は **1 つも無かった**。

★★ 30 分で 2 回直したのに、機械で数えたら **31 件**在った（私の手勘定は 3 件だった）。
  手で潰すのをやめて、分母を機械から取る ── これがこの番人。

★ 分母の作り方: 製品コードの**画面に出うる文字列**（docstring とコメントは除く）を
  AST で全部取り、台帳 tests/screen_emphasis_register.json に理由つきで載っているものだけ
  除く。★ 台帳には「画面に出ないもの」しか載せられない
  （『画面に出るが ** を残したい』という枠は用意していない）。

★ 模型に渡るプロンプトを除くのは**手加減ではない**: プロンプトは 1 バイト変えると
  op 一致が動く（OPS_DOC に 16 行足したら 98.1% → 94.2% の実測）。
  `**` が模型への強調として効いているかは測っていないので、測らずに消さない。
"""
import ast
import json
from pathlib import Path

from _product_source import product_files, product_strings

REGISTER = Path(__file__).resolve().parent / "screen_emphasis_register.json"


def _exempt_names() -> dict:
    d = json.loads(REGISTER.read_text(encoding="utf-8"))
    return {e["name"]: e for e in d["exempt"]}


def _owner_name(node, parent) -> str:
    """その文字列が属する代入の名前（module 直下の定数名）。無ければ空。"""
    while node in parent:
        node = parent[node]
        if isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name):
            return node.targets[0].id
    return ""


def _strings_with_stars():
    """`**` を持つ、画面に出うる文字列を (ファイル, 行, 定数名, 中身) で返す。"""
    out = []
    for f in product_files():
        tree = ast.parse(f.read_text(encoding="utf-8"))
        docs, parent = set(), {}
        for n in ast.walk(tree):
            if isinstance(n, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                b = n.body[0] if n.body else None
                if (isinstance(b, ast.Expr) and isinstance(b.value, ast.Constant)
                        and isinstance(b.value.value, str)):
                    docs.add(id(b.value))
            for c in ast.iter_child_nodes(n):
                parent[c] = n
        for n in ast.walk(tree):
            if (isinstance(n, ast.Constant) and isinstance(n.value, str)
                    and "**" in n.value and id(n) not in docs):
                out.append((f, n.lineno, _owner_name(n, parent), n.value))
    return out


def test_the_screen_does_not_show_markdown_bold():
    """★★ 本体 ── 台帳に載っていない文字列は `**` を持たないこと。"""
    exempt = _exempt_names()
    strays = [(f, ln, name, v) for f, ln, name, v in _strings_with_stars()
              if name not in exempt]
    assert not strays, (
        "画面に出る文字列に `**` があります（端末も GUI も markdown を描きません）:"
        + chr(10)
        + chr(10).join(f"  {Path(f).name}:{ln}  {v[:70]!r}" for f, ln, _, v in strays)
        + chr(10) + "  削るか、画面に出ない理由を tests/screen_emphasis_register.json に"
        " 1 行で書いてください。")


def test_every_exemption_is_still_real():
    """★ 台帳に残っているのに `**` がもう無いなら、その行は消す（古い不安を配らない）。

    ★ この向きが無いと、台帳は増える一方になり「載っているから安心」が嘘になる。
    """
    names_with_stars = {name for _, _, name, _ in _strings_with_stars() if name}
    stale = [n for n in _exempt_names() if n not in names_with_stars]
    assert not stale, (
        f"台帳に在るのに `**` を持たなくなったもの: {stale} ── 台帳から消してください")


def test_every_exemption_says_why_it_is_not_on_screen():
    """★ 免除には**理由**が要る（数を通すために足させない）。"""
    for name, e in _exempt_names().items():
        assert e.get("why"), f"{name} に why がありません"
        assert e.get("kind") in ("模型に渡る", "描画されない記録"), (
            f"{name} の kind が台帳の 2 種類のどちらでもありません: {e.get('kind')!r}")
