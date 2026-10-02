# -*- coding: utf-8 -*-
"""断定の文言の盤 ── 「観測していないことを観測したように言う」文を、書かれた場所から導く（2026-10-02）。

★★ なぜ要るか（盲検の欠陥の形 7）: 断り・警告の文言が、**その場面でしか正しくない**のに広く出ていたり、
  **見たこと**と**推し量り**を分けずに断定していたりした。直しは件ごとに入れたが、次の断定の文は
  次の人（次の自分）が書く ── 書いた時点では、嘘になる場面が見えない。
  ★ だから「断定の語」を含む文が、**畳んだ関数の外に現れたら赤**にする。許す文は理由つきで台帳に持つ。

★ 何を断定の語と見るか（実際に直した形から決めた）:
    closed_two_way   「…のどちらかです」             閉じた二択（第三の可能性を塞ぐ）
    after_change     「そのあと変更されて」           見たのは指紋の不一致だけ・誰が・いつ、は見ていない
    placed_by_human  「人が置いた」                   誰の物かは観測していない
    mixup_or         「…取り違えか」                 原因の断定
    not_supported    「対応していません」             能力の否定（宣言表に当たった時だけ言ってよい）
    not_supported_now「今のところ対応」「未対応です」  同上
    no_such_op       「操作はありません」             一覧に在る操作を無いと言う形
    tool_lacks       「この道具には…ありません」       同上
    not_a_value      「書き込む値ではありません」       名前と同じ字なだけの語を値でないと断定
    absent_or_unchanged 「存在しません/変更されていません」 無いのと変わらないのを畳んでいた
    scopeless_col / scopeless_row  列・行が「無い」と言うのに、探したシートを言わない
★ 語の列挙で正しい ── 判定しているのは「人が何を言ったか」でなく**文を書く側の癖**で、癖は言葉に出る。
  漏れた語は「今までどおり」になるだけ（台帳に載らないだけで、嘘が増えるわけではない）。

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

REGISTER = REPO / "tests" / "claim_wording_register.json"

_JA = re.compile(r"[぀-ヿ一-鿿]")

#: 断定の語（名前 → 正規表現）。
WORDS = {
    "closed_two_way": re.compile(r"のどちらかです"),
    "after_change": re.compile(r"そのあと変更されて"),
    "placed_by_human": re.compile(r"人が置いた"),
    "mixup_or": re.compile(r"取り違えか"),
    "not_supported": re.compile(r"対応していません"),
    "not_supported_now": re.compile(r"今のところ対応|未対応です"),
    "no_such_op": re.compile(r"操作はありません"),
    "tool_lacks": re.compile(r"この道具には[^。]{0,30}ありません"),
    "not_a_value": re.compile(r"書き込む値ではありません"),
    "absent_or_unchanged": re.compile(r"存在しません/変更されていません"),
    "scopeless_col": re.compile(r"列『[^』]*』(?:が|は)(?:この表に)?ありません|『[^』]*』という列がありません"),
    "scopeless_row": re.compile(r"『[^』]*』という行が見つかりません"),
}

#: 範囲を言わない不在の語は、**同じ文に探した範囲が載っていれば**断定でない（探したシートを言う口
#:   `anchor.searched` / `anchor.column_missing` が組んだ文）。★ 式の埋め込みは `{...}` として描かれる。
SCOPE_STATED = re.compile(r"探したシート|\{searched\(|\{where\}|シート『")
_SCOPE_WORDS = ("scopeless_col", "scopeless_row")


def _docstring_ids(tree) -> set:
    ids = set()
    for n in ast.walk(tree):
        if isinstance(n, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            body = n.body
            if (body and isinstance(body[0], ast.Expr)
                    and isinstance(body[0].value, ast.Constant)
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


def sentences(source: str) -> list:
    """source の中の日本語の文を [(関数名, 行, 文)] で返す（docstring を除く・連結は 1 文）。"""
    tree = ast.parse(source)
    doc = _docstring_ids(tree)
    parent = {}
    for p in ast.walk(tree):
        for c in ast.iter_child_nodes(p):
            parent[id(c)] = p

    def func_of(n):
        while id(n) in parent:
            n = parent[id(n)]
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
                return n.name
        return "<module>"

    out = []
    for n in ast.walk(tree):
        if not isinstance(n, (ast.Constant, ast.JoinedStr, ast.BinOp)) or id(n) in doc:
            continue
        p = parent.get(id(n))
        if p is not None and isinstance(p, (ast.BinOp, ast.JoinedStr)) and _render(p) is not None:
            continue                      # 親が既に 1 文として描ける（部分を 2 度数えない）
        text = _render(n)
        if text is None or not _JA.search(text):
            continue
        out.append((func_of(n), n.lineno, text))
    return out


def scan(files=None) -> list:
    """断定の語を含む文を [{file, func, word, line, text}] で返す（製品コード全体）。"""
    found = []
    for f in (files if files is not None else product_files()):
        try:
            rel = str(Path(f).resolve().relative_to(REPO)).replace(chr(92), "/")
        except ValueError:                # repo の外（変異試験の一時ファイル）── 名前だけで持つ
            rel = Path(f).name
        for func, line, text in sentences(Path(f).read_bytes().decode("utf-8")):
            for word, rx in WORDS.items():
                if word in _SCOPE_WORDS and SCOPE_STATED.search(text):
                    continue
                if rx.search(text):
                    found.append({"file": rel, "func": func, "word": word, "line": line,
                                  "text": text})
    return found


def tally(found: list) -> dict:
    """{(file, func, word): 件数}。★ 行番号は鍵に入れない（行が動くだけで赤くしない）。"""
    out: dict = {}
    for x in found:
        k = (x["file"], x["func"], x["word"])
        out[k] = out.get(k, 0) + 1
    return out


def load_register(path: Path = REGISTER) -> dict:
    data = json.loads(Path(path).read_bytes().decode("utf-8"))
    return {(e["file"], e["func"], e["word"]): e for e in data["allowed"]}


def compare(found_tally: dict, register: dict) -> dict:
    """台帳と実装の差。★ 等号で縛る ── 足した文は理由を書くまで赤・台帳に残った文も赤（死んだ免除）。"""
    unlisted = {k: n for k, n in found_tally.items() if k not in register}
    wrong_count = {k: (found_tally[k], register[k]["count"]) for k in found_tally
                   if k in register and found_tally[k] != register[k]["count"]}
    stale = [k for k in register if k not in found_tally]
    return {"unlisted": unlisted, "wrong_count": wrong_count, "stale": stale}


if __name__ == "__main__":
    t = tally(scan())
    for k, n in sorted(t.items()):
        print(n, *k)
    print("total keys", len(t), "sentences", sum(t.values()))
