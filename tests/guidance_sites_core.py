"""「次に打つもの・書くもの」を画面に出す所（案内）を、製品のソースから全部数える（2026-10-01）。

★★ なぜ在るか（盲検で同じ形が 6 回 ──「嘘の案内」）:
  道具が「こう打て」「こう書け」と言い、その通りにすると落ちる／抜けられない／手段が無い。
    ・「依頼文に列名を含めて（例:『取引先コードをキーに』）」── 依頼者は**既にそう書いていた**（3・6・7 体目）
    ・「`--header-row 4` のように指定して再実行」── その通りに打っても同じ文で止まる（`--sheet` も要った）
    ・「`--header-row 3` のように」── その冊の見出しは 4 行目（3 は決め打ちの例）
  ★ 導線の台帳（typable_hints_register.json の hints）は既に在ったが、分母が
    「バッククォートの中の `ailine …` / `--…`」だけだった ──「例:『…』」も「〜と書いてください」も
    **分母の外**に在った。1 体目の事故（3 体目）も 7 体目の再来も、台帳に 1 行も無かった。

★ ここが数えるもの（分母は機械が出す・手書きの名簿にしない）:
    製品の画面に出うる文字列（tests/_product_source.product_strings ── docstring を除く定数と
    f-string の固定部分）のうち、日本語の文の中に**次の一手を示す印**を含むもの。印は MARKS。
  ★ hints（バッククォートの `ailine …`/`--…`）と重なるものは hints 側の台帳が持つ ── ここは**その外**だけ。
  ★ f-string・定数の連結で印が**断片の境目で割れた**案内も拾う（split_sites）── 断片ごとに見るだけだと
    「…のように」+ 変数 +「指定して」の形が分母から落ちる。

★ 判定（歩いたか）は持たない ── 台帳（guidance 欄）と歩き手（walk_refusals_core）の仕事。
"""
from __future__ import annotations

import ast
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _product_source import product_files, product_strings  # noqa: E402

REGISTER = Path(__file__).resolve().parent / "typable_hints_register.json"

#: ★ hints 側の分母（test_every_typable_hint_is_accounted_for.TYPABLE と同じ式 ── 重なりを外すのに使う）
TYPABLE = re.compile(r"`(ailine [a-z-]+|--[a-z-]+)")

#: 次の一手を示す印。★ 足す時は「その印が付いた文は、人に何かを打たせる／書かせる」ことを確かめてから。
MARKS = {
    # 日本語の文の中の --フラグ（バッククォート無し）
    "bare_flag": re.compile(r"(?<![`\w-])--[a-z][a-z-]+"),
    # 日本語の文の中の `ailine <サブコマンド>`（バッククォート無し）
    "bare_sub": re.compile(r"(?<![`\w])ailine [a-z][a-z-]+"),
    # 例文
    "example": re.compile(r"例[:：]|例えば"),
    # 書け・打て・やり直せ
    "imperative": re.compile(
        r"と書いて|書き足して|のように指定|付けて(?:ください|もう一度|再実行|出し直|実行)"
        r"|再実行|もう一度実行|言い直して|と頼んで"),
}

#: 日本語の文か（★ 印だけの断片 ── argv の '--sheet' や 'ailine accounts' という名札 ── は人に向けた文でない）
_JAPANESE = re.compile(r"[぀-ヿ一-鿿]")
#: いま走っている入口の名札（「■ ailine run（2冊の照合）  A=…」）── 打てと言っていない
_HEADLINE = re.compile(r"^\s*■ ailine ")


def marks_of(text: str) -> list:
    """その文字列が持つ印（MARKS の名前）。人に向けた文でなければ空。"""
    if not _JAPANESE.search(text) or _HEADLINE.search(text):
        return []
    return [k for k, p in MARKS.items() if p.search(text)]


def _all_hits() -> list:
    """(ファイル名, 行, 文字列, 印) ── 印を持つ文字列を全部（hints との重なりも含む）。"""
    out = []
    for f, ln, v in product_strings():
        m = marks_of(v)
        if m or TYPABLE.search(v):
            out.append((Path(f).name, ln, v, m))
    return out


def sites() -> list:
    """★ 本体 ── hints の外に在る案内を (ファイル, 断片, 何番目か, 印) で全部。

    鍵は hints と同じ『ファイル＋断片（完全一致）＋何番目か』（行番号は上に 1 行足すだけで全部ずれる）。
    何番目かは**この分母の中で**数える。
    """
    seen: Counter = Counter()
    out = []
    for name, _ln, v, m in _all_hits():
        if TYPABLE.search(v) or not m:
            continue
        k = (name, v.strip())
        seen[k] += 1
        out.append((name, v.strip(), seen[k], tuple(m)))
    # ★ 断片に割ると印が消える案内は、丸ごとの型（穴は「{…}」）で載せる ── 分母から落とさない
    for name, _ln, t in split_sites():
        k = (name, t.strip())
        seen[k] += 1
        out.append((name, t.strip(), seen[k], tuple(marks_of(t))))
    return out


def _template(node) -> str | None:
    """文字列の式を 1 本の型にする（f-string の穴・連結の変数は「{…}」）。文字列の式でなければ None。"""
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):
        return "".join(p.value if isinstance(p, ast.Constant) else "{…}" for p in node.values)
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
        a, b = _template(node.left), _template(node.right)
        if a is None and b is None:
            return None
        return (a if a is not None else "{…}") + (b if b is not None else "{…}")
    return None


def split_sites() -> list:
    """★ 断片に割ると印が消える案内 ── f-string・定数の連結の**丸ごと**には印があるのに、
    どの断片にも印が無いもの (ファイル, 行, 型)。ここが空でない間は、分母に穴が在る。"""
    out = []
    for f in product_files():
        tree = ast.parse(f.read_bytes().decode("utf-8"))
        inner = set()
        for node in ast.walk(tree):
            if isinstance(node, (ast.JoinedStr, ast.BinOp)):
                for ch in ast.walk(node):
                    if ch is not node:
                        inner.add(id(ch))
        for node in ast.walk(tree):
            if id(node) in inner or not isinstance(node, (ast.JoinedStr, ast.BinOp)):
                continue
            t = _template(node)
            if t is None or not (marks_of(t) or TYPABLE.search(t)):
                continue
            pieces = [ch.value for ch in ast.walk(node)
                      if isinstance(ch, ast.Constant) and isinstance(ch.value, str)]
            if not any(marks_of(p) or TYPABLE.search(p) for p in pieces):
                out.append((f.name, node.lineno, t))
    return out


def register() -> dict:
    return json.loads(REGISTER.read_bytes().decode("utf-8"))


def census() -> dict:
    """数（直す前→後を報告に出すための形）。★ 分母は機械・台帳は突き合わせるだけ。"""
    reg = register()
    hints = reg.get("hints", [])
    guide = reg.get("guidance", [])
    now = sites()
    keys = {(e["file"], e["text"], e.get("nth", 1)) for e in guide}
    unlisted = [s for s in now if (s[0], s[1], s[2]) not in keys]
    by_mark = Counter(m for s in now for m in s[3])
    walked = Counter(e.get("walked") for e in guide)
    return {
        "hints（バッククォートの打てるもの）": len(hints),
        "hints の walked": dict(Counter(e.get("walked") for e in hints)),
        "guidance（それ以外の案内）": len(now),
        "guidance の印ごと": dict(by_mark),
        "guidance のうち台帳に無い": len(unlisted),
        "guidance の walked": dict(walked),
        "断片に割ると印が消える案内": len(split_sites()),
    }


def main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(description="案内（次に打つもの・書くもの）を製品のソースから数える")
    ap.add_argument("--list", action="store_true", help="台帳に無い案内を 1 行ずつ出す")
    a = ap.parse_args(argv)
    print(json.dumps(census(), ensure_ascii=False, indent=1))
    if a.list:
        keys = {(e["file"], e["text"], e.get("nth", 1)) for e in register().get("guidance", [])}
        for s in sites():
            if (s[0], s[1], s[2]) not in keys:
                print(f"{s[0]}  {'/'.join(s[3])}  {s[1][:90]}")
        for f, ln, t in split_sites():
            print(f"[割れ] {f}:{ln}  {t[:90]}")
    return 0
