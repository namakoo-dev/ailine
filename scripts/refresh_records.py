# -*- coding: utf-8 -*-
"""導出でしかない記録を作り直す（2026-09-16）。

★ なぜ在るか（Namakoo「これって ailine をいじる度に更新しないといけないよね」）:
  2026-09-16 の 1 日で、記録ファイルを触った commit は
  行数 5 回 / README 10 回 / 依存の図 3 回。数え直すだけの作業が push を 2 回止めた。

★★ ただし記録は 2 種類あり、**扱いを逆にすると害になる**:

  ① 導出でしかないもの  ── コードを見れば答えが一意に決まる。人が数える意味はゼロ。
     `ailine.py` の行数 / 試験の本数 / 依存の図。**ここだけをこの道具が作り直す。**

  ② 測定を記録したもの  ── 実機を回して初めて出る（効果の行列 245 件・翻訳精度）。
     ★ **自動更新してはいけない。** 自動で揃えたら、記録は常に実測と一致し、
       二度と警告しなくなる ── 恒真。2026-09-16 に、依存の図が
       「生成器と番人が同じ盲点を共有していたせいで 13 日間 48% 間違ったまま緑」
       だったのを見つけた。②を自動化すると、同じ壊れ方を自分から作ることになる。

  だからこの道具は②に**触らない**。触ろうとしたら、そう言って止まる。

    python scripts/refresh_records.py            # 何が古いかを見るだけ
    python scripts/refresh_records.py --write    # ①だけ作り直す
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
MAIN = REPO / "src" / "ailine" / "__init__.py"
BUDGET = REPO / "tests" / "ailine_py_line_budget.txt"

#: ★ この道具が**触らない**印（②測定の記録）。名前で守る ── 増えたらここに足す。
MEASURED_MARKS = ("MATRIX", "MATRIX_CASES", "MATRIX_REFUSED", "BATTERY", "ACCURACY")


def _count_tests() -> tuple:
    """pytest が実際に集める本数（全体・実機）。★ 数えるのは pytest であって人ではない。"""
    def collect(*extra):
        r = subprocess.run([sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider",
                            "--no-header", "--collect-only", *extra],
                           cwd=str(REPO), capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        m = re.search(r"(\d+)(?:/(\d+))? tests collected", r.stdout)
        if not m:
            raise SystemExit("pytest の収集結果を読めない:\n" + r.stdout[-600:])
        return int(m.group(1))
    return collect(), collect("-m", "local")


def _replace_mark(mark: str, value: str, write: bool) -> list:
    """印の中身を差し替える。★ bytes で読み書きする（write_text は CRLF を壊す）。"""
    changed = []
    pat = re.compile(("<!-- " + mark + " -->(.*?)<!-- /" + mark + " -->").encode("utf-8"), re.S)
    want = ("<!-- " + mark + " -->" + value + "<!-- /" + mark + " -->").encode("utf-8")
    for p in sorted(REPO.rglob("*.md")):
        if any(x in str(p) for x in (".git", "node_modules")):
            continue
        b = p.read_bytes()
        if not pat.search(b):
            continue
        nb = pat.sub(want, b)
        if nb == b:
            continue
        changed.append(str(p.relative_to(REPO)))
        if write:
            crlf = b.count(b"\r\n")
            p.write_bytes(nb)
            assert p.read_bytes().count(b"\r\n") == crlf, f"改行を壊した: {p}"
    return changed


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--write", action="store_true", help="①だけ作り直す")
    ap.add_argument("--parts", default="lines,tests,graph",
                    help="見るものを絞る（lines / tests / graph をカンマ区切り）。"
                         "★ commit の瞬間に走らせる時は、変わったファイルに合わせて絞る "
                         "── 全部だと 10 秒、絞れば 0〜7 秒")
    a = ap.parse_args(argv)
    parts = {x.strip() for x in a.parts.split(",") if x.strip()}
    unknown = parts - {"lines", "tests", "graph"}
    if unknown:
        raise SystemExit(f"知らない部位: {sorted(unknown)}（lines / tests / graph）")

    todo = []

    if "lines" in parts:
        lines = len(MAIN.read_bytes().decode("utf-8").splitlines())
        if BUDGET.read_bytes().decode("utf-8").strip() != str(lines):
            todo.append(f"ailine.py の行数 → {lines}")
            if a.write:
                BUDGET.write_bytes((str(lines) + "\n").encode("utf-8"))
        for f in _replace_mark("MAIN_FILE_LINES", str(lines), a.write):
            todo.append(f"  印 MAIN_FILE_LINES: {f}")

    if "tests" in parts:
        total, local = _count_tests()
        for mark, val in (("TOTAL_TESTS", total), ("LOCAL_TESTS", local)):
            for f in _replace_mark(mark, str(val), a.write):
                todo.append(f"  印 {mark} → {val}: {f}")

    r = None
    if "graph" in parts:
        r = subprocess.run([sys.executable, str(REPO / "scripts" / "deps_graph.py"),
                            *(["--write"] if a.write else [])],
                           cwd=str(REPO), capture_output=True, text=True, encoding="utf-8")
        if r.returncode != 0:
            raise SystemExit("依存の図の生成に失敗:\n" + r.stderr[-600:])
    doc = REPO / "docs" / "依存関係.md"
    if r is not None and not a.write and doc.exists():
        have = doc.read_bytes().decode("utf-8").replace("\r\n", "\n").strip()
        if have != r.stdout.replace("\r\n", "\n").strip():
            todo.append("依存の図 → scripts/deps_graph.py --write")

    if a.write:
        print("作り直した:" if todo else "作り直すものは無かった")
    else:
        print("古い記録:" if todo else "記録は実体と揃っている")
    for t in todo:
        print("  " + t)

    print("\n★ この道具が**触らないもの**（実機を回して初めて出る数字）:")
    print("   " + " / ".join(MEASURED_MARKS))
    print("   自動で揃えると記録は常に一致し、二度と警告しなくなる（恒真）。")
    print("   測り直したい時は bench/basic_ops_matrix.py を回し、"
          "出た三つ組を人が tests/battery_recorded.json へ書く。")
    return 1 if (todo and not a.write) else 0


if __name__ == "__main__":
    raise SystemExit(main())
