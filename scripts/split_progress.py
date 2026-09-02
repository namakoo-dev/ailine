# -*- coding: utf-8 -*-
"""単一ファイルを割る作業の**分母を出す**（番人ではなく測定器）。

★ なぜ在るか: `src/ailine/__init__.py` は 17,000 行を超え、README も
「単一ファイルを割るべきだが、挙動を変えずに割ったことを確かめる番人を用意できて
いない」と書いている。番人の 1 本目は `tests/test_public_surface_is_frozen.py`
（名前と署名の凍結）で、これはその**手前**にある道具 ── 「あと何が残っているか」
を毎回同じ手順で数える。

★ 数を凍結しない理由: 分割の最中はこの数が激しく動く。閾値を置くと更新が仕事になり、
番人が形骸化する（この repo が何度も踏んだ「在っても鳴らない」の形）。
**発火条件つきで保留する** ── 分割が終わって数が落ち着いたら、
「本体の純ロジックが増えたら赤」を tests/ 側に置く。それまでは数えるだけ。

★ 判定の粗さ（先に書いておく）:
  ・「I/O を触らない」は呼び出し名の一致で見ている。動的な呼び出しは捕まらない
  ・モジュール定数への依存は見ていない（外へ出す時は定数も一緒に連れて行く必要がある）
  ・**この一覧は「出せる」ではなく「出せる候補」** ── まとまりを決めるのは人。
    関数単位で外へ出すと ailine_core が雑多になり、モジュールの仕様説明が書けなくなる。

    python scripts/split_progress.py            # 進捗の数
    python scripts/split_progress.py --clusters # 名前のまとまりの候補（分け方の材料）
"""
from __future__ import annotations

import ast
import re
import sys
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
MAIN = REPO / "src" / "ailine" / "__init__.py"
CORE = REPO / "src" / "ailine_core"

# 触っていたら「純ロジックではない」と見なす名前（呼び出し名・属性名で照合）
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


def survey():
    src = MAIN.read_text(encoding="utf-8")
    tree = ast.parse(src)
    top = {n.name: n for n in tree.body
           if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
    rows = []
    for name, fn in top.items():
        called = _called_names(fn)
        if called & IMPURE:
            continue
        deps = sorted(called & set(top))
        rows.append({"name": name, "lines": fn.end_lineno - fn.lineno + 1,
                     "deps": deps})
    return src, top, rows


def main(argv):
    src, top, pure = survey()
    total = len(src.splitlines())
    core_files = sorted(CORE.glob("*.py"))
    core_lines = sum(len(p.read_text(encoding="utf-8").splitlines())
                     for p in core_files)
    free = [r for r in pure if not r["deps"]]
    near = [r for r in pure if 1 <= len(r["deps"]) <= 2]

    print(f"本体            {total:>6} 行   関数 {len(top)} 個")
    print(f"ailine_core     {core_lines:>6} 行   {len(core_files)} モジュール"
          f"（既に外へ出ている割合 {core_lines / (core_lines + total) * 100:.0f}%）")
    print(f"純ロジックの候補  依存ゼロ {sum(r['lines'] for r in free):>5} 行 / "
          f"{len(free)} 関数   ・ 依存 1〜2 本 "
          f"{sum(r['lines'] for r in near)} 行 / {len(near)} 関数")

    if "--clusters" in argv:
        print("\n=== 名前のまとまり（分け方の材料・人が決めるための手がかり）===")
        words = Counter()
        for r in pure:
            for w in re.findall(r"[a-z]+", r["name"]):
                if len(w) >= 4:
                    words[w] += 1
        for w, n in words.most_common(14):
            members = [r for r in pure if re.search(rf"\b{w}\b|_{w}_|^{w}_|_{w}$",
                                                     r["name"])]
            if len(members) < 2:
                continue
            ln = sum(m["lines"] for m in members)
            print(f"  {w:<12} {len(members):>2} 関数 {ln:>4} 行  "
                  f"{', '.join(m['name'] for m in members[:4])}"
                  f"{' …' if len(members) > 4 else ''}")
    else:
        print("\n（分け方の材料は --clusters）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
