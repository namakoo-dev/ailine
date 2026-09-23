# -*- coding: utf-8 -*-
"""単一ファイルを割る作業の**分母を出す**（番人ではなく測定器）。

★ なぜ在るか: `src/ailine/__init__.py` は 2 万行近くあり、割る作業が続いている。
  「あと何が残っているか」を毎回同じ手順で数える。
★★ 2026-09-23: 中身は tests/split_progress_core.py に移した（番人が scripts/ を import すると
  素の環境で弾かれるため・つめ車 tests/test_pure_logic_only_shrinks.py と**同じ測定器**を使う）。
  ここは入口だけ。数え方の粗さと、第 1 版の盲点（モジュール定数を見ていなかった）は芯の冒頭に書いた。

    python scripts/split_progress.py            # 進捗の数
    python scripts/split_progress.py --pure     # 純と数えた関数（ailine_core へ持ち出せる候補）
    python scripts/split_progress.py --clusters # 名前のまとまりの候補（分け方の材料）
"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))
import split_progress_core as spc  # noqa: E402


def main(argv):
    result = spc.survey()
    funcs = result["funcs"]
    total = len(result["src"].splitlines())
    core_lines, core_files = spc.core_lines()
    pure_lines, pure_n = spc.pure_logic(result)
    held = [r for r in funcs.values() if r["held"] and not r["io"]]

    print(f"本体            {total:>6} 行   関数 {len(funcs)} 個")
    print(f"ailine_core     {core_lines:>6} 行   {core_files} モジュール"
          f"（既に外へ出ている割合 {core_lines / (core_lines + total) * 100:.0f}%）")
    print(f"純ロジック       {pure_lines:>6} 行   {pure_n} 関数"
          "（I/O なし・差し替え名を読まない・残る関数を呼ばない）")
    print(f"  ★ I/O は無いが差し替え名を読むので残る: {len(held)} 関数"
          f"（例: {', '.join(r['name'] for r in held[:4])}）")

    pure = [r for r in funcs.values() if r["pure"]]
    if "--pure" in argv:
        print("\n=== 純と数えた関数（行数の多い順）===")
        for r in sorted(pure, key=lambda r: -r["lines"]):
            extra = []
            if r["consts"]:
                extra.append("定数 " + ",".join(r["consts"][:4]))
            if r["imports"]:
                extra.append("import " + ",".join(r["imports"][:4]))
            print(f"  {r['lines']:>4} {r['name']}  {' / '.join(extra)}")
    elif "--clusters" in argv:
        print("\n=== 名前のまとまり（分け方の材料・人が決めるための手がかり）===")
        words = Counter()
        for r in pure:
            for w in re.findall(r"[a-z]+", r["name"]):
                if len(w) >= 4:
                    words[w] += 1
        for w, n in words.most_common(14):
            members = [r for r in pure if re.search(rf"\b{w}\b|_{w}_|^{w}_|_{w}$", r["name"])]
            if len(members) < 2:
                continue
            ln = sum(m["lines"] for m in members)
            print(f"  {w:<12} {len(members):>2} 関数 {ln:>4} 行  "
                  f"{', '.join(m['name'] for m in members[:4])}"
                  f"{' …' if len(members) > 4 else ''}")
    else:
        print("\n（一覧は --pure、分け方の材料は --clusters）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
