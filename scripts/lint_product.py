# -*- coding: utf-8 -*-
"""製品コード（src/・scripts/）のリンタ番人。

★ なぜ在るか（2026-08-24 の盲検査定）: 「リンタが CI にも pre-push にも入っていない」と
指摘され、実際に測ったら製品コードに 12 件（死んだ変数・未使用 import・placeholder の
無い f-string）が溜まっていた。どれも動作は壊さないが、**11,653 行の 1 ファイルが
道具で手入れされていない**という読み方をされる ── そしてそれは事実だった。

★ pyflakes は `# noqa` を読まない。意図して残す import（再輸出・在否確認）は在るので、
   ここで noqa を汲む ── ただし**黙って捨てない**: 何件を宣言で見逃したかを必ず印字する
   （silent cap を作らない）。

★ tests/ は対象外。50 件あり、今夜まとめて掃くと本物の変更が埋もれる。
   件数をここに書いて可視化しておく（見えている借金にする）。
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGETS = ["src/", "scripts/"]


def main() -> int:
    r = subprocess.run([sys.executable, "-m", "pyflakes", *TARGETS],
                        cwd=str(ROOT), capture_output=True, text=True)
    real, allowed = [], []
    for line in (r.stdout or "").splitlines():
        m = re.match(r"(.+?):(\d+):\d+: (.+)", line)
        if not m:
            continue
        path, lineno, msg = m.group(1).replace("\\", "/"), int(m.group(2)), m.group(3)
        try:
            src_line = (ROOT / path).read_bytes().decode("utf-8").split("\n")[lineno - 1]
        except Exception:
            src_line = ""
        (allowed if "# noqa" in src_line else real).append((path, lineno, msg))

    if allowed:
        print(f"（宣言で見逃した残置: {len(allowed)} 件 ── # noqa 付き）")
        for path, lineno, msg in allowed:
            print(f"    {path}:{lineno} {msg}")
    if real:
        print(f"✗ 製品コードに {len(real)} 件:", file=sys.stderr)
        for path, lineno, msg in real:
            print(f"    {path}:{lineno} {msg}", file=sys.stderr)
        print("  意図して残すなら、その行に # noqa と**理由**を書いてください。",
              file=sys.stderr)
        return 1
    print("✓ 製品コードのリンタ: 指摘なし")
    # ★ 見えている借金を毎回言う（tests/ は今のところ対象外）。
    t = subprocess.run([sys.executable, "-m", "pyflakes", "tests/"],
                        cwd=str(ROOT), capture_output=True, text=True)
    n = len([l for l in (t.stdout or "").splitlines() if re.match(r".+?:\d+:\d+: ", l)])
    if n:
        print(f"（未処置: tests/ に {n} 件 ── 対象外にしているだけで、消えてはいません）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
