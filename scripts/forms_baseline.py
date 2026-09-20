# -*- coding: utf-8 -*-
"""請求書の読みの基準線 ── **薄い入口**（中身は tests/forms_baseline_core.py）。

★★ なぜ在るか（`docs/PENDING-20260918-請求元と宛先の取り違え.md`）:
  読みの規則に触る前に、305 冊の「宛先/請求元が出たか・区分・値」を凍らせる。
  2026-09-18 はそれが在ったから、225 冊と 16 冊の退行をどちらも commit 前に止められた。

使い方:
    python scripts/forms_baseline.py            # いまの姿と基準線の差を出す（赤で落ちる）
    python scripts/forms_baseline.py --write    # ★ 基準線を取り直す（人が明示的に打つ）

★ `--write` を自動で走らせないこと ── 記録が常に一致して、二度と警告しなくなる（恒真）。

★ なぜ中身が tests/ に在るか: 素の環境の番人は `scripts/` 同士の import を弾く（3 度踏んだ）。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

import forms_baseline_core as core  # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="請求書の読みの基準線")
    ap.add_argument("--write", action="store_true",
                    help="★ 基準線を取り直す（差を読んでから打つこと）")
    a = ap.parse_args(argv)
    now = core.survey()
    if a.write:
        before = core.load()
        if before:
            print(core.render(before, now))
            print("")
        p = core.freeze(now)
        print(f"✓ 基準線を取り直しました: {p}")
        print(f"  {core.counts(now)}")
        return 0
    before = core.load()
    if not before:
        print("？ 基準線がありません ── `--write` で取ってください")
        return 3
    print(core.render(before, now))
    return 1 if core.diff(before, now) else 0


if __name__ == "__main__":
    raise SystemExit(main())
