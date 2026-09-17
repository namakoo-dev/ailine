# -*- coding: utf-8 -*-
"""配線盤を表示する ── 中身は tests/wiring_board_core.py（2026-09-16）。

★ ここは**薄い入口**だけ。中身を tests/ に置くのは、素の環境の番人が
  「宣言外の import」を全部止めるため（scripts/ から import すると弾かれる）。

    python scripts/wiring_board.py           # 人が読む形
    python scripts/wiring_board.py --json    # 機械が読む形（可視化が使う）
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

from wiring_board_core import mismatches, survey  # noqa: E402

LABEL = {"derived": "導出あり", "watched": "番人あり", "explained": "導けない",
         "unstudied": "★ 未調査", "bare": "★ 無防備"}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--json", action="store_true", help="機械が読む形で出す")
    a = ap.parse_args(argv)
    data = survey()
    if a.json:
        print(json.dumps(data, ensure_ascii=False, indent=1))
        return 0

    cols, ops = data["columns"], data["ops"]
    by = {k: [c for c in cols if c["stance"] == k] for k in LABEL}
    print(f"配線盤 — {data['commit']}")
    print(f"op {len(ops)} × 判断 {len(cols)}")
    print("  " + " / ".join(f"{LABEL[k]} {len(by[k])}" for k in LABEL))
    risky = by["unstudied"] + by["bare"]
    print(f"  ★ 調べていない列 {len(risky)} ── ここが次の盲点の育つ場所"
          f"（{sum(len(c['declared']) for c in risky)} マス）\n")
    for c in cols:
        print(f"  [{LABEL[c['stance']]:8}] {c['name']:32} {len(c['declared']):2}/{len(ops)} op"
              + (f"  番人 {len(c['watchers'])} 本" if c["stance"] == "watched" else ""))
    ms = mismatches(data)
    print(f"\n宣言と実測の食い違い: {len(ms)} 列")
    for m in ms:
        print(f"  ★ {m['column']}")
        if m["declared_only"]:
            print(f"      宣言だけ: {' '.join(m['declared_only'])}")
        if m["derived_only"]:
            print(f"      実測だけ（候補・確定ではない）: {' '.join(m['derived_only'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
