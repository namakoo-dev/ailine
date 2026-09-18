# -*- coding: utf-8 -*-
"""盲検 1 回ぶんを凍らせて・流し直して・採る ── 中身は tests/blind_run_core.py（2026-09-18）。

★ ここは**薄い入口**だけ。中身を tests/ に置くのは、素の環境の番人が
  「宣言外の import」を全部止めるため（scripts/ から import すると弾かれる）。

    python scripts/blind_run.py freeze <回> <作業フォルダ>
    python scripts/blind_run.py replay <回>
    python scripts/blind_run.py floor  <回>     # 揺れの床
    python scripts/blind_run.py sheet  <回>     # 機械で決まる 5 条だけの採点表
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

from blind_run_core import cmd_freeze, cmd_floor, cmd_replay, cmd_sheet  # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("freeze"); f.add_argument("n"); f.add_argument("workdir")
    for name in ("replay", "floor", "sheet"):
        sub.add_parser(name).add_argument("n")
    a = ap.parse_args(argv)
    if a.cmd == "freeze":
        return cmd_freeze(a.n, a.workdir)
    return {"replay": cmd_replay, "floor": cmd_floor, "sheet": cmd_sheet}[a.cmd](a.n)


if __name__ == "__main__":
    raise SystemExit(main())
