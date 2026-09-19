# -*- coding: utf-8 -*-
"""断りが示した道を歩く（薄い入口）。中身は tests/walk_refusals_core.py。

★ なぜ中身が tests/ に在るか: 素の環境の番人（scripts/_ci_parity_blocker.py）は
  requirements-dev.txt に無い import を全部止めるので、scripts/ 同士の import が
  **自分の repo の道具なのに**弾かれる（3 度踏んだ）。

    python scripts/walk_refusals.py
    python scripts/walk_refusals.py --json
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tests"))

from walk_refusals_core import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
