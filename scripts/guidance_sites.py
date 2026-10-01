# -*- coding: utf-8 -*-
"""案内（次に打つもの・書くもの）を製品のソースから数える（薄い入口）。中身は tests/guidance_sites_core.py。

★ なぜ中身が tests/ に在るか: 素の環境の番人（scripts/_ci_parity_blocker.py）は
  requirements-dev.txt に無い import を全部止めるので、scripts/ 同士の import が弾かれる。

    python scripts/guidance_sites.py          # 数（hints / guidance / 台帳に無いもの / 歩いたか）
    python scripts/guidance_sites.py --list   # 台帳に無い案内を 1 行ずつ
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tests"))

from guidance_sites_core import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
