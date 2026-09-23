"""「4行目」「４行」── 行番号で行を名指しする言い方。

★ 2026-09-23 に src/ailine/__init__.py から**移しただけ**。本体の位置の解決と、
  コード生成（SWAP）が同じ 1 本を使う（書き写さない）。
"""
from __future__ import annotations

import re


_re_row_number_word = re.compile(r"[0-9０-９]+\s*行(?:目)?")
