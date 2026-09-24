"""色の名前 → RGB（16 進）の表 ── 引数の検査・Basic の生成・事後条件の 3 つが同じ表を引く。

★ 2026-09-24 に postconditions/_shared.py から**移しただけ**（本文は 1 文字も変えていない）。
  生成（codegen）が事後条件の内輪のモジュールを import していた配置を解くため、中立な冊にした。
  _shared は同じ実体を再輸出する。
"""
from __future__ import annotations


COLOR_MAP = {
    "red": "FF0000", "green": "00B050", "blue": "0000FF", "yellow": "FFFF00",
    "orange": "FFA500", "purple": "800080", "pink": "FFC0CB", "black": "000000",
    "white": "FFFFFF", "gray": "808080", "grey": "808080",
    "lightblue": "ADD8E6", "lightgreen": "90EE90", "lightyellow": "FFFFE0",
    "lightred": "FFCCCC", "lightgray": "D3D3D3", "lightgrey": "D3D3D3",
}
