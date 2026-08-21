"""demo ファイルの鮮度番人（operator 盲検 7 度目・摩擦①の再発防止・2026-08-21）。

★ なぜ在るか: README の一発目の例文を demo/sample.xlsx に実行した結果の列（売上-原価）が
commit に焼き込まれたまま残り、新規購入者が最初にコピペするコマンドが exit 7
（「同じ依頼を 2 回実行した可能性」の関所）で止まった ── 初見の最初の 1 分で萎えさせる。
検証中の実行結果を demo に commit しない、を機械で見張る。"""
import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))


def test_demo_sample_has_no_readme_example_output_baked_in():
    ws = openpyxl.load_workbook(REPO / "demo" / "sample.xlsx").active
    headers = [c.value for c in ws[1]]
    assert "売上-原価" not in headers, (
        f"README 例文の実行結果が demo に焼き込まれている（初見の一発目が exit 7 になる）: {headers}"
    )
    assert headers == ["商品", "金額", "在庫", "売上", "原価"], headers
