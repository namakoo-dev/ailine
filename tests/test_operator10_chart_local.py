"""operator 盲検10度目 ① の実機検体: 合計行のある表へ棒グラフを挿すと、グラフの
   データ範囲（値列の numRef）が合計行を含まず1行縮むことを実 LO で確かめる。
   sandbox の pytest は basrun/ollama をモックするため構造的に捕まえられない ── 実機で確かめる。
"""
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402

_NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}


def _chart_val_ref(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        chart_paths = sorted(n for n in z.namelist()
                              if re.fullmatch(r"xl/charts/chart\d+\.xml", n, re.IGNORECASE))
        assert chart_paths, "chart XML が無い"
        root = ET.fromstring(z.read(chart_paths[0]))
        val_f = root.find(".//c:ser/c:val/c:numRef/c:f", _NS)
        assert val_f is not None and val_f.text, "値列の参照(c:val)が読めない"
        return val_f.text


@pytest.mark.local
def test_chart_range_excludes_total_row_on_real_lo(tmp_path):
    book = tmp_path / "summary.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "集計"
    for r in [["部門", "金額"], ["営業", 800], ["開発", 700], ["総務", 250], ["合計", 1750]]:
        ws.append(r)
    wb.save(book)
    p = subprocess.run(
        [sys.executable, str(REPO / "ailine.py"), "run", str(book),
         "集計シートで部門ごとの棒グラフを作って", "--copy", "--timeout", "60"],
        capture_output=True, text=True, timeout=300, encoding="utf-8")
    assert p.returncode == 0, f"実機 CHART(bar) が失敗:\n{p.stdout[-600:]}"
    out = book.with_name(book.stem + ".out" + book.suffix)
    ref = _chart_val_ref(out)
    # 期待: 集計!$B$2:$B$4（合計行=5行目を含まない・第4の柱として混入しない）
    assert re.search(r"\$B\$2:\$B\$4$", ref), f"合計行がグラフ範囲に混入している: {ref}"
    status, reason = ailine.check_chart_series(out, kind="bar", value_col_letter="B",
                                                category_col_letter="A")
    assert status == "pass", reason
