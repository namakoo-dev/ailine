"""グラフ段の実機（LO・basrun）検体: 一気通貫（「売上の推移を折れ線で」→ 挿入 →
   check_chart_series pass）。DEDUP の実機検体（test_dedup_local.py）と同じ作法:
   sandbox の pytest は basrun/ollama をモックするため構造的に捕まえられない ── 実機で確かめる。"""
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402


@pytest.mark.local
def test_chart_line_runs_on_real_lo_and_series_verifies(tmp_path):
    book = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["月", "売上"])
    for r in [["1月", 100], ["2月", 200], ["3月", 150]]:
        ws.append(r)
    wb.save(book)
    p = subprocess.run(
        [sys.executable, str(REPO / "ailine.py"), "run", str(book),
         "売上の推移を折れ線で見せて", "--copy", "--timeout", "60"],
        capture_output=True, text=True, timeout=300, encoding="utf-8")
    assert p.returncode == 0, f"実機 CHART(line) が失敗:\n{p.stdout[-600:]}"
    out = book.with_name(book.stem + ".out" + book.suffix)
    status, reason = ailine.check_chart_series(out, kind="line", value_col_letter="B")
    assert status == "pass", reason
