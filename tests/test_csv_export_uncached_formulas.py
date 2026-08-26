# データの出入口の盲検・致命3（2026-08-26）── キャッシュ値の無い数式セルが
# 分母から消え、空欄で書き出しても「欠落 0」が成立していた。
#
# ★ 実測: 同じブック・同じコマンドで、LO の再計算の有無だけで結果が変わり、
#   **どちらも ✓** だった。金額列が全部数式の見積書は、金額が全部空の CSV が ✓ で出る。
#
# ★ 根: xml_readback.read_grid が「空セル」と「数式だがキャッシュ値が無いセル」を
#   同じ `value is None` で捨てていた。捨てられた側は declared（分母）にも入らない。
#   ★ この週に 4 度出た形 ── **検算の分母が、疑うべき対象と同じ 1 回の読みから来る**。
#
# 契約:
#   ① 読み手が「空」と「読めなかった」を区別して数える
#   ② 出口で名指しして開示する（何行目何列目か）
#   ③ ✓ を名乗らない（書けた分は本当なので × でもない ── △）
#   ④ 数式でもキャッシュ値が在れば従来どおり（誤爆しない）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402
from ailine_core import csv_export, xml_readback  # noqa: E402
from test_golden_transcripts import _isolate, _run_main  # noqa: E402
from test_readback_claim import _inject_formula_cache  # noqa: E402


def _quote_book(tmp_path, name="mitsumori.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "見積"
    ws.append(["品目", "数量", "単価", "金額"])
    ws.append(["A", 2, 500, "=B2*C2"])
    ws.append(["B", 3, 800, "=B3*C3"])
    wb.save(p)
    return p


# --- ① 読み手が区別する -------------------------------------------------------------

def test_reader_separates_empty_from_unreadable(tmp_path):
    p = _quote_book(tmp_path)
    data = xml_readback.read_grid(p, sheet_name="見積")
    assert data["uncached_formulas"] == [(2, 4), (3, 4)], data["uncached_formulas"]
    assert (2, 4) not in data["grid"], "前提: 値としては読めていない"


def test_reader_is_silent_when_the_cache_is_there(tmp_path):
    """④ 誤爆しない: キャッシュ値が在る数式は従来どおり値として読める。"""
    p = _quote_book(tmp_path)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"D2": 1000, "D3": 2400})
    data = xml_readback.read_grid(p, sheet_name="見積")
    assert data["uncached_formulas"] == [], data["uncached_formulas"]
    assert data["grid"][(2, 4)] == 1000


# --- ②③ 出口まで通っている -----------------------------------------------------------

def test_export_csv_discloses_and_refuses_the_checkmark(tmp_path, monkeypatch, capsys):
    """★ 実測の形そのもの: 金額が全部空の CSV に ✓ が出ていた。"""
    _isolate(monkeypatch, tmp_path)
    p = _quote_book(tmp_path)
    out = tmp_path / "o.csv"
    rc, printed = _run_main(["export-csv", str(p), "--sheet", "見積", "--out", str(out)], capsys)
    assert rc == 0, printed
    assert "✓" not in printed, f"金額が空なのに ✓ を名乗った: {printed}"
    assert "△" in printed, printed
    assert "2行目4列目" in printed and "3行目4列目" in printed, \
        f"どのセルが空になったか名指ししていない: {printed}"
    assert "空欄で書き出しました" in printed, \
        f"『検算していません』では実害が伝わらない: {printed}"
    # 実物も確かめる（画面の文字だけで満足しない）
    body = out.read_text(encoding="utf-8-sig")
    assert '"A",2,500,""' in body, body


def test_export_csv_still_claims_when_the_cache_is_there(tmp_path, monkeypatch, capsys):
    """④ 誤爆防止: キャッシュ値が在れば従来どおり ✓。"""
    _isolate(monkeypatch, tmp_path)
    p = _quote_book(tmp_path)
    _inject_formula_cache(p, "xl/worksheets/sheet1.xml", {"D2": 1000, "D3": 2400})
    out = tmp_path / "o.csv"
    rc, printed = _run_main(["export-csv", str(p), "--sheet", "見積", "--out", str(out)], capsys)
    assert rc == 0, printed
    assert "✓" in printed, printed
    assert "空欄で書き出しました" not in printed, printed
