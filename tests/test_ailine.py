"""ailine の純粋ロジックの単体テスト（ollama / LibreOffice を要さない部分）。
   生成・適用の統合は実機（basrun_spike）で検証済み。ここは回帰用の土台。
"""
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine


# --- コード抽出 / 署名 -----------------------------------------------------

def test_extract_bas_strips_markdown_fence():
    raw = "```basic\nOption Explicit\nSub Run(oDoc As Object)\nEnd Sub\n```"
    assert ailine.extract_bas(raw).startswith("Option Explicit")
    assert "```" not in ailine.extract_bas(raw)

def test_extract_bas_passthrough_without_fence():
    raw = "Sub Run(oDoc As Object)\nEnd Sub"
    assert ailine.extract_bas(raw) == raw

@pytest.mark.parametrize("code,ok", [
    ("Sub Run(oDoc As Object)\nEnd Sub", True),
    ("sub run( oDoc as object )", True),          # 大文字小文字・空白ゆらぎ
    ("Sub Run()\nEnd Sub", False),                # 引数なし
    ("Sub Other(oDoc As Object)", False),         # 別名
    ("' コメントだけ", False),
])
def test_valid_signature(code, ok):
    assert ailine.valid_signature(code) is ok


# --- 参照ライブラリ --------------------------------------------------------

def test_load_refs_bundles_examples():
    text = ailine.load_refs(ailine.DEFAULT_REFS)
    assert "Sub Run(oDoc As Object)" in text
    assert "参考" in text

def test_load_refs_missing_dir_is_empty(tmp_path):
    assert ailine.load_refs(tmp_path / "nope") == ""


# --- ヘルパ・ライブラリ（呼ぶだけ） ----------------------------------------

def test_load_helpers_catalog_and_files():
    catalog, files = ailine.load_helpers(ailine.DEFAULT_HELPERS)
    assert any(f.name.endswith(".bas") for f in files)
    assert "SortByColumn" in catalog
    assert "InsertBarChart" in catalog
    assert "Call" in catalog          # Call 形式で呼ばせる指示が入っている

def test_load_helpers_missing_dir(tmp_path):
    catalog, files = ailine.load_helpers(tmp_path / "nope")
    assert catalog == "" and files == []


# --- 自作の道: openpyxl による太字後付け（LO が xlsx に書けない分） ----------

def test_apply_style_directives_bold(tmp_path):
    p = _book(tmp_path, [["h1", "h2"], ["a", 1]])
    code = "Sub Run(oDoc As Object)\n Call StyleBold(oDoc, 0, 0, 1, 0)\nEnd Sub"
    n = ailine.apply_style_directives(p, code)
    assert n == 1
    ws = openpyxl.load_workbook(p).active
    assert ws.cell(1, 1).font.bold is True
    assert ws.cell(1, 2).font.bold is True
    assert ws.cell(2, 1).font.bold in (None, False)   # データ行は太字にしない

def test_apply_style_directives_none(tmp_path):
    p = _book(tmp_path, [["a", 1]])
    assert ailine.apply_style_directives(p, "Sub Run(oDoc As Object)\nEnd Sub") == 0


# --- snapshot / 差分（no-op ガードの核） -----------------------------------

def _book(tmp_path, rows):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p

def test_diff_detects_value_change(tmp_path):
    p = _book(tmp_path, [["a", 1], ["b", 2]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.active.cell(1, 3, "new"); wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any("new" in ln for ln in lines)

def test_diff_noop_when_unchanged(tmp_path):
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    after = ailine.snapshot(p)   # 何も変えない
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is False       # ← no-op を正しく no-op と判定
    assert lines == []

def test_diff_detects_new_sheet(tmp_path):
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.create_sheet("集計"); wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any("集計" in ln for ln in lines)

def test_diff_detects_fill_only_change(tmp_path):
    # 値でなく背景色だけ変えても『変化した』と見えること（no-op 誤検出を防ぐ）
    from openpyxl.styles import PatternFill
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(1, 1).fill = PatternFill("solid", fgColor="FFCCCC")
    wb.save(p)
    after = ailine.snapshot(p)
    changed, _ = ailine.diff_snapshots(before, after)
    assert changed is True

def test_diff_detects_border_only_change(tmp_path):
    # 罫線だけの変更も検出すること（罫線ヘルパが no-op 誤判定されないため）
    from openpyxl.styles import Border, Side
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    thin = Side(style="thin")
    wb.active.cell(1, 1).border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb.save(p)
    after = ailine.snapshot(p)
    changed, _ = ailine.diff_snapshots(before, after)
    assert changed is True

def test_diff_detects_merge(tmp_path):
    p = _book(tmp_path, [["a", "b"]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.active.merge_cells("A1:B1"); wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any("結合" in ln for ln in lines)

def test_diff_detects_colwidth(tmp_path):
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p); wb.active.column_dimensions["A"].width = 30; wb.save(p)
    after = ailine.snapshot(p)
    changed, lines = ailine.diff_snapshots(before, after)
    assert changed is True
    assert any("列幅" in ln for ln in lines)

def test_diff_detects_align_only_change(tmp_path):
    # 中央揃えだけの変更も検出すること（AlignCenter ヘルパが no-op 誤判定されないため）
    from openpyxl.styles import Alignment
    p = _book(tmp_path, [["a", 1]])
    before = ailine.snapshot(p)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(1, 1).alignment = Alignment(horizontal="center")
    wb.save(p)
    after = ailine.snapshot(p)
    changed, _ = ailine.diff_snapshots(before, after)
    assert changed is True


# --- 文書の説明 ------------------------------------------------------------

def test_describe_book_lists_headers(tmp_path):
    p = _book(tmp_path, [["商品", "金額", "在庫"], ["りんご", 1200, 8]])
    desc = ailine.describe_book(p)
    assert "列0=商品" in desc
    assert "列1=金額" in desc
    assert "シート一覧" in desc
