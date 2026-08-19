"""EXTRACT の事後条件(check_extract)の純ロジック検体（DESIGN-20260820-extract-op.md §事後条件）。

★ 昨夜の実弾2件をそのまま検体化する:
  - 行抽出は意味は当たったが全セルが文字列化（getString/setString コピーで
    '59,400' のようにカンマごと焼き込む） → test_check_extract_fail_type_stringified
  - 列抽出は空シートを作って exit 0（「できたふり」） → test_golden_postcondition.py の
    extract_fail_missing_sheet（run_postcondition golden）が同型を既に凍結している。
    ここでは check_extract を直接呼ぶ純ロジックの4点（行数一致・値と型の保存・両側の網羅・
    元シート無変更）を1つずつ fail させる検体 + pass 検体を集める。
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ailine  # noqa: E402

_ARGS = {"col": "金額", "cmp": "gte", "value": 40000.0,
         "_target_sheet": "Sheet", "_new_sheet": "金額40000以上"}


def _book(tmp_path, name, src_rows, out_rows=None, out_sheet_name="金額40000以上"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"
    for row in src_rows:
        ws.append(row)
    if out_rows is not None:
        out = wb.create_sheet(out_sheet_name)
        for row in out_rows:
            out.append(row)
    wb.save(p)
    return p


_SRC_ROWS = [["商品", "金額"], ["a", 30000], ["b", 50000], ["c", 45000]]


def test_check_extract_pass(tmp_path):
    """3行中2行(b,c)が一致 → 出力が同じ順・同じ値/型で2行 → pass。"""
    p = _book(tmp_path, "b.xlsx", _SRC_ROWS,
              out_rows=[["商品", "金額"], ["b", 50000], ["c", 45000]])
    status, reason = ailine.check_extract(p, dict(_ARGS), header_row=1)
    assert status == "pass", reason
    assert "3行中2行が一致" in reason
    assert "2行を抽出" in reason


def test_check_extract_fail_row_count_mismatch(tmp_path):
    """① 行数一致: 出力が1行しかない（期待2行）→ fail。"""
    p = _book(tmp_path, "b.xlsx", _SRC_ROWS, out_rows=[["商品", "金額"], ["b", 50000]])
    status, reason = ailine.check_extract(p, dict(_ARGS), header_row=1)
    assert status == "fail"
    assert "行数が期待と不一致" in reason


def test_check_extract_fail_type_stringified(tmp_path):
    """② 値と型の保存: 昨夜の実弾そのもの ── 数値のはずの金額セルが文字列 '50000' で
       書かれている（getString/setString で焼いたことの再現）。行数・見た目の値は一致
       していても fail し、理由に型の食い違いが分かる形で出ること。"""
    p = _book(tmp_path, "b.xlsx", _SRC_ROWS,
              out_rows=[["商品", "金額"], ["b", "50000"], ["c", 45000]])
    status, reason = ailine.check_extract(p, dict(_ARGS), header_row=1)
    assert status == "fail"
    assert "str" in reason and "int" in reason, reason


def test_check_extract_fail_one_sided_coverage(tmp_path):
    """③ 両側の網羅: 出力の行数は期待どおり(2行)だが、条件を満たさない行('a')を含み、
       本来含むべき行('c')が抜けている（多く含める/少なく埋めるを同時に検体化）。"""
    p = _book(tmp_path, "b.xlsx", _SRC_ROWS,
              out_rows=[["商品", "金額"], ["b", 50000], ["a", 30000]])
    status, reason = ailine.check_extract(p, dict(_ARGS), header_row=1)
    assert status == "fail"
    assert "3行中2行が一致" in reason


def test_check_extract_fail_source_modified(tmp_path):
    """④ 元シートが無変更: source_book(適用前)と path(適用後)で元シートの値が違う
       （読むだけのはずの EXTRACT が元データを書き換えた）→ fail。"""
    before = _book(tmp_path, "before.xlsx", _SRC_ROWS)
    after = _book(tmp_path, "after.xlsx",
                  [["商品", "金額"], ["a", 30000], ["b", 999999], ["c", 45000]],
                  out_rows=[["商品", "金額"], ["b", 999999], ["c", 45000]])
    status, reason = ailine.check_extract(after, dict(_ARGS), header_row=1, source_book=before)
    assert status == "fail"
    assert "変更されています" in reason


def test_check_extract_pass_with_unmodified_source_book(tmp_path):
    """④ を実際に確認できる形の pass: source_book が元シートと完全一致 → 元シート無変更、
       と明言した理由文で pass。"""
    before = _book(tmp_path, "before.xlsx", _SRC_ROWS)
    after = _book(tmp_path, "after.xlsx", _SRC_ROWS,
                  out_rows=[["商品", "金額"], ["b", 50000], ["c", 45000]])
    status, reason = ailine.check_extract(after, dict(_ARGS), header_row=1, source_book=before)
    assert status == "pass", reason
    assert "元シート無変更" in reason


def test_check_extract_fail_zero_target_rows(tmp_path):
    """止血1: 元シートにデータ行が0件（何も検証できない）を合格にしない。"""
    p = _book(tmp_path, "b.xlsx", [["商品", "金額"]], out_rows=[["商品", "金額"]])
    status, reason = ailine.check_extract(p, dict(_ARGS), header_row=1)
    assert status == "fail"
    assert reason == ailine._ZERO_TARGET_REASON
