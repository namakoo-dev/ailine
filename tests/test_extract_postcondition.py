"""EXTRACT の事後条件(check_extract)の純ロジック検体（4点を1つの位置対応比較で同時に見る
   設計・コミット 2edcb08「EXTRACT op」参照）。

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

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
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


# ── ★ 2026-09-15: この道具が作った**計算列**で絞れること（盲検の買い手役が踏んだ）──────────
#
# ★★ 事故: 「在庫数から発注点を引いた過不足という列を作って」→ ✓。続けて
#   「過不足がマイナスの行を抜き出して」→ **出力は正解 3 行なのに「6行中0行が一致」で × ・exit 1**。
#   原本に入らず、正しい答えが「不合格にされた .out.xlsx」の中にだけ残る。
#   ★ 原因: 述語が**式ビュー**を読んでいた（`'=C2-D2' < 0` は常に偽）。キャッシュ値は同じ
#     ファイルに在り、行の中身の比較（_row_as_shown）は既に値ビューを読んでいた
#     ── **同じ関数の中で述語だけが配線されていなかった**。

def _book_with_formula_column(tmp_path, name="f.xlsx", *, sheet_title="Sheet",
                              lead_sheet=None, out_rows=None):
    """C 列 = A 列 − B 列 の**式**を持つ表（キャッシュ値つき）。

    ★ キャッシュ値は openpyxl では書けないので、**式ビューと値ビューの 2 冊**を作って
      値ビュー側を data_only 相当として渡す…のではなく、ここでは実際の保存形に合わせて
      「式 + キャッシュ値」を直接 XML に持たせる代わりに、LibreOffice が保存した形を模す。
      → openpyxl だけで再現できないので、**式セルにキャッシュが無い**場合の
        フォールバック（式のまま返す）も含めて、下の 2 本で両方を縛る。
    """
    p = tmp_path / name
    wb = openpyxl.Workbook()
    if lead_sheet:                      # ★ 対象が 1 枚目でない配置（ヘルパのシート指定の番人）
        wb.active.title = lead_sheet
        ws = wb.create_sheet(sheet_title)
    else:
        ws = wb.active
        ws.title = sheet_title
    ws.append(["品番", "在庫数", "発注点", "過不足"])
    for i, (code, stock, point) in enumerate([("A-100", 45, 50), ("A-200", 8, 20),
                                              ("B-010", 120, 40), ("B-020", -3, 30)], start=2):
        ws.append([code, stock, point, f"=B{i}-C{i}"])
    if out_rows is not None:
        out = wb.create_sheet("過不足0未満")
        for row in out_rows:
            out.append(row)
    wb.save(p)
    return p


def test_a_formula_column_without_a_cache_is_not_counted_as_matching(tmp_path):
    """★ キャッシュ値が**無い**式（openpyxl が書いたまま・一度も計算されていない）は、
       数値として当たらない ── ここで当ててしまうと「計算していない値で判定した」ことになる。
       ★ この検体は**直す前も後も同じ**（フォールバックの契約）。動いたら意味が変わった印。"""
    p = _book_with_formula_column(tmp_path, out_rows=[["品番", "在庫数", "発注点", "過不足"]])
    args = {"col": "過不足", "cmp": "lt", "value": 0.0,
            "_target_sheet": "Sheet", "_new_sheet": "過不足0未満"}
    status, reason = ailine.check_extract(p, args, header_row=1)
    assert status == "pass", reason
    assert "4行中0行が一致" in reason, reason


def test_the_predicate_reads_the_value_view_not_the_formula_text(tmp_path, monkeypatch):
    """★★ 事故そのもの。式のセルに**キャッシュ値が在る**とき、述語はそれを読むこと。

    ★ 実機（LibreOffice）が保存したキャッシュを openpyxl では書けないので、BookView の
      値ビューだけを差し替えて「キャッシュが在る」状態を作る（窒息点は BookView の 1 口）。
    """
    p = _book_with_formula_column(
        tmp_path, out_rows=[["品番", "在庫数", "発注点", "過不足"],
                            ["A-100", 45, 50, -5], ["A-200", 8, 20, -12], ["B-020", -3, 30, -33]])
    cache = {2: -5, 3: -12, 4: 80, 5: -33}

    real = ailine.BookView.cell_value

    def fake_cell_value(self, row, col, sheet=None):
        if col == 4 and row in cache and (sheet or "Sheet") == "Sheet":
            return cache[row]
        return real(self, row, col, sheet)

    monkeypatch.setattr(ailine.BookView, "cell_value", fake_cell_value)
    args = {"col": "過不足", "cmp": "lt", "value": 0.0,
            "_target_sheet": "Sheet", "_new_sheet": "過不足0未満"}
    status, reason = ailine.check_extract(p, args, header_row=1)
    assert status == "pass", reason
    assert "4行中3行が一致" in reason, reason


def test_the_helper_reads_the_formula_from_the_target_sheet_not_the_first(tmp_path, monkeypatch):
    """★ ついでに塞いだ潜在の穴: 「写す側が実際に写す値」のヘルパが、シート名なしで
       式ビューを引いていた ── 対象が 2 枚目以降だと**別のシートの式**を見る。

    ★★ 初版の検体は**この穴を踏めなかった**（変異を当てても緑のまま）。理由は
      「キャッシュが無ければ元の値に落とす」フォールバックが、間違ったシートを読んだ回まで
      救ってしまうから。★ だから 1 枚目に**式とキャッシュの両方**が在る状態を作る ──
      そこを読んだら答えが変わる、という形にして初めて番人になる。
    """
    p = tmp_path / "two.xlsx"
    wb = openpyxl.Workbook()
    lead = wb.active
    lead.title = "説明"
    lead.append(["x", "y", "z", "w"])
    for _ in range(4):
        lead.append([None, None, None, "=999"])      # 1 枚目は式（キャッシュは下で足す）
    ws = wb.create_sheet("Sheet")
    ws.append(["品番", "在庫数", "発注点", "過不足"])
    for code, stock, point, diff in [("A-100", 45, 50, -5), ("A-200", 8, 20, -12),
                                     ("B-010", 120, 40, 80), ("B-020", -3, 30, -33)]:
        ws.append([code, stock, point, diff])        # 対象は素の数値（式ではない）
    out = wb.create_sheet("過不足0未満")
    for row in [["品番", "在庫数", "発注点", "過不足"], ["A-100", 45, 50, -5],
                ["A-200", 8, 20, -12], ["B-020", -3, 30, -33]]:
        out.append(row)
    wb.save(p)

    real = ailine.BookView.cell_value

    def fake_cell_value(self, row, col, sheet=None):
        # ★ 1 枚目（sheet 省略時の既定）にだけキャッシュが在る状態
        if col == 4 and sheet in (None, "説明"):
            return 999
        return real(self, row, col, sheet)

    monkeypatch.setattr(ailine.BookView, "cell_value", fake_cell_value)
    args = {"col": "過不足", "cmp": "lt", "value": 0.0,
            "_target_sheet": "Sheet", "_new_sheet": "過不足0未満"}
    status, reason = ailine.check_extract(p, args, header_row=1)
    assert status == "pass", reason
    assert "4行中3行が一致" in reason, reason


def test_dedup_keys_read_the_value_view_too(tmp_path, monkeypatch):
    """★ 抽出の兄弟。鍵の列が式のとき、式ビューを読むと '=A2' と '=A3' を別の鍵と数え、
       **重複が永久に見つからない**。片方だけ直さない。"""
    p = tmp_path / "d.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"
    ws.append(["商品", "区分"])
    # ★ 式の**文字列は行ごとに違う**が、計算後の値は全部同じ ── ここが肝。
    #   文字列が同じ検体だと、式ビューを読む壊れた版でも同じ鍵になり、番人が鳴らない
    #   （初版はこれで変異試験が緑のままだった）。
    for i, f in enumerate(["=1+0", "=2-1", "=3-2", "=4-3"], start=2):
        ws.append([f, i])
    out = wb.create_sheet("重複除去")
    out.append(["商品", "区分"])
    out.append(["A", 2])                       # 1 行だけ残るのが正解
    wb.save(p)

    real = ailine.BookView.cell_value

    def fake_cell_value(self, row, col, sheet=None):
        if col == 1 and row >= 2:
            return "A"                          # 4 行とも同じ鍵
        return real(self, row, col, sheet)

    monkeypatch.setattr(ailine.BookView, "cell_value", fake_cell_value)
    st, reason = ailine.check_dedup(p, {"_new_sheet": "重複除去", "keys": ["商品"],
                                        "_target_sheet": "Sheet"}, 1)
    assert st == "pass", reason
    assert "4行中1行を残しました" in reason, reason
