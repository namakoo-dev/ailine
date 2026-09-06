"""行の名指しは、**見出し行を data 行として拾わない**こと（2026-09-06）。

★★ 出所（実物の請求書で実測）: 見出しが 16 行目の表に「金額の列を全部 0 にして」と
  頼むと、こう返っていた ──

      （『行挿入』でなく『行追加』として読み直しました ──
        依頼文が場所を**『金額』の行＝16行目**と指しています）
      ？ できませんでした ── 行番号『-1』が不正です

  ★ **見出しの語『金額』を「行の名前」として拾い、見出し行そのものを指していた。**
    列の書き換えが行の追加に読み替わっている ── ちょうど別の操作だ。

★ 真因は片配線: `resolve_row_anchor` の呼び出し **4 箇所のうち 1 箇所だけ**が
  `header_row` を渡していなかった（既定の 1 行目で走査していた）。
  ★ `book_meta` は `header_rows` を持っている ── 呼び出し側に持たせず、
    `insert_rows_should_have_been_add_row` の中で引く（次に呼ぶ人も間違えない）。

★★ 同じ日に**同じ形の穴を 2 つ**踏んでいる: 朝に足した `formula_loss_advisory` も
  「見出しを 1 行目と決め打ち」で実物では黙っていた。
  ★ **見出し行は 1 行目とは限らない** ── 実物の帳票はほぼ例外なく違う。

★ 止まったのは引数の検査（`行番号『-1』が不正`）で、**誤読を見抜いたからではない**。
  別の言い回しで偶然に正しい引数が出れば、黙って違う操作が走りうる ── だから
  「壊す前に止まった」を根拠に放置しない。
"""
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import ailine  # noqa: E402
from _product_source import product_text  # noqa: E402

ROWS = [("ボルト", 10, 50), ("ナット", 5, 20), ("ワッシャ", 8, 15)]


def _meta(tmp_path, header_row: int):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明細"
    for _ in range(header_row - 1):
        ws.append([])
    ws.append(["品名", "数量", "単価"])
    for r in ROWS:
        ws.append(r)
    p = tmp_path / f"h{header_row}.xlsx"
    wb.save(p)
    return ailine.build_book_meta(p, header_rows={"明細": header_row})


# --- ① 見出しの語を「行の名前」として拾わない -------------------------------

@pytest.mark.parametrize("header_row", [1, 3, 16])
@pytest.mark.parametrize("task", [
    "品名の列を全部 0 にして",
    "数量の列を全部 0 にして",
    "単価の列をすべて 0 にして",
])
def test_a_header_word_is_never_read_as_a_row(tmp_path, header_row, task):
    """★ 実物の再現形 ── 見出しの語で『行追加』に読み替わらないこと。"""
    bm = _meta(tmp_path, header_row)
    assert ailine.insert_rows_should_have_been_add_row(task, {}, bm, "明細") is None, task


@pytest.mark.parametrize("header_row", [1, 3, 16])
def test_the_anchor_resolver_ignores_the_header_row(tmp_path, header_row):
    """★ 器官そのもの ── 見出し行を渡せば、見出しの語では行が決まらない。"""
    bm = _meta(tmp_path, header_row)
    at, _note = ailine.resolve_row_anchor("品名の列を全部 0 にして", bm, "明細",
                                          header_row=header_row)
    assert at is None, at


# --- ② 発火すべき側は発火する（★ 対で縛る ── 潰しすぎない）------------------

@pytest.mark.parametrize("header_row", [1, 3, 16])
@pytest.mark.parametrize("task", [
    "ボルトの下に新品を追加して",
    "ナットの上に新品を追加して",
])
def test_a_real_row_name_still_fires(tmp_path, header_row, task):
    """★ 中身で行を指す依頼は、見出し行が何行目でも従来どおり読み替わること。"""
    bm = _meta(tmp_path, header_row)
    assert ailine.insert_rows_should_have_been_add_row(task, {}, bm, "明細"), task


# --- ③ 配線（★ 呼び出し側 4 箇所すべてが見出し行を渡す）---------------------

def test_every_caller_passes_the_header_row():
    """★ 片配線の再発を止める ── `resolve_row_anchor` を素で呼ばない。

    ★ 2026-09-06 に測ったとき、4 箇所のうち 1 箇所だけが渡していなかった。
      1 箇所直して満足せず、**入口側から数える**（この repo の系譜）。
    """
    src = product_text()
    calls = [ln.strip() for ln in src.splitlines() if "resolve_row_anchor(" in ln
             and not ln.lstrip().startswith("def ")]
    assert calls, "呼び出しが 1 つも見つからない（検出が壊れている疑い）"
    bare = [c for c in calls if "header_row" not in c]
    assert not bare, f"見出し行を渡していない呼び出しがある: {bare}"


def test_the_organ_reads_the_header_row_from_the_book(tmp_path):
    """★ 呼び出し側に持たせない ── book_meta から引いていること。"""
    body = product_text().split("def insert_rows_should_have_been_add_row")[1]
    body = body[:body.index(chr(10) + "def ")]
    assert "header_rows" in body, "book_meta の header_rows を引いていない"
    assert "header_row=" in body, "resolve_row_anchor に渡していない"


# --- ④ 実物（★ 他社の請求書・repo の外）------------------------------------

def test_it_holds_on_the_real_invoice(tmp_path):
    src = Path("C:/Dev/_fixtures/misoca_invoice_blackline.xlsx")
    if not src.is_file():
        pytest.skip("実物の検体が手元に無い（C:/Dev/_fixtures）")
    import shutil
    p = tmp_path / "inv.xlsx"
    shutil.copy2(src, p)
    bm = ailine.build_book_meta(p, header_rows={"misoca_invoice": 16})
    got = ailine.insert_rows_should_have_been_add_row(
        "金額の列を全部 0 にして", {}, bm, "misoca_invoice")
    assert got is None, got
