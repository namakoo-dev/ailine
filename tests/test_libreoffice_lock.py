# LibreOffice で開いたまま書き換えない ── 2026-08-30。
# Namakoo「LO を開いた状態で表の更新はできないの？」
#
# ★★ 実測: **Excel のロック（~$名前）は見ていたのに、LibreOffice のロック
#   （.~lock.名前#）は見ていなかった** ── ailine 自身が LibreOffice を使う道具なのに、
#   片方だけ守られていた（また「行と列の非対称」と同じ形）。
#   ロックファイルを置いても素通りして書き込めた。
#
# ★★ 危ないのは「書けること」ではなく、この順序:
#     ① 人が LO で開く → ② ailine が書く（成功する）→ ③ 人が LO 側で保存する
#     → **開いた時点の古い中身で上書き**され、ailine の変更が黙って消える。
#   この道具が一番嫌う「静かに失われる」形なので、書く前に止める。
#
# ★ 断りは Excel と同じ形にそろえる ── 何が起きるかを言い、残骸の可能性も言う
#   （断れない時は開示する、の裏返し: 断る時は理由と逃げ道を出す）。

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from _product_source import window_around  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む


@pytest.fixture()
def book(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["商品", "売上"])
    ws.append(["りんご", 100])
    wb.save(p)
    return p


def test_a_libreoffice_lock_is_detected(book):
    """★ .~lock.<名前># が在れば止める（実測で素通りしていた）。"""
    (book.parent / f".~lock.{book.name}#").write_text("x", encoding="utf-8")
    got = ailine.check_excel_lock(book)
    assert got is not None, "LibreOffice のロックを見ていない"
    kind, detail = got
    assert kind == "libreoffice", kind
    assert ".~lock." in detail


def test_an_excel_lock_is_still_detected(book):
    """★ 元から在った Excel 側を壊していないこと。"""
    (book.parent / f"~${book.name}").write_text("x", encoding="utf-8")
    kind, _detail = ailine.check_excel_lock(book)
    assert kind == "excel"


def test_no_lock_no_complaint(book):
    """★ 黙りすぎ・鳴りすぎの両方を見る ── ロックが無ければ何も言わない。"""
    assert ailine.check_excel_lock(book) is None


def test_both_lock_shapes_are_checked_in_one_place():
    """★ 片方だけ守る形に戻らないよう、2 つが同じ関数に在ることを縛る。"""
    seg = window_around("def check_excel_lock(", after=2600)
    assert 'f"~${book.name}"' in seg, "Excel のロックを見ていない"
    assert 'f".~lock.{book.name}#"' in seg, "LibreOffice のロックを見ていない"


def test_the_refusal_says_what_would_be_lost():
    """★★ 「書けません」で終わらせない ── **あとで消える**ことを言う。
       それがこの断りの理由そのものなので、文言が消えたら赤くする。"""
    seg = window_around('elif kind == "libreoffice":', after=900)
    assert "開いた時点の内容で上書き" in seg
    assert "残骸" in seg, "残骸だった場合の逃げ道が無い"
