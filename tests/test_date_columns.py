# 日付の列を「数値として」検証できるようにする（2026-08-24）。
#
# ★ 実測した穴（実 7B + 実 LibreOffice）: 出納帳を「日付の古い順に並べ替えて」と頼むと、
#   LibreOffice は**正しく並べたのに** ailine が
#   「事後条件の検証対象が0件（何も検証できていない）（数値でない 3 行は対象外）」で拒否し、
#   原本に反映しない。理由は 1 つ ── 事後条件が openpyxl の datetime を「数値でない」と
#   見ていた。表計算の日付は**シリアル値という数値**なので、これは検証側の取り違え。
#   ★ 出納帳・領収書を扱う道具で日付の並べ替えができないのは致命的（台帳の DATE_CALC
#   2 件より重い穴だった）。
#
# 契約:
#   ① 日付/日時のセルは検証の対象になる（「対象外」に落ちない）
#   ② 日付の並べ替えの順序が実際に検証される（恒真でない ── 逆順なら fail）
#   ③ 数値でも日付でもないセル（文字列）は今までどおり対象外

import datetime as dt
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "_numeric_value"),
    reason="日付の数値化 未実装（契約は凍結済み）",
    strict=True,
)


@needs_impl
@pytest.mark.parametrize("value,expected", [
    (5, 5.0),
    (2.5, 2.5),
    (dt.date(2026, 3, 26), 46107.0),
    (dt.datetime(2026, 3, 26, 12, 0), 46107.5),
    ("abc", None),
    (None, None),
    (True, None),                     # bool は数値セルとして扱わない（従来どおり）
])
def test_numeric_value_maps_dates_to_serials(value, expected):
    assert ailine._numeric_value(value) == expected


def _book(tmp_path, order):
    p = tmp_path / "s.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "出納"
    ws.append(["日付", "金額"])
    for d in order:
        ws.append([dt.datetime(d.year, d.month, d.day), 1])
    wb.save(p)
    return p


@needs_impl
def test_sort_postcondition_accepts_a_date_column(tmp_path):
    """① ② 昇順に並んだ日付列 → ok（対象外に落ちない）。"""
    book = _book(tmp_path, [dt.date(2026, 1, 5), dt.date(2026, 3, 20), dt.date(2026, 5, 1)])
    status, reason = ailine.check_sort(book, {"col": "日付", "order": "asc",
                                                "_target_sheet": "出納"})
    # ★ 治具の訂正（封印者ナギ・2026-08-24）: 事後条件の合格語は "ok" でなく "pass"。
    #   assert の意図（日付列が検証され、対象外に落ちない）は不変。
    assert status == "pass", f"日付列が検証できていない: {status} / {reason}"
    assert "対象外" not in (reason or ""), reason


# ★ needs_impl を付けない ── 実装前は恒真（今は「対象0件」で ok にならないだけ）。
#   実装後に初めて意味を持つ恒真殺しとして置く。
def test_sort_postcondition_still_catches_a_wrong_order(tmp_path):
    """恒真殺し: 降順に並んでいるのに昇順を宣言したら fail。"""
    book = _book(tmp_path, [dt.date(2026, 5, 1), dt.date(2026, 3, 20), dt.date(2026, 1, 5)])
    status, _reason = ailine.check_sort(book, {"col": "日付", "order": "asc",
                                                 "_target_sheet": "出納"})
    assert status == "fail", "順序が逆なのに合格にした（恒真）"


def test_string_column_is_still_excluded(tmp_path):
    """③ 退行防止: 文字列の列は今までどおり検証対象にならない。"""
    p = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "s"
    ws.append(["名前", "金額"])
    for n in ("b", "a"):
        ws.append([n, 1])
    wb.save(p)
    status, reason = ailine.check_sort(p, {"col": "名前", "order": "asc", "_target_sheet": "s"})
    assert status != "pass", f"文字列列を数値として検証した: {status} / {reason}"
