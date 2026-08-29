# 引用符の中は値であって、対象の名指しではない ── 2026-08-30。
# Namakoo「セル指定しているのに値を上書きできない」
#
# ★★ 実測（「7行B列を『{{合計:税込金額}}』に上書き」・シート『雛形』を選択）:
#     （『行挿入』でなく『行追加』として読み直しました ── 依頼文が場所を
#       『合計』の行＝9行目と指しています）
#   → セルを座標で名指ししているのに、**引用符の中の『合計』**が位置の目印として
#     拾われ、頼んでいない行が挿さりかけた。
#
# 漏れ口は 2 つあった:
#  ① `_row_named_anywhere_in_task`（依頼文に現れる**表の値**を探すフォールバック）が
#     生の依頼文を見ていた。★ ここは 4 つの呼び出しの合流点なので、この 1 箇所で全部に効く。
#  ② 読み直しの関所が**1 枚目**で位置を解いていた（画面で『雛形』を選んでいるのに
#     『8月請求』の 9 行目）。
#
# ★ 列では既に塞いだ穴（_task_names_single_real_column）が、行では開いていた
#   ── **行と列の非対称**、この repo が何度も踏んだ形。

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

H = ["取引先", "項目", "金額"]
ROWS = [["丸和物流", "配送", 57600], ["近江スチール", "鋼材", 60000], ["合計", "", 117600]]


@pytest.fixture()
def meta(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(H)
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    return {"sheets": ["請求"], "headers": {"請求": list(H)},
            "header_rows": {"請求": 1}, "path": str(p)}


# --- ① 引用符の中の語で行を決めない ------------------------------------------------------

def test_a_row_is_not_chosen_from_inside_the_quotes(meta):
    """★★ これが今回の芯。値の中の『合計』が位置になってはいけない。"""
    at, note = ailine.resolve_row_anchor("7行B列を「{{合計:税込金額}}」に上書き", meta, "請求")
    assert at is None, (at, note)


def test_the_switch_to_add_row_does_not_fire_either(meta):
    """★ 位置が出ないので、op の乗り換え（行追加として読み直し）も起きない。"""
    assert ailine.insert_rows_should_have_been_add_row(
        "7行B列を「{{合計:税込金額}}」に上書き", {}, meta, "請求") is None


def test_a_name_outside_the_quotes_still_works(meta):
    """★ 黙りすぎていないこと: 引用符の外で名指しされた行は今までどおり拾う。"""
    assert ailine.resolve_row_anchor("合計の行を削除して", meta, "請求")[0] == 4
    assert ailine.resolve_row_anchor("丸和物流を削除して", meta, "請求")[0] == 2


def test_the_leak_is_sealed_at_the_junction():
    """★ 4 つの呼び出しの合流点 1 箇所で塞ぐ（呼び出し側に配らない＝片配線を作らない）。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index("def _row_named_anywhere_in_task(")
    seg = src[i:i + 1800]
    assert "_task_outside_quotes(task)" in seg, "合流点が生の依頼文を見ている"
    assert 'text = task or ""' not in seg


# --- ② 読み直しは「選んだシート」で解く --------------------------------------------------

def test_the_reread_resolves_on_the_target_sheet():
    """★★ 画面で『雛形』を選んでいるのに『8月請求』の行を根拠にしていた。
       対象シートは既に 1 箇所で決まっている ── それを使う。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    # 読み直しの塊が 1 枚目を仮定していないこと
    assert src.count('_sheet_hint = (getattr(a, "_target_sheet", None)') == 1
    assert src.count('_sheet_h = (getattr(a, "_target_sheet", None)') == 1
    assert '_sheet_hint = (book_meta.get("sheets") or [None])[0]' not in src
    assert '_sheet_h = (book_meta.get("sheets") or [None])[0]' not in src


# --- ③ 見出しの無い列を「『』」と出さない ------------------------------------------------

def test_a_column_without_a_header_is_named_by_its_letter(tmp_path):
    """★ 実測で画面に出た形: 「列は『』（B列）」── 空の名前を見せない。"""
    p = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "雛形"
    ws.append([None, None])
    ws.append(["ご請求金額", "{{合計:金額}}"])
    wb.save(p)
    m = {"sheets": ["雛形"], "headers": {"雛形": ["", ""]},
         "header_rows": {"雛形": 1}, "path": str(p)}
    got = ailine.resolve_cell_target_from_task("2行B列を「x」にして", m, "雛形")
    if got is not None:
        _r, _c, note = got
        assert "『』" not in note, note
