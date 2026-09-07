# -*- coding: utf-8 -*-
"""壊れた式が生まれた回は、原本へ被せない（2026-09-08・盲検の検品が拾った）。

★★ 出所: 盲検で「丸和物流の行を削除して」を打ったところ、合計式が
  `=SUM(#REF!:INDEX(E:E,ROW()-1))` に壊れ、**原本に `#REF!` が残った**。
  判定は `⚠`（嘘はついていない）。★ だが **undo を打たない限り、合計が壊れた
  請求書が手元に残る** ── 実害としては false ✓ より重い。

★★ 検出は**既にできていた**（formula_health が「新たにエラーになったセル」を数え、
  ⚠ を出し ✓ も降ろしていた）。足りなかったのは**帰結**で、「言ったうえで書いて」いた。

★ だから直しは検出器を増やすのではなく、**同じ器官の呼び先を 1 つ増やす**:
  反映の関所（✓ を出す唯一の場所）で、新たに壊れた式が在れば原本へ被せない。
  ★ 削除に限らない位置に置いた ── Namakoo「同様のケースが別例で発生しないようにしたい」。
★ 逃げ道は増やさない。結果が要るなら `--copy` で今までどおり受け取れる。
"""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import openpyxl
import pytest

import ailine
from ailine_core import formula_health


def _book_with_total(path, broken: bool):
    """合計式を持つ表。broken=True なら、その式が #REF! に壊れている。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(["取引先", "金額"])
    ws.append(["丸和物流", 57600])
    ws.append(["近江スチール", 60000])
    ws.append(["合計", None])
    ws["B4"] = "=SUM(#REF!:B3)" if broken else "=SUM(B2:B3)"
    # ★ キャッシュ値（openpyxl は data_only でこれを読む）
    wb.save(path)
    return path


def _inject_cached(path, cell: str, value):
    """LibreOffice/Excel が書くキャッシュ値を手で入れる（検体の治具）。"""
    import re
    import zipfile
    tmp = Path(str(path) + ".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                s = data.decode("utf-8")
                s = re.sub(
                    r'(<c r="' + cell + r'"[^>]*)(/>|>.*?</c>)',
                    '<c r="' + cell + '" t="e"><f>SUM(#REF!:B3)</f>'
                    "<v>" + str(value) + "</v></c>", s, count=1)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)
    return path


def test_the_organ_names_only_newly_broken_cells(tmp_path):
    """★ 判定そのもの ── **新たに**壊れたセルだけを返す（前から壊れている分は数えない）。"""
    before = _inject_cached(_book_with_total(tmp_path / "b.xlsx", broken=True), "B4", "#REF!")
    after = _inject_cached(_book_with_total(tmp_path / "a.xlsx", broken=True), "B4", "#REF!")
    assert formula_health.new_error_cells(before, after) == {}, "前から壊れている分を数えている"

    clean = _book_with_total(tmp_path / "c.xlsx", broken=False)
    got = formula_health.new_error_cells(clean, after)
    assert got, "新しく壊れたのに気づいていない"


def _run_finish_apply(tmp_path, before_path, after_path, capsys):
    """✓ を出す唯一の関所を、**原本を置き換える経路**（--inplace）で通す。"""
    work = tmp_path / "w"
    work.mkdir(exist_ok=True)
    a = argparse.Namespace(inplace=True, task="丸和物流の行を削除して", json=False,
                           model="m", keep_backups=3, copy=False)
    ok = ailine._finish_apply(a, before_path, after_path, work, {"op": "DELETE_ROWS"},
                              machine_verified=True, scope="操作:行削除 削除位置:2 行数:1",
                              scope_note="", warning_count=0)
    return ok, capsys.readouterr().out


def test_a_newly_broken_formula_stops_the_write(tmp_path, capsys):
    """★ 原本へ被せないこと。★ どの op でも効く位置に在ることを、削除以外でも縛る。"""
    before = _book_with_total(tmp_path / "orig.xlsx", broken=False)
    keep = before.read_bytes()
    after = _inject_cached(_book_with_total(tmp_path / "orig.out.xlsx", broken=True),
                           "B4", "#REF!")
    ok, shown = _run_finish_apply(tmp_path, before, after, capsys)
    assert ok is False, shown
    assert "壊れた式" in shown and "#REF!" in shown, shown
    assert before.read_bytes() == keep, "★ 原本が書き換わっている"
    assert "変更していません" in shown, shown


def test_a_clean_result_still_goes_through(tmp_path, capsys):
    """★ 対で縛る ── 壊れていない回は今までどおり反映されること（塞ぎすぎの検分）。"""
    before = _book_with_total(tmp_path / "ok.xlsx", broken=False)
    after = _book_with_total(tmp_path / "ok.out.xlsx", broken=False)
    ok, shown = _run_finish_apply(tmp_path, before, after, capsys)
    assert ok is True, shown
    assert "壊れた式" not in shown, shown
