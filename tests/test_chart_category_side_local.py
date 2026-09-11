# -*- coding: utf-8 -*-
"""グラフ: 項目列が値列より**右**にあっても、役割が反転しないこと（実機・2026-09-11）。

★★ 実装より先に置いた赤い検体（docs/PENDING-20260910-グラフの項目列と転記のキー列.md ①）。

  真因（2026-09-10 に実機で特定）:
    `addNewByName(..., True, True)` の「先頭列＝項目名」は、渡した範囲配列の**順序**でなく
    **シート上で左にある方**で決まる。項目列が値列より右だと、
    系列名・項目名・値の 3 つがそろって反転する。

  battery で名前が出た（`売上/chart` 値列の参照が C でなく D）── 売上表は
  `部門/担当/金額/月` で、「金額の棒グラフ」だと項目=月(D) > 値=金額(C)。

★★ 検分の入り方（棚の記録どおり）:
  **列の並びを入れ替えた 2 検体で同じ結果になること。**
  片方だけ通る直し方（例: 右のときだけ断る）はここで赤くなる。

★ LLM は使わない ── `InsertChart` を basrun で直接呼ぶ。測るのはヘルパの役割の付け方で、
  翻訳の質ではない（tests/test_bold_local.py と同じ作法）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

HELPERS_DIR = REPO / "src" / "ailine" / "helpers"

VALUES = [100, 200, 150, 120, 180]


def _rows(numeric_month: bool) -> list:
    """★ 月を**数値**にした形も持つ（battery の売上表は月が 4/5 の数値）。

    項目が文字なら、反転した"値"は文字列になり系列そのものが消える（目に見えて壊れる）。
    項目が数値なら、反転しても系列は在って**違う列を描く**（静かに間違う ── こちらが危ない）。
    """
    return [[(i + 3) if numeric_month else f"{i}月", v] for i, v in enumerate(VALUES, start=1)]


def _book(path: Path, category_first: bool, numeric_month: bool = False) -> tuple:
    """同じ中身で、列の並びだけ違う表を作る。戻り値: (項目列の記号, 値列の記号, 項目idx, 値idx)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    rows = _rows(numeric_month)
    if category_first:
        ws.append(["月", "金額"])
        for m, v in rows:
            ws.append([m, v])
        cat, val = ("A", "B"), (0, 1)
    else:
        ws.append(["金額", "月"])          # ★ 項目が値より右
        for m, v in rows:
            ws.append([v, m])
        cat, val = ("B", "A"), (1, 0)
    wb.save(path)
    wb.close()
    return cat[0], cat[1], val[0], val[1]


def _insert_chart(book: Path, workdir: Path, cat_idx: int, val_idx: int, kind: str = "bar"):
    _catalog, helper_files = ailine.load_helpers(HELPERS_DIR)
    assert helper_files, "helpers/*.bas が見つからない"
    code = (
        "Option VBASupport 1\nOption Explicit\n\n"
        "Sub Run(oDoc As Object)\n"
        f"    Call InsertChart(oDoc, 0, {cat_idx}, {val_idx}, \"{kind}\", {len(VALUES)})\n"
        "End Sub\n"
    )
    ok, err, raw = ailine.basrun_apply(book, code, workdir, helper_files)
    assert ok, f"basrun_apply が失敗した: {err}\n{(raw or '')[-500:]}"


@pytest.mark.local
@pytest.mark.parametrize("numeric_month", [False, True], ids=["月が文字", "月が数値"])
@pytest.mark.parametrize("category_first", [True, False],
                         ids=["項目が値の左", "項目が値の右"])
def test_chart_roles_do_not_depend_on_which_column_is_left(tmp_path, category_first,
                                                            numeric_month):
    """★★ 同じ依頼・同じ中身で、列の並びだけ変えても**同じ役割**で描けること。

    直す前の実測（2026-09-10）:
      項目が左 → 系列名 $C$1（金額）／項目 $B$2:$B$6（月）／値 $C$2:$C$6（金額）  正しい
      項目が右 → 系列名 $D$1（月）  ／項目 $C$2:$C$6（金額）／値 $D$2:$D$6（月）  3 つとも反転
    """
    book = tmp_path / "b.xlsx"
    cat_letter, val_letter, cat_idx, val_idx = _book(book, category_first, numeric_month)
    workdir = tmp_path / "work"
    workdir.mkdir()
    try:
        _insert_chart(book, workdir, cat_idx, val_idx)
        status, reason = ailine.check_chart_series(
            book, kind="bar", value_col_letter=val_letter, category_col_letter=cat_letter)
        assert status == "pass", (
            f"★ 項目が値の{'左' if category_first else '右'}で役割が崩れた: {reason}")
    finally:
        ailine._stop_office()


@pytest.mark.local
@pytest.mark.parametrize("kind", ["line", "pie"])
def test_other_kinds_also_keep_roles_when_category_is_right(tmp_path, kind):
    """★ 棒だけ直して折れ線・円が残る片配線を塞ぐ（種別は Diagram の差し替えで分岐する）。"""
    book = tmp_path / "b.xlsx"
    cat_letter, val_letter, cat_idx, val_idx = _book(book, category_first=False)
    workdir = tmp_path / "work"
    workdir.mkdir()
    try:
        _insert_chart(book, workdir, cat_idx, val_idx, kind=kind)
        status, reason = ailine.check_chart_series(
            book, kind=kind, value_col_letter=val_letter, category_col_letter=cat_letter)
        assert status == "pass", f"★ {kind}: 項目が右で役割が崩れた: {reason}"
    finally:
        ailine._stop_office()
