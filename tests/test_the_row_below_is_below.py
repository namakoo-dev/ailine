# -*- coding: utf-8 -*-
"""「〜の行の下に … 追加して」が、通って、しかも**下**に入る（2026-09-09・盲検 D）。

★★ 出所（総務の立場で打たれた、いちばん自然な言い方）:

    依頼   「北斗精機**の行の下に**、取引先「西村工業」の行を追加して、
             項目は事務机、件数は 2、単価は 15000 にして」
    出力   **？ 1 回の依頼で行を 2 回足そうとしています**（断り）

  1 つの指摘から欠陥が **3 つ**出た:

    ① モデルは INSERT_ROWS（空行）＋ADD_ROW（値）の 2 段で返す（3/3 で同じ）。
       人が手でやる手順どおりで、モデルは間違っていない。ADD_ROW はそれ自体が
       行を作るので、機械の「二重宣言」判定も正しい ── **畳めば済む**。
    ② 畳む時に**モデルの位置**（at）を持ち込んだら、北斗精機は 6 行目なのに
       at=8 で、「下に」が**上**に入った。★ 位置は実表から解き直す（A' 原則）。
    ③ ★ そのうえで、なお**上**に入った ── 「北斗精機**の行**の下に」で掴むのが
       『北斗精機の行』になり、実表に無いので落ちて、後段の「<X>の行」規則が
       **その行そのもの**を返していた。「の下に」が丸ごと消えていた。

  ★ ③ は ① を直した**副作用で見えた** ── 直さなければ、断られたまま気づかなかった。
"""
from __future__ import annotations

import openpyxl
import pytest

import ailine
from ailine_core.fold_insert_add import fold_insert_then_add

ROWS = [["取引先", "項目", "件数", "単価"],
        ["丸和物流", "配送", 12, 4800],
        ["近江スチール", "鋼材", 5, 12000],
        ["北斗精機", "精密部品", 3, 38000],
        ["みどり建設", "内装", 9, 7200]]


@pytest.fixture()
def meta(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    return {"sheets": ["請求"], "headers": {"請求": list(ROWS[0])},
            "header_rows": {"請求": 1}, "path": str(p)}


# --- ① 畳み（純ロジック）-----------------------------------------------------

def test_insert_then_add_is_folded_into_one_step():
    plan = [{"op": "INSERT_ROWS", "args": {"at": 8, "count": 1}},
            {"op": "ADD_ROW", "args": {"at": 9, "values": {"取引先": "西村工業"}}}]
    got, note = fold_insert_then_add(plan)
    assert len(got) == 1 and got[0]["op"] == "ADD_ROW", got
    assert note and "1 段にまとめました" in note
    # ★★ モデルの位置は持ち込まない（実表から解き直す）── 持ち込むと「下」が「上」になる
    assert "at" not in got[0]["args"], got[0]["args"]
    assert got[0]["args"]["values"] == {"取引先": "西村工業"}


@pytest.mark.parametrize("plan", [
    # ★ 値が無い ── 本当に空行が欲しい回（畳まない）
    [{"op": "INSERT_ROWS", "args": {"at": 8, "count": 1}},
     {"op": "ADD_ROW", "args": {"at": 9, "values": {}}}],
    # ★ 3 行空けて 1 行だけ埋める ── 釣り合わないので畳まない
    [{"op": "INSERT_ROWS", "args": {"at": 8, "count": 3}},
     {"op": "ADD_ROW", "args": {"at": 9, "values": {"取引先": "西村工業"}}}],
    # ★ 空けた行と関係ない行へ入れる
    [{"op": "INSERT_ROWS", "args": {"at": 8, "count": 1}},
     {"op": "ADD_ROW", "args": {"at": 20, "values": {"取引先": "西村工業"}}}],
    # ★ 組み合わせが違う
    [{"op": "SORT", "args": {"col": "単価", "order": "desc"}},
     {"op": "ADD_ROW", "args": {"at": 9, "values": {"取引先": "西村工業"}}}],
])
def test_what_should_not_be_folded_is_left_alone(plan):
    got, note = fold_insert_then_add(plan)
    assert note is None and got == plan


# --- ② 位置（実表から解き直す）----------------------------------------------

@pytest.mark.parametrize("task, want, why", [
    # ★ 実測した事故そのもの ── 「の行の下に」で位置語が消えていた
    ("北斗精機の行の下に、取引先「西村工業」の行を追加して", 5, "の下＝5行目"),
    ("北斗精機の下に西村工業の行を追加して", 5, "の下＝5行目"),
    ("北斗精機の行の上に1行足して", 4, "の上＝4行目"),
    # ★ 対の試験: 「<X>の行」だけの回は、その行そのもの（位置語は無い）
    ("みどり建設の行を削除して", 5, "の行＝5行目"),
])
def test_the_structure_word_does_not_eat_the_position(meta, task, want, why):
    got, note = ailine.resolve_row_anchor(task, meta, "請求", 1)
    assert got == want, (got, note)
    assert why in (note or ""), note


# --- ③ 実機（通って、しかも下に入る）----------------------------------------

@pytest.mark.local
def test_it_really_lands_below_on_real_libreoffice(tmp_path):
    import os
    import subprocess
    import sys
    from pathlib import Path

    repo = Path(ailine.__file__).resolve().parents[2]
    book = tmp_path / "seikyu.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    for r in ROWS:
        ws.append(r)
    wb.save(book)

    r = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(book),
         "北斗精機の行の下に、取引先「西村工業」の行を追加して、項目は事務机、"
         "件数は2、単価は15000にして", "--copy", "--sheet", "請求", "--timeout", "300"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900,
        env={**os.environ, "PYTHONPATH": str(repo / "src")})
    assert r.returncode == 0, r.stdout[-800:]
    assert "行を 2 回足そう" not in r.stdout, r.stdout[-800:]
    out = openpyxl.load_workbook(book.with_name(book.stem + ".out.xlsx"))["請求"]
    names = [out.cell(i, 1).value for i in range(1, 7)]
    assert names == ["取引先", "丸和物流", "近江スチール", "北斗精機",
                     "西村工業", "みどり建設"], names
