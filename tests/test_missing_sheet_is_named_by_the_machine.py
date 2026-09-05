"""無いシート名は、**機械が**名指しして断ること（2026-09-05）。

★★ 出所（Namakoo が画面の文に引っかかった）: 1 枚しかないブックに
    「売上シートの金額を並べ替えて」と頼むと、道具はこう返していた ──

        ？ 売上シートとは何シートですか？

  意味を成していない。しかも **4 回に 1 回しか出ず**、残りは「照合できませんでした」で
  **揺れていた** ── その問いは LLM が作文した CLARIFY の question だったから。
★ 機械は答えを持っている（**シート名の一覧**）。今日 5 回目の
  「知識は在るのに機械が使っていない」。持っている事実で言い直させた:

        ？ このブックに『売上』というシートはありません（あるシート: 在庫）
          シート名を確かめるか、--sheet で明示してください

★ 誤爆（**在るのに「ありません」**／**指示語を名前と読む**）の方がこわいので、
  ②③ を①と**対で**縛る。★ ②③ は番人を書く前の 24 形の測定で、実際に
  **自分の実装から 2 件出た**（「別のシート」→『別の』／「集計シートを作って」）。
"""
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402
from ailine_core.target_sheet import (  # noqa: E402
    render_missing_sheet_refusal, sheet_named_but_missing)

ONE = ["在庫"]
TWO = ["注文一覧", "商品マスタ"]


# --- ① 無い名前は名指しで返る ----------------------------------------------

@pytest.mark.parametrize("task, sheets, want", [
    ("売上シートの金額を並べ替えて", ONE, "売上"),
    ("売上一覧シートを見て", ONE, "売上一覧"),
    ("「売上」シートの金額を並べ替えて", ONE, "売上"),
    ("Sheet1シートを太字に", ONE, "Sheet1"),
    ("商品のシートを直して", ONE, "商品"),
    ("売上シートに列を追加して", ONE, "売上"),
])
def test_a_sheet_that_is_not_there_is_named(task, sheets, want):
    assert sheet_named_but_missing(task, sheets) == want


# --- ② 在るもの・指示語・序数には黙る（★ こちらの誤爆が一番こわい）----------

@pytest.mark.parametrize("task, sheets", [
    ("在庫シートを並べ替えて", ONE),
    ("在庫のシートを並べ替えて", ONE),
    ("注文一覧シートの金額を合計して", TWO),
    ("商品表から商品名を注文シートに写して", TWO),   # ★ 貪欲マッチの罠
    ("このシートを並べ替えて", ONE),
    ("この表のシートを直して", ONE),                  # ★ 指示語＋の（前方一致で落とす）
    ("別のシートに書き出して", ONE),                  # ★ 実際に『別の』と名指ししていた
    ("元のシートに戻して", ONE),
    ("他のシートも同じに", ONE),
    ("現在のシートを太字に", ONE),
    ("次のシートに移して", ONE),
    ("新しいシートを作って", ONE),
    ("同じシートに追記して", ONE),
    ("各シートに罫線を引いて", ONE),
    ("全部のシートに罫線", ONE),
    ("2枚目のシートを並べ替えて", ONE),               # ★ 序数は既存の判定に譲る
    ("シートを並べ替えて", ONE),
    ("金額を並べ替えて", ONE),
    ("売上シートの金額を並べ替えて", []),             # ★ 一覧が取れなければ黙る
])
def test_it_stays_quiet_when_the_name_is_not_a_missing_sheet(task, sheets):
    assert sheet_named_but_missing(task, sheets) is None


# --- ③ これから作る名前に「ありません」と言わない（★ でたらめの側）----------

@pytest.mark.parametrize("task", [
    "集計シートを作って", "集計シートを追加して", "集計シートを新しく作って",
])
def test_a_sheet_the_user_is_creating_is_not_called_missing(task):
    assert sheet_named_but_missing(task, ONE) is None


def test_but_adding_a_column_to_a_missing_sheet_is_still_named():
    """★ ③ と対で縛る ── 「追加」の目的語がシートでなければ従来どおり名指しする。"""
    assert sheet_named_but_missing("売上シートに列を追加して", ONE) == "売上"


# --- ④ 断り文は「あるシート」を実際に名指しする ----------------------------

def test_the_refusal_names_the_sheets_that_do_exist():
    said = chr(10).join(render_missing_sheet_refusal("売上", TWO))
    assert "『売上』" in said and "ありません" in said
    for s in TWO:
        assert s in said, said
    assert "--sheet" in said


# --- ⑤ 配線（★ 純ロジックだけでは守れないことを今日学んだ）------------------

def test_the_gate_asks_the_real_book_for_its_sheets():
    src = Path(ailine.__file__).read_text(encoding="utf-8")
    gate = src.split("def cmd_refuse_vocab_miss")[1].split(chr(10) + "def ")[0]
    assert "sheet_named_but_missing(" in gate, "門に配線されていない"
    call = gate.split("sheet_named_but_missing(")[1][:80]
    assert "a.task" in call and "_sheets_now" in call, f"実表を渡していない: {call[:60]}"
    assert "build_book_meta" in gate, "シート名の出所が実ブックでない"


@pytest.mark.local
def test_the_machine_answers_instead_of_asking_what_a_sheet_is(tmp_path):
    """★ 実機 ── 出所の再現。**LLM の作文**が出ないことまで見る。"""
    import openpyxl
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "在庫"
    ws.append(["品名", "金額"]); ws.append(["机", 12000]); ws.append(["椅子", 8000])
    src = tmp_path / "in.xlsx"; wb.save(src)
    repo = Path(ailine.__file__).resolve().parents[2]
    env = {**os.environ, "PYTHONPATH": str(repo / "src")}
    got = subprocess.run([sys.executable, "-m", "ailine", "run", str(src),
                          "売上シートの金額を並べ替えて", "--copy", "--timeout", "150"],
                         capture_output=True, text=True, encoding="utf-8",
                         errors="replace", cwd=str(repo), env=env)
    assert got.returncode == 3, got.stdout[-400:]
    assert "『売上』というシートはありません" in got.stdout, got.stdout[-400:]
    assert "在庫" in got.stdout, got.stdout[-400:]
    assert not re.search(r"とは何シート", got.stdout), got.stdout[-400:]
