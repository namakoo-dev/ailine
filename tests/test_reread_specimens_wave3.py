"""読み直しの塊に、配線を通す検体を足す（2026-09-05・第 3 波 ── 残り 7 塊）。

★★ なぜ第 3 波まで書くか: `_translate_and_dispatch` の 681 行を**畳みたい**が、
  `tests/test_reread_ledger.py` が出している分母は 15 塊。第 2 波で 8 塊まで来た。
  **ゴールデンの無いコードを畳むのは、この repo が一番戒めている形**なので、
  残りに検体が付くまで畳まない（README「番人を作ってから割った」）。

★ 発火形は**実物に当てて確かめてから**書いた（今日の教訓「検体は窒息点と治具まで
  含めて仮説」）。実際、最初に書こうとした形は 2 つとも空振りした:
    ・`removal_reading("3行目を削除して")` → None（★ 行番号ではなく**名前**で解く器官だった）
    ・`add_row_values_from_request(..., {"品名": "ねじ"})` → {}（★ 位置の目印になる値は
      篩で落ちる ── 「置く物の名前」だけでは値にならない）

★ どれも `--dry` で走る（LibreOffice も ollama も要らない ＝ CI で回る）。
"""
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import ailine  # noqa: E402
from test_golden_transcripts import _isolate, _run_main  # noqa: E402


@pytest.fixture
def book(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["品名", "担当", "金額", "備考"])
    for row in [["ボルト", "田中", 12000, ""], ["ナット", "鈴木", 4500, ""],
                ["ワッシャ", "佐藤", 30000, ""]]:
        ws.append(row)
    p = tmp_path / "t.xlsx"
    wb.save(p)
    return p


def _returns(monkeypatch, plan):
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"plan": plan})


def _gives_up(monkeypatch):
    _returns(monkeypatch, [{"op": "OUT_OF_VOCAB", "about": "よく分からない依頼"}])


def _never_asks_the_llm(monkeypatch):
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda *a, **k: pytest.fail("機械で解けるのに LLM に聞いている"))


# --- 塊 2: 一括書換 → 1 セル書換（★ 実測で最も高くついた事故の形）--------------

def test_a_whole_column_write_is_narrowed_to_one_cell(book, tmp_path, monkeypatch, capsys):
    """★★ 出所（2026-08-28・Namakoo が実測・「今日いちばん悪い形」）:
      「7行目の担当を『佐藤』に」で**担当列が全行『佐藤』になり ✓ が出た**。
      一括書換の契約としては ✓ は正しいが、依頼は 1 行だった ──
      三項（依頼・宣言・実体）のうち**依頼を見ていなかった**。
    """
    _isolate(monkeypatch, tmp_path)
    _returns(monkeypatch, [{"op": "SET_COLUMN_VALUE", "args": {"col": "備考", "value": "済"}}])
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw:
                        {"op": op, "args": {"col": "備考", "value": "済"}})
    # ★ 検体を実物に当てて分かったこと: 「3行目の備考を…」は**別の塊**（列まで名指し
    #   できる回）が先に拾う。ここが狙う塊は「行は分かるが列が依頼文に無い」回なので、
    #   `resolve_cell_target_from_task` が解けない形にする（実測で確かめた）。
    _rc, out = _run_main(["run", str(book), "3行目を「済」にして", "--dry"], capsys)
    assert "『一括書換』でなく『1セル書換』として読み直しました" in out, out
    assert "3行目" in out, out


def test_when_it_cannot_pick_one_cell_it_refuses_instead_of_writing_the_column(
        book, tmp_path, monkeypatch, capsys):
    """★ 対の側 ── 1 セルに落とせなかった回は、**列ぜんぶを書かずに断る**。

    ★ ここは「本物の 2 択」が残っている唯一の場所（1 セルか、列ぜんぶか）。
      行き止まりにせず、選べる形で返すこと（モーダルは使わない ── 画面を止める）。
    """
    _isolate(monkeypatch, tmp_path)
    _returns(monkeypatch, [{"op": "SET_COLUMN_VALUE", "args": {"col": "備考", "value": "済"}}])
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw: {"op": op, "args": {}})
    _rc, out = _run_main(["run", str(book), "ナットの備考を「済」にして", "--dry"], capsys)
    assert "決められませんでした" in out or "読み直しました" in out, out
    assert "列全体は勝手に書き換えません" in out or "1セル書換" in out, out


# --- 塊 3: 書式の対象を 1 セルに絞る（★ 台帳が「黙る」と数えていた塊）----------

def test_formatting_a_named_value_stays_in_one_cell(book, tmp_path, monkeypatch, capsys):
    """★ 台帳（test_reread_ledger.py）はこの塊を「文言を出さない」側に数えていたが、
      実際は出している ── 文面が『{名前}』の**1 セル**として…で、
      台帳の正規表現（『…』として読み直しました）が**間に挟まる文字で外れていた**。
      ★ 「黙る塊が何本か」も、機械の数え方次第で変わる。
    """
    _isolate(monkeypatch, tmp_path)
    _gives_up(monkeypatch)
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw: {"op": op, "args": {}})
    _rc, out = _run_main(["run", str(book), "ボルトのセルを太字にして", "--dry"], capsys)
    assert "1 セル" in out and "読み直しました" in out, out
    assert "列ぜんぶには広げません" in out, out


# --- 塊 13: 行削除（★ 「除く」の 2 通りの読みを取り違えない）-------------------

def test_the_row_removal_reread_resolves_the_row_by_name(book, tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    _gives_up(monkeypatch)
    _never_asks_the_llm(monkeypatch)
    _rc, out = _run_main(["run", str(book), "ナットの行を消して", "--dry"], capsys)
    assert "『行削除』として読み直しました" in out, out
    assert "3行目" in out, out


def test_it_does_not_delete_when_the_task_says_except(book, tmp_path, monkeypatch, capsys):
    """★ 対で縛る ── 「〜以外」は**残したい**側。①に化けさせると取り返しがつかない。"""
    _isolate(monkeypatch, tmp_path)
    _gives_up(monkeypatch)
    _rc, out = _run_main(["run", str(book), "ナット以外の行を残して", "--dry"], capsys)
    assert "『行削除』として読み直しました" not in out, out


# --- 塊 15: 行挿入 → 行追加（★ 空行を挿すだけで終わらせない）-------------------

def test_an_empty_row_insert_becomes_a_row_with_values(book, tmp_path, monkeypatch, capsys):
    """★★ 出所（2026-08-27・Namakoo が実測）:「みかんとぶどうの間に梨を追加して。
      売上は600 原価は300」が**空行 1 本の挿入**になった。op の取り違え。
    ★ 値は機械が依頼文から決める（LLM の出した値は篩にかけるだけ）── 実測で
      「位置の目印が値になり・置く物が別の列に入り・未設定がでっち上げられる」が同時に起きた。
    """
    _isolate(monkeypatch, tmp_path)
    _returns(monkeypatch, [{"op": "INSERT_ROWS", "args": {"at": 4, "count": 1}}])
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw:
                        {"op": op, "args": {"values": {"品名": "ねじ", "担当": "山田",
                                                        "金額": 800}}})
    _rc, out = _run_main(["run", str(book),
                          "ねじの行を一番下に追加して。担当は山田 金額は800", "--dry"], capsys)
    assert "『行挿入』でなく『行追加』として読み直しました" in out, out


def test_a_request_that_really_wants_a_blank_row_is_left_alone(
        book, tmp_path, monkeypatch, capsys):
    """★ 誤爆側 ── 「空行が欲しい」と言っている依頼には触らない。"""
    _isolate(monkeypatch, tmp_path)
    _returns(monkeypatch, [{"op": "INSERT_ROWS", "args": {"at": 3, "count": 1}}])
    _rc, out = _run_main(["run", str(book), "3行目に空行を挿入して", "--dry"], capsys)
    assert "『行追加』として読み直しました" not in out, out


# --- 塊 10〜12: 入れ替えの 3 本目（列の入れ替え）------------------------------

def test_the_column_swap_reread_fires_without_the_llm(book, tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    _gives_up(monkeypatch)
    # ★ 列の入れ替えは機械だけでは解けない（どの 2 列かは二段目が返す）── 返る形を差し替える。
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw:
                        {"op": op, "args": {"a": "品名", "b": "担当"}})
    _rc, out = _run_main(["run", str(book), "品名の列と担当の列を入れ替えて", "--dry"], capsys)
    assert "読み直しました" in out, out
    assert "入れ替え" in out, out
