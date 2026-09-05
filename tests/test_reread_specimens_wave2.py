"""読み直しの塊に、**配線を通す検体**を足す（2026-09-05・第 2 波）。

★★ 出所（盲検の査定・所見⑨）: `_translate_and_dispatch` が 681 行あり、その 6 割が
  「読み直し」の `if` の積層だった。査定者の言葉:「バグを直すたびに同じ関数へ if を
  足す形で、著者自身が『1 つの関数に畳んで呼び出し側に持たせない』と書いた原則が、
  ここには適用されていない」。

★ 畳みたい。だが **`tests/test_reread_ledger.py` が分母を出していた** ──
  15 塊のうち、配線を通す検体を持つのは **3 塊だけ**。
  ★★ ゴールデンの無いコードを畳むのは、この repo が一番戒めている形だ
    （README「番人を作ってから割った」）。**順番は「検体 → 畳む」**。
  この束は、その分母を減らすための第 2 波。

★ 型は既に在るものを踏襲する（`test_swap_two_cells.py` の `..._without_the_llm`）:
  一段目の翻訳を**実測で起きた誤りの形**に差し替え、`--dry` で走らせる
  （LibreOffice も ollama も要らない ＝ CI で回る）。
★ 二段目翻訳（LLM）が要る塊は、**返る形だけ**を差し替える ── 呼ばれること自体も
  検体の一部なので、呼ばれなかったら落ちるようにする。
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
    ws.append(["品名", "数量", "金額", "備考"])
    for row in [["ボルト", 10, 12000, ""], ["ナット", 5, 4500, "確認済"],
                ["ワッシャ", 8, 30000, "確認済"]]:
        ws.append(row)
    p = tmp_path / "t.xlsx"
    wb.save(p)
    return p


def _gives_up(monkeypatch, about="よく分からない依頼"):
    """一段目が語彙外を返す回（実測で最も多い誤りの形）。"""
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1:
                        {"plan": [{"op": "OUT_OF_VOCAB", "about": about}]})


def _returns(monkeypatch, plan):
    monkeypatch.setattr(ailine, "translate_task",
                        lambda model, task, book_meta, temperature=0.1: {"plan": plan})


def _never_asks_the_llm(monkeypatch):
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda *a, **k: pytest.fail("機械で解けるのに LLM に聞いている"))


# --- ① 数値書式（★ 二段目を呼ばずに機械だけで解ける塊）----------------------

def test_the_number_format_reread_fires_without_the_llm(book, tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    _gives_up(monkeypatch)
    _never_asks_the_llm(monkeypatch)
    _rc, out = _run_main(["run", str(book), "金額に桁区切りを付けて", "--dry"], capsys)
    assert "『数値書式』として読み直しました" in out, out
    assert "金額" in out, out


def test_the_number_format_reread_stays_quiet_when_the_plan_is_already_right(
        book, tmp_path, monkeypatch, capsys):
    """★ 対で縛る ── 既に正しい計画が来ている回に、読み直しが割り込まないこと。"""
    _returns(monkeypatch, [{"op": "NUMBER_FORMAT", "args": {"col": "金額", "style": "thousands"}}])
    _isolate(monkeypatch, tmp_path)
    _never_asks_the_llm(monkeypatch)
    _rc, out = _run_main(["run", str(book), "金額に桁区切りを付けて", "--dry"], capsys)
    assert "として読み直しました" not in out, out


# --- ② 列抽出（★ 同上・機械だけ）--------------------------------------------

def test_the_column_extraction_reread_fires_without_the_llm(book, tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    _gives_up(monkeypatch)
    _never_asks_the_llm(monkeypatch)
    _rc, out = _run_main(["run", str(book), "品名と金額の列だけ抜き出して", "--dry"], capsys)
    assert "『列抽出』として読み直しました" in out, out


# --- ③ 列追加（★ 二段目が要る塊 ── 返る形だけ差し替える）--------------------

def test_the_add_column_reread_uses_the_second_pass(book, tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    _gives_up(monkeypatch)
    asked = []

    def fake_fixed(model, op, task, book_meta, **kw):
        asked.append(op)
        return {"op": op, "args": {"name": "区分"}}

    monkeypatch.setattr(ailine, "translate_task_fixed_op", fake_fixed)
    _rc, out = _run_main(["run", str(book), "区分という列を追加して", "--dry"], capsys)
    assert "『列追加』として読み直しました" in out, out
    assert "ADD_COLUMN" in asked, f"二段目を呼んでいない: {asked}"


# --- ④ 置き換え（★ 実測の事故形: 列を丸ごと潰しかけた）----------------------

def test_the_replace_reread_does_not_overwrite_the_whole_column(
        book, tmp_path, monkeypatch, capsys):
    """★ 出所（2026-08-27・Namakoo「置き換えができない」）: 一段目が
      SET_COLUMN_VALUE（列を丸ごと）を返し、空欄の行まで潰すところだった。"""
    _isolate(monkeypatch, tmp_path)
    _returns(monkeypatch, [{"op": "SET_COLUMN_VALUE", "args": {"col": "備考", "value": "済"}}])
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw: {"op": op, "args": {"col": "備考"}})
    _rc, out = _run_main(["run", str(book), "備考の『確認済』を全て『済』にして", "--dry"], capsys)
    assert "『置き換え』として読み直しました" in out, out
    assert "確認済" in out and "済" in out, out


# --- ⑤ 条件つき書換（★ 実測: 別シートを作って列を潰すところだった）----------

def test_the_conditional_write_reread_fires_on_the_split_plan(
        book, tmp_path, monkeypatch, capsys):
    """★ 出所（2026-08-30）: 「売上が1000以上の行の担当を『佐藤』に」が
      **抽出＋一括書換の 2 段**になり、別シートを作って列を丸ごと潰しかけた。"""
    _isolate(monkeypatch, tmp_path)
    _returns(monkeypatch, [{"op": "EXTRACT", "args": {}},
                           {"op": "SET_COLUMN_VALUE", "args": {"col": "備考", "value": "○"}}])
    monkeypatch.setattr(ailine, "translate_task_fixed_op",
                        lambda model, op, task, book_meta, **kw:
                        {"op": op, "args": {"col": "備考", "cond_col": "金額",
                                            "cmp": "ge", "cond_value": 10000, "value": "○"}})
    _rc, out = _run_main(["run", str(book), "金額が10000以上の行の備考に「○」を付けて", "--dry"], capsys)
    assert "『条件つき書換』として読み直しました" in out, out


# --- ⑥ 変異試験の代わり（★ 塊を無効化したら、上の検体が落ちること）----------

def test_these_specimens_actually_pass_through_the_blocks():
    """★ 「在るのに効かない」を防ぐ ── 検体が読み直しの**文言**を見ていること。

    ★ 2026-09-02 の教訓: 文字の順序しか見ない番人は `if False and …` で無効化しても
      緑のままだった。ここは実際に画面へ出た文字列で判定している（上の 5 本）ので、
      塊を殺せば必ず落ちる。この 1 本はその約束を明文化するだけ。
    """
    src = Path(__file__).read_text(encoding="utf-8")
    fired = src.count("として読み直しました' in out") + src.count("として読み直しました\" in out")
    assert fired >= 5, f"文言で確かめている検体が {fired} 本しかない"
