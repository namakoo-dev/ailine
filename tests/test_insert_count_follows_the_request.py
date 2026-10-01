# 挿入する行の数は依頼文から機械が決める（2026-10-01・依頼の項の台帳で D だった項目）。
#
# ★ 形: 件数（count）を LLM だけが決めていた。「5行目に1行挿入して」に count=3 が返ると
#   3 行入り、事後条件は「宣言どおり 3 行ずれた」を確かめて ✓ を出す（削除の件数と同じ穴）。
#
# 契約:
#   ① 依頼文の数（「5行目に3行挿入」「2行空けて」「5〜7行目に」「5行目に空行を」＝1）が勝つ
#   ② 一致していれば何も足さない（出典も出さない）
#   ③ 読めない形（数の言い方が漢字・「数行」・始まりの行が at と違う）は何も変えない
#   ④ 数字も数の言い方も無ければ 1
#   ⑤ 見出しが 1 行目でない表でも同じ
#   ⑥ 削除の件数の読み方は変えない（裸の「3行」は削除では読まない）

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import argcheck  # noqa: E402
from ailine_core.anchor import row_count_in_task  # noqa: E402

HEAD = ["商品", "数量", "金額"]
DATA = [["りんご", 3, 300], ["みかん", 5, 500], ["ぶどう", 2, 700], ["もも", 1, 900]]


def _book(tmp_path, header_row=1):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for i in range(header_row - 1):
        ws.append(["売上一覧"] if i == 0 else [])
    ws.append(HEAD)
    for r in DATA:
        ws.append(r)
    wb.save(p)
    return p


def _meta(path, header_row=1):
    return {"sheets": ["売上"], "headers": {"売上": list(HEAD)},
            "header_rows": {"売上": header_row}, "path": str(path)}


def _insert(tmp_path, task, args, header_row=1):
    ok, r, inferred, err = ailine.verify_dsl_args(
        "INSERT_ROWS", dict(args), _meta(_book(tmp_path, header_row), header_row), task=task)
    assert ok, err
    return r, inferred


# --- ① 食い違えば依頼文が勝つ -----------------------------------------------------------

def test_the_number_in_the_request_wins(tmp_path):
    r, inferred = _insert(tmp_path, "3行目に2行挿入して", {"at": 3, "count": 1})
    assert r["count"] == 2
    assert r["_sources"]["count"] == "依頼文: 『2行』"
    assert "count" not in inferred


def test_blank_rows_by_a_bare_count(tmp_path):
    r, _i = _insert(tmp_path, "2行空けて", {"at": 3, "count": 1})
    assert r["count"] == 2


def test_a_row_range_is_a_count(tmp_path):
    r, _i = _insert(tmp_path, "3〜4行目に空行を入れて", {"at": 3, "count": 1})
    assert r["count"] == 2


def test_one_row_number_is_one_row(tmp_path):
    r, _i = _insert(tmp_path, "3行目に空行を入れて", {"at": 3, "count": 3})
    assert r["count"] == 1 and "count" in r["_sources"]


def test_the_count_is_shown_on_the_interpretation_line(tmp_path):
    r, inferred = _insert(tmp_path, "3行目に2行挿入して", {"at": 3, "count": 1})
    line = ailine.format_confirmation_line("INSERT_ROWS", r, inferred)
    assert "依頼文: 『2行』" in line


# --- ② 一致していれば黙る ----------------------------------------------------------------

def test_agreement_adds_nothing(tmp_path):
    r, _i = _insert(tmp_path, "3行目に2行挿入して", {"at": 3, "count": 2})
    assert r["count"] == 2 and "count" not in r.get("_sources", {})


# --- ③ 読めない形は何も変えない -----------------------------------------------------------

def test_a_count_in_kanji_is_not_read_as_one(tmp_path):
    r, _i = _insert(tmp_path, "みかんの下に空行を三行入れて", {"at": 4, "count": 3})
    assert r["count"] == 3 and "count" not in r.get("_sources", {})


def test_vague_counts_are_not_read(tmp_path):
    r, _i = _insert(tmp_path, "みかんの下に空行を数行入れて", {"at": 4, "count": 2})
    assert r["count"] == 2 and "count" not in r.get("_sources", {})


def test_a_start_that_differs_from_at_is_left_alone(tmp_path):
    """「2行目の下に」は at=3。始まり（2）が at と違う時は件数も決めない。"""
    r, _i = _insert(tmp_path, "2行目の下に3行挿入して", {"at": 3, "count": 1})
    assert r["count"] == 1 and "count" not in r.get("_sources", {})


def test_quoted_numbers_are_values(tmp_path):
    """引用符の中の「2行」は値 ── 件数として読まない（外に数が無いので 1 行）。"""
    r, _i = _insert(tmp_path, "「2行」の上に空行を入れて", {"at": 3, "count": 2})
    assert r["count"] == 1 and "2行" not in r["_sources"]["count"]


# --- ④ 数の無い依頼は 1 行 ------------------------------------------------------------------

def test_no_number_means_one_row(tmp_path):
    r, _i = _insert(tmp_path, "みかんの下に空行を入れて", {"at": 4, "count": 3})
    assert r["count"] == 1
    assert "指定が無い" in r["_sources"]["count"]


# --- ⑤ 見出しが 1 行目でない表 -------------------------------------------------------------

def test_header_not_on_the_first_row(tmp_path):
    r, _i = _insert(tmp_path, "5行目に2行挿入して", {"at": 5, "count": 1}, header_row=3)
    assert r["at"] == 5 and r["count"] == 2


# --- ⑥ 削除の読み方は変えない ---------------------------------------------------------------

def test_the_delete_reading_is_unchanged():
    assert row_count_in_task("3行削除して") is None
    assert row_count_in_task("5行目に3行挿入して") is None
    assert row_count_in_task("5行目に3行挿入して", argcheck._INSERT_COUNT_VERBS) == (3, 5, "3行")


# --- 変異: 配線を外すと LLM の数が黙って通る ---------------------------------------------

def test_without_the_wiring_the_llm_count_passes(tmp_path, monkeypatch):
    monkeypatch.setattr(argcheck, "_insert_count_from_request", lambda *a, **k: None)
    r, _i = _insert(tmp_path, "3行目に2行挿入して", {"at": 3, "count": 1})
    assert r["count"] == 1        # ← 直す前の挙動（嘘の ✓ の材料）
