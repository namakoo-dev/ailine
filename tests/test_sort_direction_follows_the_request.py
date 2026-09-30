# 並べ替えの向きは依頼文が決める（2026-09-30・盲検 7 体目の致命 ①）。
#
# ★ 実測: 古い順に並んだ請求一覧で「請求日の新しい順に並べ替えて」
#   → 解釈「順:昇順」／（文書に変化は検出されなかった）／ ✓ 機械検証済み。
#   向き（order）を決めていたのは LLM だけで、事後条件は宣言↔実体しか見ない ──
#   依頼の項が欠けていたので、取り違えても ✓ になった。
#
# 契約:
#   ① 依頼文から向きが読めて LLM と食い違えば、依頼文の向きに置き換える
#   ② 置き換えたことは解釈行に出る（出典「依頼文: 『新しい順』」・警告ではない）
#   ③ 食い違わない／読めない時は何も変えない（出典も付けない）
#   ④ 打ち消された語（「新しい順ではなく」）は数えない
#   ⑤ 見出しが 1 行目でない表でも同じ

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core.sort_direction import read_direction  # noqa: E402

HEAD = ["請求番号", "請求日", "金額"]


def _book(tmp_path, header_row=1):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    for _ in range(header_row - 1):
        ws.append([])
    ws.append(HEAD)
    ws.append(["INV-0901", "2026-09-01", 1000])
    ws.append(["INV-0902", "2026-09-05", 3000])
    ws.append(["INV-0903", "2026-09-09", 2000])
    wb.save(p)
    return p


def _meta(path, header_row=1):
    return {"sheets": ["請求"], "headers": {"請求": list(HEAD)},
            "header_rows": {"請求": header_row}, "path": str(path)}


def _sort(tmp_path, task, order, col="請求日", header_row=1):
    ok, r, _i, err = ailine.verify_dsl_args(
        "SORT", {"col": col, "order": order},
        _meta(_book(tmp_path, header_row), header_row), task=task)
    assert ok, err
    return r


# --- ①② 食い違えば依頼文が勝ち、解釈行に出る ------------------------------------------

def test_newest_first_overrides_an_ascending_answer(tmp_path):
    r = _sort(tmp_path, "請求日の新しい順に並べ替えて", "asc")
    assert r["order"] == "desc"
    line = ailine.format_confirmation_line("SORT", r, set())
    assert "順:降順（依頼文: 『新しい順』）" in line, line
    assert not r.get("_warnings"), r.get("_warnings")   # 警告ではなく解釈行の側


def test_oldest_first_overrides_a_descending_answer(tmp_path):
    r = _sort(tmp_path, "古い順に並べ替えて", "desc")
    assert r["order"] == "asc"
    assert r["_sources"]["order"] == "依頼文: 『古い順』"


# --- ③ 食い違わない／読めない時は何もしない -------------------------------------------

def test_agreeing_answer_is_left_alone(tmp_path):
    r = _sort(tmp_path, "金額の大きい順に並べ替えて", "desc", col="金額")
    assert r["order"] == "desc"
    assert "order" not in (r.get("_sources") or {})


@pytest.mark.parametrize("order", ["asc", "desc"])
def test_no_direction_word_changes_nothing(tmp_path, order):
    r = _sort(tmp_path, "請求日で並べ替えて", order)
    assert r["order"] == order
    assert "order" not in (r.get("_sources") or {})


def test_both_directions_named_changes_nothing(tmp_path):
    r = _sort(tmp_path, "昇順・降順どちらでもいいので請求日で並べ替えて", "asc")
    assert r["order"] == "asc"


# --- ④ 打ち消された語は数えない -------------------------------------------------------

def test_negated_word_is_not_taken(tmp_path):
    r = _sort(tmp_path, "新しい順ではなく古い順に並べ替えて", "desc")
    assert r["order"] == "asc"


def test_word_inside_quotes_is_not_taken(tmp_path):
    # ★ 引用符の中は値・名前であって向きの指示ではない（『昇順番号』のような列名）
    r = _sort(tmp_path, "請求日で並べ替えて、備考に「昇順」と書いて", "desc")
    assert r["order"] == "desc"
    assert "order" not in (r.get("_sources") or {})


@pytest.mark.parametrize("task, expected", [
    ("新しい順じゃなくて古い順に", "asc"),
    ("新しい順にではなく古い順に", "asc"),
    ("新しい順ではなく", None),          # 打ち消しだけ ── 向きの語が残らない
    ("請求日の新しい順", "desc"),
    ("金額が大きい順", "desc"),
    ("昇順と降順", None),
    ("日付で並べ替えて", None),
    ("", None),
])
def test_read_direction(task, expected):
    got = read_direction(task)
    assert (got[0] if got else None) == expected, got


# --- ⑤ 見出しが 1 行目でない表 --------------------------------------------------------

def test_works_when_the_header_is_not_on_row_one(tmp_path):
    r = _sort(tmp_path, "請求日の新しい順に並べ替えて", "asc", header_row=4)
    assert r["order"] == "desc"
    assert r["_sources"]["order"] == "依頼文: 『新しい順』"
