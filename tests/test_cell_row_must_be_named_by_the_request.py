# 1 セルの上書きで、行は依頼文が名指ししたものに限る（2026-10-01・依頼の項の台帳で B だった項目）。
#
# ★ 形: 行の名前（row）は「実表に在るか」しか見ていなかった。依頼が行を言っていないのに
#   LLM が実在する別の行の名前を返すと、その行を上書きして ✓ が出る（宣言↔実体は合っている）。
#
# 契約:
#   ① 依頼文がその行を名指ししていれば（その行のどれかの値・LLM が挙げた名前が照合できる）黙る
#   ② 名指しが無ければ ⚠（_warnings）で開示する ── _warnings は ✓→△ の降格に数えられる
#   ③ 別の行の名前の断片（『山田』工業）を、この行（山田商事）の証拠にしない
#   ④ 行番号で指した依頼（row_number）の道は変えない
#   ⑤ 見出しが 1 行目でない表でも同じ

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import argcheck  # noqa: E402

HEAD = ["取引先", "伝票", "売上", "担当"]
DATA = [["山田商事", "INV-1", 1000, "佐藤"],
        ["山田工業", "INV-2", 2000, "鈴木"],
        ["丸和物流", "INV-3", 3000, "高橋"]]


def _book(tmp_path, header_row=1):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "取引"
    for i in range(header_row - 1):
        ws.append(["取引一覧"] if i == 0 else [])
    ws.append(HEAD)
    for r in DATA:
        ws.append(r)
    wb.save(p)
    return p


def _meta(path, header_row=1):
    return {"sheets": ["取引"], "headers": {"取引": list(HEAD)},
            "header_rows": {"取引": header_row}, "path": str(path)}


def _cell(tmp_path, task, args, header_row=1):
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", dict(args), _meta(_book(tmp_path, header_row), header_row), task=task)
    assert ok, err
    return r


def _row_warnings(r):
    return [w for w in r.get("_warnings", []) if "行を指す語が見当たりません" in w]


# --- ① 名指しされていれば黙る ------------------------------------------------------------

def test_named_row_is_silent(tmp_path):
    r = _cell(tmp_path, "丸和物流の担当を『田中』にして",
              {"row": "丸和物流", "col": "担当", "value": "田中"})
    assert r["_row_index"] == 4 and not _row_warnings(r)


def test_row_named_by_another_cell_of_it_is_silent(tmp_path):
    """伝票番号で指した行に、LLM が取引先名を返した ── 同じ行なので名指しされている。"""
    r = _cell(tmp_path, "INV-3の担当を『田中』にして",
              {"row": "丸和物流", "col": "担当", "value": "田中"})
    assert r["_row_index"] == 4 and not _row_warnings(r)


# --- ② 名指しが無ければ ⚠ ---------------------------------------------------------------

def test_unnamed_row_is_disclosed(tmp_path):
    r = _cell(tmp_path, "担当を『田中』にして",
              {"row": "丸和物流", "col": "担当", "value": "田中"})
    assert len(_row_warnings(r)) == 1
    assert "『丸和物流』" in _row_warnings(r)[0] and "4行目" in _row_warnings(r)[0]


def test_a_different_row_named_by_the_request_is_disclosed(tmp_path):
    r = _cell(tmp_path, "丸和物流の担当を『田中』にして",
              {"row": "山田商事", "col": "担当", "value": "田中"})
    assert _row_warnings(r)


# --- ③ 別の行の名前の断片は証拠にしない ----------------------------------------------------

def test_a_fragment_of_another_row_is_not_evidence(tmp_path):
    r = _cell(tmp_path, "山田工業の担当を『田中』にして",
              {"row": "山田商事", "col": "担当", "value": "田中"})
    assert _row_warnings(r)
    r2 = _cell(tmp_path, "山田工業の担当を『田中』にして",
               {"row": "山田工業", "col": "担当", "value": "田中"})
    assert not _row_warnings(r2)


# --- ④ 行番号の道は変えない --------------------------------------------------------------

def test_row_number_path_is_untouched(tmp_path):
    r = _cell(tmp_path, "3行目の担当を『田中』にして",
              {"row_number": 3, "col": "担当", "value": "田中"})
    assert r["_row_index"] == 3 and not _row_warnings(r)


# --- ⑤ 見出しが 1 行目でない表 -----------------------------------------------------------

def test_header_not_on_the_first_row(tmp_path):
    r = _cell(tmp_path, "担当を『田中』にして",
              {"row": "丸和物流", "col": "担当", "value": "田中"}, header_row=3)
    assert r["_row_index"] == 6 and _row_warnings(r)
    r2 = _cell(tmp_path, "丸和物流の担当を『田中』にして",
               {"row": "丸和物流", "col": "担当", "value": "田中"}, header_row=3)
    assert not _row_warnings(r2)


# --- ② の帰結: _warnings は ✓ を降ろす材料に数えられている -------------------------------

def test_warnings_are_what_the_check_counts():
    """★ 帰結の配線: resolved["_warnings"] は ✓→△ の降格に数えられる（在っても鳴らない、を避ける）。
    ★ 場所でなく名前で引く（単発の入口 cmd_run_dsl が数えている）。"""
    import inspect
    assert 'len(resolved.get("_warnings", []))' in inspect.getsource(ailine.cmd_run_dsl)


# --- 変異: 配線を外すと黙る ---------------------------------------------------------------

def test_without_the_wiring_the_guess_would_pass_silently(tmp_path, monkeypatch):
    monkeypatch.setattr(argcheck, "_row_named_by_the_request", lambda *a, **k: True)
    r = _cell(tmp_path, "担当を『田中』にして",
              {"row": "丸和物流", "col": "担当", "value": "田中"})
    assert not _row_warnings(r)      # ← 直す前の挙動（嘘の ✓ の材料）
