# セル結合の範囲は依頼文から・消える値は先に聞く（2026-10-01）。
#
# ★ 形 1（依頼の項の台帳で D だった項目）: 範囲（range）は「A1:C1 の形か」しか見ていなかった。
#   地図の実例:「ナットの右に東棟」が A1:D1 という作られた範囲になり、事後条件は
#   「宣言の範囲が結合されたか」を確かめて ✓ を出していた。
# ★ 形 2（実機で確かめた）: LibreOffice は結合すると左上以外の値を消す。
#   「A1からC1を結合して」で B1『数量』・C1『金額』が空になった（宣言は「書式だけ」）。
#
# 契約:
#   ① 依頼文のセルの書き方（A1:D1・A1からD1・A1とB1・1行目のA〜D列）が LLM の範囲に勝つ
#   ② 一致していれば何も足さない
#   ③ セルの書き方は在るが読めない（単独のセル・違う範囲が 2 つ）時は何も変えない
#   ④ 依頼がセルを言っていない: タイトル行が 1 つに決まれば表の幅で機械が決める・
#      決められなければ ⚠（_warnings ＝ ✓→△ の材料）で開示する
#   ⑤ 左上以外に値が在れば削除の関所（_confirm_delete）で聞く・無ければ黙る
#   ⑥ 見出しが 1 行目でない表でも同じ

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import argcheck  # noqa: E402

HEAD = ["品名", "東棟", "西棟", "数量"]
DATA = [["ナット", 3, 4, 7], ["ボルト", 5, 1, 6]]


def _book(tmp_path, header_row=1, title="部品在庫"):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "在庫"
    for i in range(header_row - 1):
        ws.append([title] if i == 0 else [])
    ws.append(HEAD)
    for r in DATA:
        ws.append(r)
    wb.save(p)
    return p


def _meta(path, header_row=1):
    return {"sheets": ["在庫"], "headers": {"在庫": list(HEAD)},
            "header_rows": {"在庫": header_row}, "path": str(path)}


def _merge(tmp_path, task, rng, header_row=1):
    ok, r, inferred, err = ailine.verify_dsl_args(
        "MERGE", {"range": rng}, _meta(_book(tmp_path, header_row), header_row), task=task)
    assert ok, err
    return r, inferred


def _range_warnings(r):
    return [w for w in r.get("_warnings", []) if "結合する範囲" in w]


# --- ① 依頼文の範囲が勝つ ------------------------------------------------------------------

def test_the_range_in_the_request_wins(tmp_path):
    r, inferred = _merge(tmp_path, "A1からC1を結合して", "A1:D1")
    assert r["range"] == "A1:C1"
    assert r["_sources"]["range"] == "依頼文: 『A1からC1』"
    assert "依頼文: 『A1からC1』" in ailine.format_confirmation_line("MERGE", r, inferred)


def test_other_ways_of_writing_a_range(tmp_path):
    for task, want in (("A1:B1を結合して", "A1:B1"), ("Ａ１：Ｂ１を結合して", "A1:B1"),
                       ("B1からA1まで結合して", "A1:B1"), ("A1とB1を結合して", "A1:B1"),
                       ("1行目のA〜C列を結合して", "A1:C1"), ("1行目のA列とB列を結合して", "A1:B1"),
                       ("A〜C列の1行目を結合して", "A1:C1")):
        r, _i = _merge(tmp_path, task, "A1:D1")
        assert r["range"] == want, task


# --- ② 一致していれば黙る ----------------------------------------------------------------

def test_agreement_adds_nothing(tmp_path):
    r, _i = _merge(tmp_path, "A1からC1を結合して", "A1:C1")
    assert "range" not in r.get("_sources", {}) and not _range_warnings(r)


# --- ③ 書き方は在るが読めない ---------------------------------------------------------------

def test_unreadable_cell_writing_changes_nothing(tmp_path):
    for task in ("A1を結合して", "A1:B1とC1:D1を結合して", "A1とB2を結合して"):
        r, _i = _merge(tmp_path, task, "A1:D1")
        assert r["range"] == "A1:D1" and "range" not in r.get("_sources", {}), task
        assert not _range_warnings(r), task


# --- ④ 依頼がセルを言っていない ------------------------------------------------------------

def test_a_made_up_range_is_disclosed(tmp_path):
    """地図の実例:「ナットの右に東棟」が A1:D1 になった ── 依頼は範囲を言っていない。"""
    r, _i = _merge(tmp_path, "ナットの右に東棟", "A1:D1")
    assert r["range"] == "A1:D1" and len(_range_warnings(r)) == 1
    assert "A1:D1 は解釈が選んだ範囲" in _range_warnings(r)[0]


def test_the_title_row_is_decided_by_the_machine(tmp_path):
    r, _i = _merge(tmp_path, "タイトルを結合して", "A1:B1", header_row=2)
    assert r["range"] == "A1:D1" and "タイトル行" in r["_sources"]["range"]
    assert not _range_warnings(r) and "_confirm_delete" not in r


def test_no_title_row_means_a_warning(tmp_path):
    r, _i = _merge(tmp_path, "タイトルを結合して", "A1:D1")
    assert _range_warnings(r)


def test_empty_task_is_silent(tmp_path):
    r, _i = _merge(tmp_path, "", "A1:D1")
    assert not _range_warnings(r)


# --- ⑤ 消える値は先に聞く ---------------------------------------------------------------------

def test_values_that_would_vanish_go_to_the_delete_gate(tmp_path):
    r, _i = _merge(tmp_path, "A1からC1を結合して", "A1:C1")
    ask = r["_confirm_delete"]
    assert "2 件消えます" in ask and "B1『東棟』" in ask and "C1『西棟』" in ask


def test_a_merge_that_loses_nothing_is_silent(tmp_path):
    r, _i = _merge(tmp_path, "A1からD1を結合して", "A1:D1", header_row=2)
    assert "_confirm_delete" not in r


# --- ⑥ 見出しが 1 行目でない表 -------------------------------------------------------------

def test_header_not_on_the_first_row(tmp_path):
    r, _i = _merge(tmp_path, "2行目のA〜B列を結合して", "A1:D1", header_row=2)
    assert r["range"] == "A2:B2" and "B2『東棟』" in r["_confirm_delete"]


# --- 変異: 配線を外すと黙って通る ---------------------------------------------------------

def test_without_the_reader_the_llm_range_passes(tmp_path, monkeypatch):
    monkeypatch.setattr(argcheck, "range_named_in_task", lambda *a, **k: None)
    monkeypatch.setattr(argcheck, "task_names_a_cell", lambda *a, **k: True)
    r, _i = _merge(tmp_path, "A1からC1を結合して", "A1:D1")
    assert r["range"] == "A1:D1" and not _range_warnings(r)


def test_without_the_gate_values_vanish_silently(tmp_path, monkeypatch):
    monkeypatch.setattr(argcheck, "_merge_would_erase", lambda *a, **k: [])
    r, _i = _merge(tmp_path, "A1からC1を結合して", "A1:C1")
    assert "_confirm_delete" not in r
