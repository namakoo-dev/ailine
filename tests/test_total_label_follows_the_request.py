# 合計のラベルは依頼文から・既にある合計式は黙って書き換えない（2026-10-01）。
#
# ★ 形 1（依頼の項の台帳で D だった項目）: ラベルは LLM の値（無ければ『合計』）がそのまま
#   書かれていた。「小計を出して」に『売上合計』が返っても、事後条件は宣言のラベルが
#   書かれたかを確かめて ✓ を出す。
# ★ 形 2（実機で確かめた）: 既にある合計行の =SUM(C2:C3) が確認なしで
#   =SUM(C2:INDEX(C:C,ROW()-1)) に書き換わり、値が 800 → 1500 に変わって ✓ が出た。
#
# 契約:
#   ① 依頼文のラベルの語（合計・小計・税込み合計…）が LLM のラベルに勝つ（解釈行に出典）
#   ② LLM のラベルが依頼文に在れば何も足さない
#   ③ 依頼文にラベルの語が無ければ既定の『合計』。ただし 税/込 を含む LLM のラベルは残す
#      （税の関所の材料）。2 種類の語が読める時・列名の中の語は何も変えない
#   ④ 既にある合計行のラベル（実表）が正の回は触らない
#   ⑤ 既にある =SUM が、書く式と範囲（か倍率）が違えば上書きの関所に載せる・同じなら黙る
#   ⑥ 見出しが 1 行目でない表でも同じ

import sys
from pathlib import Path
from types import SimpleNamespace

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import argcheck, dsl_step  # noqa: E402

HEAD = ["商品", "数量", "金額"]
DATA = [["りんご", 3, 300], ["みかん", 5, 500], ["ぶどう", 2, 700]]


def _book(tmp_path, header_row=1, total=None, head=HEAD):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for i in range(header_row - 1):
        ws.append(["売上一覧"] if i == 0 else [])
    ws.append(list(head))
    for r in DATA:
        ws.append(r)
    if total is not None:
        ws.append(["合計", None, total])
    wb.save(p)
    return p


def _meta(path, header_row=1, head=HEAD):
    return {"sheets": ["売上"], "headers": {"売上": list(head)},
            "header_rows": {"売上": header_row}, "path": str(path)}


def _total(tmp_path, task, args, header_row=1, total=None, head=HEAD):
    ok, r, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", dict(args),
        _meta(_book(tmp_path, header_row, total, head), header_row, head), task=task)
    assert ok, err
    return r, inferred


# --- ① 依頼文の語が勝つ ------------------------------------------------------------------

def test_the_label_in_the_request_wins(tmp_path):
    r, inferred = _total(tmp_path, "金額の小計を出して", {"col": "金額", "label": "売上合計"})
    assert r["label"] == "小計"
    assert r["_sources"]["label"] == "依頼文: 『小計』"
    assert "依頼文: 『小計』" in ailine.format_confirmation_line("APPEND_TOTAL", r, inferred)


def test_a_made_up_label_goes_back_to_the_requested_word(tmp_path):
    r, _i = _total(tmp_path, "金額の合計を出して", {"col": "金額", "label": "売上合計"})
    assert r["label"] == "合計" and "label" in r["_sources"]


# --- ② 依頼文に在るラベルはそのまま -------------------------------------------------------

def test_a_label_in_the_request_is_kept(tmp_path):
    r, _i = _total(tmp_path, "金額の税込み合計を一番下に出して（消費税10%）",
                   {"col": "金額", "label": "税込み合計"})
    assert r["label"] == "税込み合計" and "label" not in r.get("_sources", {})


def test_the_default_matching_the_request_adds_nothing(tmp_path):
    r, _i = _total(tmp_path, "金額の合計を出して", {"col": "金額"})
    assert r["label"] == "合計" and "label" not in r.get("_sources", {})


# --- ③ 語が無ければ既定・読めない形は変えない ---------------------------------------------

def test_no_label_word_means_the_default(tmp_path):
    r, _i = _total(tmp_path, "金額を一番下に足して", {"col": "金額", "label": "総額"})
    assert r["label"] == "合計" and "既定" in r["_sources"]["label"]


def test_a_tax_label_is_kept_for_the_tax_gate(tmp_path):
    """「消費税込みでいくら」── 税/込 のラベルを消すと、倍率の関所が鳴らなくなる。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", {"col": "金額", "label": "消費税込み合計"},
        _meta(_book(tmp_path)), task="消費税込みでいくらになるか教えて")
    assert r["label"] == "消費税込み合計"
    assert not ok or r.get("factor") != 1.0      # 倍率が決まらなければ聞き返す（今までどおり）


def test_two_label_words_change_nothing(tmp_path):
    r, _i = _total(tmp_path, "小計と総額を出して", {"col": "金額", "label": "売上合計"})
    assert r["label"] == "売上合計" and "label" not in r.get("_sources", {})


def test_a_column_named_like_a_label_is_not_a_label(tmp_path):
    head = ["商品", "数量", "小計"]
    r, _i = _total(tmp_path, "小計の列を一番下で合計して", {"col": "小計", "label": "合計"}, head=head)
    assert r["label"] == "合計" and "label" not in r.get("_sources", {})


# --- ④ 既にある合計行のラベルが正 ---------------------------------------------------------

def test_the_label_of_an_existing_total_row_is_kept(tmp_path):
    r, _i = _total(tmp_path, "金額の小計を出して", {"col": "金額"}, total=None)
    assert r["label"] == "小計"
    r2, _i = _total(tmp_path, "金額の小計を出して", {"col": "金額"}, total="")
    assert r2["_at_row"] == 5 and r2["label"] == "合計"   # 実表の行のラベル


# --- ⑤ 既にある =SUM の上書き -------------------------------------------------------------

def test_a_different_sum_goes_to_the_overwrite_gate(tmp_path):
    r, _i = _total(tmp_path, "金額の合計を出して", {"col": "金額"}, total="=SUM(C2:C3)")
    assert r["_at_row"] == 5
    assert "=SUM(C2:C3)" in r["_confirm_overwrite"] and "=SUM(C2:C4)" in r["_confirm_overwrite"]


def test_the_same_sum_is_silent(tmp_path):
    for f in ("=SUM(C2:C4)", "=SUM($C$2:$C$4)", "=SUM(C2:INDEX(C:C,ROW()-1))"):
        r, _i = _total(tmp_path, "金額の合計を出して", {"col": "金額"}, total=f)
        assert r["_at_row"] == 5 and "_confirm_overwrite" not in r, f


def test_a_changed_factor_goes_to_the_gate(tmp_path):
    r, _i = _total(tmp_path, "金額の税込み合計を出して（消費税10%）", {"col": "金額"},
                   total="=SUM(C2:C4)")
    assert r.get("factor") == 1.1 and "_confirm_overwrite" in r


def test_an_empty_total_cell_is_silent(tmp_path):
    r, _i = _total(tmp_path, "金額の合計を出して", {"col": "金額"}, total="")
    assert r["_at_row"] == 5 and "_confirm_overwrite" not in r


def test_the_gate_reads_the_new_key(tmp_path, capsys):
    """★ 帰結の配線: _confirm_overwrite は関所（print_dsl_confirmation）へ渡り、
    聞く文は「削除しますか？」でなく既定の上書き（prompt=None）。"""
    seen = {}

    def _gate(a, warn, **kw):
        seen["warn"], seen["prompt"] = warn, kw.get("prompt")
        return None

    deps = SimpleNamespace(
        format_confirmation_line=lambda *a, **k: "解釈: 操作:合計追加",
        maybe_warn_header_col_mismatch=lambda *a, **k: None,
        maybe_warn_target_overwrite=lambda *a, **k: None,
        interpretation_summary_line=lambda *a, **k: None,
        confirm_overwrite_or_gate=_gate,
        classify_subject_provenance=None, sheet_conflict_gate=None)
    note = "★ 合計行（5行目）の金額には既に =SUM(C2:C3) が入っています"
    dsl_step.print_dsl_confirmation(
        "APPEND_TOTAL", {"col": "金額", "_confirm_overwrite": note}, set(), "合計を出して",
        meta={"sheets": ["売上"]}, warn_book=tmp_path / "b.xlsx", new_cols=None,
        a=SimpleNamespace(), deps=deps)
    assert seen == {"warn": note, "prompt": None}
    assert note in capsys.readouterr().out


# --- ⑥ 見出しが 1 行目でない表 -------------------------------------------------------------

def test_header_not_on_the_first_row(tmp_path):
    r, _i = _total(tmp_path, "金額の小計を出して", {"col": "金額", "label": "合計"}, header_row=3)
    assert r["label"] == "小計"
    r2, _i = _total(tmp_path, "金額の合計を出して", {"col": "金額"}, header_row=3,
                    total="=SUM(C4:C5)")
    assert r2["_at_row"] == 7 and "=SUM(C4:C6)" in r2["_confirm_overwrite"]
    r3, _i = _total(tmp_path, "金額の合計を出して", {"col": "金額"}, header_row=3,
                    total="=SUM(C4:C6)")
    assert "_confirm_overwrite" not in r3


# --- 変異: 配線を外すと黙って通る ---------------------------------------------------------

def test_without_the_label_wiring_the_llm_label_passes(tmp_path, monkeypatch):
    monkeypatch.setattr(argcheck, "_total_label_from_request", lambda *a, **k: None)
    r, _i = _total(tmp_path, "金額の小計を出して", {"col": "金額", "label": "売上合計"})
    assert r["label"] == "売上合計"


def test_without_the_gate_wiring_the_sum_is_rewritten_silently(tmp_path, monkeypatch):
    monkeypatch.setattr(argcheck, "_sum_overwrite_note", lambda *a, **k: None)
    r, _i = _total(tmp_path, "金額の合計を出して", {"col": "金額"}, total="=SUM(C2:C3)")
    assert "_confirm_overwrite" not in r
