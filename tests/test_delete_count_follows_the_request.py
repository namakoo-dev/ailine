# 削除の件数は LLM に決めさせない（2026-10-01・依頼の項の台帳で D だった項目）。
#
# ★ 形: 「5行目を削除して」に LLM が count=3 を返すと 3 行消え、事後条件は
#   「宣言どおり 3 行減った」を確かめて ✓ を出す ── 宣言↔実体は合っていて、依頼の項が
#   欠けていた（並べ替えの向き 504efb6 と同じ形）。
#
# 契約:
#   ① 名前が複数行に当たった時は、件数はその行の数（機械が数えた）
#   ② 位置で指す依頼は依頼文の数字から読む（「5行目から3行」「5〜7行目」「5行目を削除」＝1）
#   ③ 名前で 1 行を指し、依頼文に数字が無ければ 1
#   ④ LLM と食い違えば機械が勝ち、解釈行に出典が出る（警告ではない）
#   ⑤ 一致する時・読めない時（裸の「3行」・始まりの行が宣言と違う）は何も変えない
#   ⑥ 見出しが 1 行目でない表でも同じ
#   ⑦ 事後条件の画面の件数は実際に消した行から数える（飛び飛びを「から N 行」と言わない）

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import argcheck  # noqa: E402
from ailine_core.anchor import row_count_in_task  # noqa: E402

HEAD = ["品名", "数量", "単価"]
DATA = [["りんご", 3, 100], ["みかん", 5, 80], ["ぶどう", 2, 300],
        ["もも", 4, 250], ["なし", 6, 120], ["みかん", 1, 90], ["かき", 7, 60]]


def _book(tmp_path, header_row=1, data=DATA, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "在庫"
    for _ in range(header_row - 1):
        ws.append(["在庫一覧"] if _ == 0 else [])
    ws.append(HEAD)
    for r in data:
        ws.append(r)
    wb.save(p)
    return p


def _meta(path, header_row=1):
    return {"sheets": ["在庫"], "headers": {"在庫": list(HEAD)},
            "header_rows": {"在庫": header_row}, "path": str(path)}


def _delete(tmp_path, task, at, count=None, header_row=1):
    args = {"at": at} if count is None else {"at": at, "count": count}
    ok, r, inf, err = ailine.verify_dsl_args(
        "DELETE_ROWS", args, _meta(_book(tmp_path, header_row), header_row), task=task)
    assert ok, err
    return r, inf


# --- 読み手（純関数）-------------------------------------------------------------------

def test_the_reader_reads_the_forms_it_claims():
    assert row_count_in_task("5行目から3行削除して") == (3, 5, "5行目から3行")
    assert row_count_in_task("5〜7行目を削除して")[:2] == (3, 5)
    assert row_count_in_task("5行目から7行目までを削除して")[:2] == (3, 5)
    assert row_count_in_task("５行目を消して")[:2] == (1, 5)


def test_the_reader_is_silent_when_it_cannot_decide():
    for t in ("3行削除して", "5行目と7行目を削除して", "7行目から5行目を削除して",
              "りんごの行を削除して", "『5行目』を削除して", "5行目から3行、8行目を削除して"):
        assert row_count_in_task(t) is None, t


# --- ②④ 位置で指した依頼 ----------------------------------------------------------------

def test_one_row_number_means_one_row(tmp_path):
    r, inf = _delete(tmp_path, "5行目を削除して", at=5, count=3)
    assert r["count"] == 1
    assert r["_sources"]["count"] == "依頼文: 『5行目』"
    assert "count" not in inf


def test_from_n_take_k_overrides_the_answer(tmp_path):
    r, _ = _delete(tmp_path, "3行目から3行削除して", at=3, count=1)
    assert r["count"] == 3 and "3行目から3行" in r["_sources"]["count"]


def test_a_range_overrides_the_answer(tmp_path):
    r, _ = _delete(tmp_path, "3〜5行目を削除して", at=3, count=2)
    assert r["count"] == 3


def test_the_confirmation_counts_the_same_number(tmp_path):
    """★ 聞く文（値の在る行を消す前の確認）と消す件数が同じ 1 つの値から出る。"""
    r, _ = _delete(tmp_path, "5行目を削除して", at=5, count=3)
    assert "から" not in r["_confirm_delete"] and r["_confirm_delete"].startswith("5行目には")


def test_the_disclosure_reaches_the_interpretation_line(tmp_path):
    r, inf = _delete(tmp_path, "5行目を削除して", at=5, count=3)
    line = ailine.format_confirmation_line("DELETE_ROWS", r, inf)
    assert "行数:1（依頼文: 『5行目』）" in line


# --- ①③ 名前で指した依頼 ----------------------------------------------------------------

def test_a_name_on_several_rows_counts_those_rows(tmp_path):
    r, _ = _delete(tmp_path, "みかんの行を削除して", at=3, count=1)
    assert r["_delete_rows"] == [3, 7] and r["count"] == 2
    assert "count" in r["_sources"]


def test_a_name_on_one_row_is_one_row(tmp_path):
    r, _ = _delete(tmp_path, "ぶどうの行を削除して", at=4, count=2)
    assert r["at"] == 4 and r["count"] == 1
    assert r["_sources"]["count"] == "依頼文が名指しした 1 行"


# --- ⑤ 何も変えない ----------------------------------------------------------------------

def test_agreeing_answer_is_left_alone(tmp_path):
    r, _ = _delete(tmp_path, "3行目から3行削除して", at=3, count=3)
    assert r["count"] == 3 and "count" not in r.get("_sources", {})


def test_unreadable_count_is_left_alone(tmp_path):
    r, _ = _delete(tmp_path, "3行削除して", at=5, count=2)
    assert r["count"] == 2 and "count" not in r.get("_sources", {})


def test_a_start_that_disagrees_with_the_declaration_is_left_alone(tmp_path):
    """★ 件数だけ直すと別の行を消す（位置の食い違いは位置の関所の仕事）。"""
    r, _ = _delete(tmp_path, "5行目を削除して", at=6, count=3)
    assert r["count"] == 3 and "count" not in r.get("_sources", {})


def test_a_digit_in_a_named_request_is_left_alone(tmp_path):
    """名前で指していても、依頼文に数字が在る（「ぶどうの行を2つ」）なら 1 と決めない。"""
    r, _ = _delete(tmp_path, "ぶどうの行から2つ削除して", at=4, count=2)
    assert r["count"] == 2


# --- ⑥ 見出しが 1 行目でない表 -----------------------------------------------------------

def test_header_not_on_the_first_row(tmp_path):
    r, _ = _delete(tmp_path, "7行目を削除して", at=7, count=3, header_row=3)
    assert r["count"] == 1
    r2, _ = _delete(tmp_path, "5〜7行目を削除して", at=5, count=1, header_row=3)
    assert r2["count"] == 3


# --- 変異: 配線を外すと直らない（試験が読み手を通っていることの証明）-------------------------

def test_without_the_wiring_the_llm_count_would_win(tmp_path, monkeypatch):
    monkeypatch.setattr(argcheck, "row_count_in_task", lambda task: None)
    r, _ = _delete(tmp_path, "5行目を削除して", at=5, count=3)
    assert r["count"] == 3          # ← 直す前の挙動（嘘の ✓ の材料）


# --- ⑦ 事後条件の画面の件数 --------------------------------------------------------------

def test_postcondition_names_non_contiguous_rows(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    kept = [r for r in DATA if r[0] != "みかん"]
    after = _book(tmp_path, data=kept, name="after.xlsx")
    # ★ 宣言の count（1）と一覧（2 行）がずれた回でも、画面は実際に消した 2 行を言う
    args = {"at": 3, "count": 1, "_delete_rows": [3, 7]}
    status, reason = ailine.check_delete_rows(after, args, header_row=1, source_book=before)
    assert status == "pass", reason
    assert reason.startswith("3、7行目の 2 行を削除"), reason


def test_postcondition_contiguous_message_is_unchanged(tmp_path):
    before = _book(tmp_path, name="before.xlsx")
    after = _book(tmp_path, data=DATA[:1] + DATA[2:], name="after.xlsx")
    status, reason = ailine.check_delete_rows(after, {"at": 3, "count": 1}, header_row=1,
                                              source_book=before)
    assert status == "pass", reason
    assert reason == "3行目から 1 行を削除（下の行は上へ詰まりますが、中身と式は保たれています）"
