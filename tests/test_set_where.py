# 条件つき書換（SET_WHERE）── 2026-08-27。Namakoo「原価が500以上の項目に『◎』を付ける」
#
# ★ 表計算のごく普通の操作なのに、一覧に無かった（SET_COLUMN_VALUE は列を丸ごと同じ値に
#   するだけ）。実測でも 4/4 で OUT_OF_VOCAB ── しかも「条件付き書式」と誤って読まれていた。
#   人が欲しいのは**値**であって書式ではない。
#
# 契約:
#   ① 比較（以上/以下/…）は依頼文から機械が取る（LLM と食い違えば機械が勝つ）
#   ② 閾値の数字も依頼文から。2 つ以上あって決まらないなら断る
#   ③ 書き込む値は「」『』で囲まれたものだけ（LLM に作らせない）
#   ④ 当てはまる行が 0 なら、**走らせる前に**断る（後から「変化なし」で × は不親切）
#   ⑤ 事後条件は両側を見る: 当てはまる行は書かれ、**当てはまらない行は 1 セルも変わらない**
#      ── 「付いたか」だけ見ると全行に付けても pass する
#   ⑥ 条件は**適用前のファイル**から独立に評価する（書き手の判定を借りない）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

ROWS = [["商品", "売上", "原価", "チェック"],
         ["りんご", 1200, 700, None],
         ["みかん", 800, 300, None],
         ["ぶどう", 1500, 900, None]]


def _book(tmp_path, rows=None, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for r in (rows or ROWS):
        ws.append(r)
    wb.save(p)
    return p


def _meta(path):
    return {"sheets": ["売上"], "headers": {"売上": ["商品", "売上", "原価", "チェック"]},
            "header_rows": {"売上": 1}, "path": str(path)}


TASK = "原価が500以上の項目のチェック列に「◎」を付けて"
ARGS = {"col": "チェック", "cond_col": "原価", "cmp": "gte"}


# --- ①②③④ 関所 ---------------------------------------------------------------------

def test_a_normal_request_resolves(tmp_path):
    ok, r, _i, err = ailine.verify_dsl_args("SET_WHERE", dict(ARGS), _meta(_book(tmp_path)),
                                             task=TASK)
    assert ok, err
    assert r["cmp"] == "gte" and r["cond_value"] == 500.0 and r["value"] == "◎"
    assert r["_cond_label"] == "『原価』が 500 以上"
    assert r["_match_rows"] == [2, 4], r["_match_rows"]      # りんご 700 / ぶどう 900


def test_the_machine_beats_the_llm_on_the_comparison(tmp_path):
    """① 「以上」と書いてあるのに LLM が lte と言ったら、機械が勝つ（境界行が静かに入れ替わる）。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {**ARGS, "cmp": "lte"}, _meta(_book(tmp_path)), task=TASK)
    assert ok, err
    assert r["cmp"] == "gte"
    assert any("機械抽出(gte)を採用" in w for w in r.get("_warnings", [])), r.get("_warnings")


def test_an_ambiguous_threshold_is_refused(tmp_path):
    """② 数字が 2 つある依頼は断る ── どちらが閾値か機械には決められない。"""
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", dict(ARGS), _meta(_book(tmp_path)),
        task="原価が500以上で売上が1000以上の行のチェック列に「◎」を付けて")
    assert not ok
    assert "一意に読み取れません" in err, err


def test_an_unquoted_value_is_refused(tmp_path):
    """③ 書き込む値は引用符から。LLM に作らせない（SET_COLUMN_VALUE と同じ関所）。"""
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", dict(ARGS), _meta(_book(tmp_path)),
        task="原価が500以上の項目のチェック列に◎を付けて")
    assert not ok
    assert "「」または『』で囲んで" in err, err


def test_zero_matching_rows_is_refused_before_running(tmp_path):
    """④ 当てはまる行が 0 なら、走らせる前に断る（列追加で同じ形を実測した）。"""
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", dict(ARGS), _meta(_book(tmp_path)),
        task="原価が99999以上の項目のチェック列に「◎」を付けて")
    assert not ok
    assert "当てはまる行がありません" in err and "何も書いていません" in err, err


def test_a_column_the_sheet_does_not_have_is_refused(tmp_path):
    """依頼文にも手掛かりが無ければ断る。"""
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "無い列", "cond_col": "無い列2", "cmp": "gte"},
        _meta(_book(tmp_path)), task="なにかが500以上の行に「◎」を付けて")
    assert not ok, err


def test_a_wrong_column_is_rescued_from_the_request_text(tmp_path):
    """★ 既存の作法（機械抽出が LLM に勝つ）はこの op でも効く: LLM が実在しない列を
       返しても、依頼文が実在列を一意に名指ししていればそちらを採り、**採ったことを言う**。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "チェック", "cond_col": "無い列", "cmp": "gte"},
        _meta(_book(tmp_path)), task=TASK)
    assert ok, err
    assert r["cond_col"] == "原価"
    assert any("依頼文が名指しする列『原価』を採用" in w for w in r.get("_warnings", [])),         r.get("_warnings")


def test_the_request_is_recognised():
    """★ 比較語と引用の**両方**が揃った時だけ疑う ── 片方だけなら別の op の領分。"""
    assert ailine.task_asks_for_a_conditional_write(TASK)
    assert not ailine.task_asks_for_a_conditional_write("原価が500以上の行を抜き出して")
    assert not ailine.task_asks_for_a_conditional_write("備考列を全部「確認済み」にして")


# --- 生成 ------------------------------------------------------------------------------

def test_codegen(tmp_path):
    ok, r, _i, err = ailine.verify_dsl_args("SET_WHERE", dict(ARGS), _meta(_book(tmp_path)),
                                             task=TASK)
    assert ok, err
    code = ailine.codegen_dsl("SET_WHERE", r, _meta(_book(tmp_path)))
    assert 'Call SetColumnValueWhere(oDoc, 0, 3, 2, 0, 500.0, "◎")' in code, code


def test_the_comparison_lives_in_one_place_in_the_bas():
    """★★ 判定を 2 箇所に書き写さない: 抽出(ExtractRows)と条件つき書換が同じ
       RowMatches を呼ぶ。書き写すと必ず片方だけ直る（この repo は今日 3 回踏んだ）。"""
    bas = (REPO / "src" / "ailine" / "helpers" / "AiLineHelpers.bas").read_text(encoding="utf-8")
    assert bas.count("Select Case cmpCode") == 1, "比較の Select Case が 2 箇所ある"
    assert "Function RowMatches(" in bas
    assert bas.count("RowMatches(") >= 3     # 定義 + ExtractRows + SetColumnValueWhere
    assert "Sub SetColumnValueWhere(" in bas


# --- ⑤⑥ 事後条件 ----------------------------------------------------------------------

def _pc_args():
    return {"col": "チェック", "cond_col": "原価", "cmp": "gte", "cond_value": 500.0,
             "value": "◎", "_header_row": 1}


def test_a_correct_conditional_write_passes(tmp_path):
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "チェック"],
                              ["りんご", 1200, 700, "◎"],
                              ["みかん", 800, 300, None],
                              ["ぶどう", 1500, 900, "◎"]])
    status, reason = ailine.check_set_where(after, _pc_args(), source_book=before)
    assert status == "pass", reason
    assert "2 行だけ" in reason, reason


def test_writing_to_every_row_fails(tmp_path):
    """★★ 恒真殺しの本命: 「付いたか」だけ見る番人なら**全行に付けても pass する**。
       当てはまらない行が変わっていないことまで見て、初めて条件つきの証明になる。"""
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "チェック"],
                              ["りんご", 1200, 700, "◎"],
                              ["みかん", 800, 300, "◎"],      # ← 300 なのに付いた
                              ["ぶどう", 1500, 900, "◎"]])
    status, reason = ailine.check_set_where(after, _pc_args(), source_book=before)
    assert status == "fail" and "広がった疑い" in reason, reason


def test_missing_a_matching_row_fails(tmp_path):
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "チェック"],
                              ["りんご", 1200, 700, "◎"],
                              ["みかん", 800, 300, None],
                              ["ぶどう", 1500, 900, None]])   # ← 900 なのに付いていない
    status, reason = ailine.check_set_where(after, _pc_args(), source_book=before)
    assert status == "fail" and "なっていない行" in reason, reason


def test_disturbing_another_column_fails(tmp_path):
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "チェック"],
                              ["りんご", 9999, 700, "◎"],     # ← 売上を巻き込んだ
                              ["みかん", 800, 300, None],
                              ["ぶどう", 1500, 900, "◎"]])
    status, reason = ailine.check_set_where(after, _pc_args(), source_book=before)
    assert status == "fail" and "以外の列が変わっています" in reason, reason


def test_the_condition_is_read_from_the_before_file(tmp_path):
    """⑥ 条件を**適用後**の値で見ていたら、この検体は通ってしまう。
       適用後の原価は書き換えられているので、before から読んでいることの証明になる。"""
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "チェック"],
                              ["りんご", 1200, 100, "◎"],     # 原価が 100 に化けている
                              ["みかん", 800, 300, None],
                              ["ぶどう", 1500, 900, "◎"]])
    status, reason = ailine.check_set_where(after, _pc_args(), source_book=before)
    assert status == "fail", f"原価が書き換わっているのに通した: {reason}"


def test_without_the_before_file_it_does_not_claim(tmp_path):
    after = _book(tmp_path, ROWS)
    status, _r = ailine.check_set_where(after, _pc_args(), source_book=None)
    assert status == "warn"


def test_the_bulk_fill_advisory_is_silent_when_the_postcondition_proves_the_rows():
    """★ 実測: 正しく 2 行だけ書いたのに「空欄への同一値の一括書き込み」が立って ✓ が
       △ に落ちた。この op の事後条件は**両方向**（書かれるべき行は書かれ、それ以外は
       1 セルも変わらない）を証明する ── 助言は証明が届かない所にだけ要る。
       ★ SET_COLUMN_VALUE は据え置き: あちらは「全データ行がその値か」しか見ず、
         元が空欄だったかを問わないので、助言がまだ仕事をする。"""
    assert ailine.OP_WRITE_TARGET["SET_WHERE"].proves_which_cells is True
    assert ailine.OP_WRITE_TARGET["SET_COLUMN_VALUE"].proves_which_cells is False
    before = {"cells": {"売上!2,4": (None, None), "売上!4,4": (None, None)}}
    after = {"cells": {"売上!2,4": ("◎", None), "売上!4,4": ("◎", None)}}
    assert ailine.detect_uniform_fill(before, after) is not None, "前提: 既定では鳴る"
    assert ailine.detect_uniform_fill(before, after, proved=True) is None
