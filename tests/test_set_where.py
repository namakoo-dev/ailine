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
from _reread_home import segment as reread_segment  # noqa: E402 ── ★ 層の場所は 1 箇所が持つ
from _product_source import product_text  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む

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
    # ★ 末尾の "" は「対象から外す行」（合計行など）。この表には無いので空。
    assert 'Call SetColumnValueWhere(oDoc, 0, 3, 2, 0, 500.0, "◎", "")' in code, code


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


# --- 表記の揺れ（2026-08-27・Namakoo「◎を入れて では動作しない」）--------------------

def test_the_trigger_does_not_depend_on_the_verb():
    """★ 引き金は**比較語と引用**だけで、動詞（付ける/入れる/書き込む/記入する）を見ない。
       動詞を並べ始めると、並べ忘れた言い方が黙って落ちる。"""
    for verb in ["「◎」を付けて", "「◎」を入れて", "「◎」を書き込んで", "「◎」と記入して",
                  "「◎」を記入して", "「◎」にして"]:
        task = "原価が500以上の行のチェック列に" + verb
        assert ailine.task_asks_for_a_conditional_write(task), task


def test_writing_into_a_column_is_not_adding_a_column():
    """★★ 実測した事故: 「列**に**『◎』を入れて」を**列追加**が横取りし、条件つき書換が
       上書きされていた（「入れ」の 1 語で当たっていた）。
       ★ 助詞が意味を運ぶ: 「列**を**入れる」＝列そのもの／「列**に**…を入れる」＝行き先。"""
    task = "原価が500以上の項目のチェック列に「◎」を入れて"
    assert ailine.task_asks_for_a_conditional_write(task)
    assert not ailine.task_asks_to_add_a_column(task)
    # ★ 恒真殺し: 塞いだ側で、正当な列追加まで殺していないこと
    assert ailine.task_asks_to_add_a_column("原価の右に空の列を入れて")


def test_a_reread_never_overwrites_another_reread():
    """★★ 構造側の真因: 読み直しの塊が 5 つ並び、**後の塊が前の結果を上書き**していた。
       ★ 個々の条件をいくら賢くしてもこの形の事故は消えない ── 塊が増えるたびに
         「まだ上書きされない」ことを人が確かめる羽目になる。1 回に縛る。"""
    # ★ 層の場所は tests/_reread_home.py が 1 箇所で持つ（2026-09-05 に切り出した時、
    #   4 つの試験が同じ関数名を手書きしていて同時に落ちた ── 試験側の片配線）。
    body = reread_segment()
    sets = body.count("_reread_done = True")
    guards = body.count("not _reread_done")
    # ★ 数を固定しない（塊は増える）。**不変**は「印を立てる塊と、印を見る塊が同数」。
    #   数で縛ると塊を足すたびに試験を直すことになり、直すついでに緩めてしまう。
    assert sets >= 5 and sets == guards, (
        f"読み直しの塊のうち、印を見ていないものがある（立てる {sets} / 見る {guards}）")


# --- 置き換え「『A』を『B』に」（2026-08-27・Namakoo「置き換えができない」）------------
#
# ★ 実測: 「チェック列の『◎』を全て『合格』に書き換えて」は一段目が
#   SET_COLUMN_VALUE（列を丸ごと『合格』に）を返していた ── **空欄の行まで潰す**。
#   機械は引用が 2 つあるので「値が一意に読み取れない」と正しく断っていたが、
#   **断って終わり**だった。★ 引用が 2 つある時の意味（置き換え）を読む。
# ★ 新しい op は要らない: 同じ列の中で A の行だけ B にする＝条件つき書換の特別な場合
#   （条件列＝書き込み先列・比較は「等しい」）。

@pytest.mark.parametrize("task,pair", [
    ("チェック列の「◎」を全て「合格」に書き換えて", ("◎", "合格")),
    ("チェック列の「◎」を「合格」にして", ("◎", "合格")),
    ("商品列の『りんご』を『林檎』に置き換えて", ("りんご", "林檎")),
    ("備考列を全部「確認済み」にして", None),                       # 引用が 1 つ＝置き換えでない
    ("原価が500以上の行のチェック列に「◎」を入れて", None),          # 条件つき書換の領分
])
def test_the_replace_pair_is_read_by_particles_not_verbs(task, pair):
    """★ 動詞（書き換え/置換/直す…）を並べない ── 並べ忘れた言い方が黙って落ちる。
       助詞が意味を運ぶ（『A』**を** … 『B』**に**）。"""
    assert ailine.extract_replace_pair(task) == pair


def test_a_replace_resolves_to_a_conditional_write(tmp_path):
    rows = [["商品", "売上", "原価", "チェック"], ["りんご", 1200, 700, "◎"],
             ["みかん", 800, 300, None], ["ぶどう", 1500, 900, "◎"]]
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "チェック"}, _meta(_book(tmp_path, rows)),
        task="チェック列の「◎」を全て「合格」に書き換えて")
    assert ok, err
    assert (r["cmp"], r["cond_value"], r["value"]) == ("eq", "◎", "合格")
    assert r["cond_col"] == "チェック", "条件列は書き込み先と同じ列のはず"
    assert r["_match_rows"] == [2, 4], r["_match_rows"]     # 空欄のみかんは入らない


def test_a_replace_of_a_value_that_is_not_there_is_refused(tmp_path):
    rows = [["商品", "売上", "原価", "チェック"], ["りんご", 1200, 700, "◎"]]
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "チェック"}, _meta(_book(tmp_path, rows)),
        task="チェック列の「×」を「合格」に書き換えて")
    assert not ok
    assert "『×』の行がありません" in err and "何も書いていません" in err, err


def test_codegen_uses_a_string_literal_for_a_text_threshold(tmp_path):
    """★★ 実測で生の traceback を出した: 閾値が数値かどうかは **cmp ではなく値**で決まる。
       eq は「金額が 100 と等しい」にも「チェックが『◎』と等しい」にも使う ──
       cmp で分けると、置き換えの『◎』を float() に渡して落ちる。"""
    rows = [["商品", "売上", "原価", "チェック"], ["りんご", 1200, 700, "◎"]]
    meta = _meta(_book(tmp_path, rows))
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "チェック"}, meta, task="チェック列の「◎」を「合格」にして")
    assert ok, err
    code = ailine.codegen_dsl("SET_WHERE", r, meta)
    assert 'SetColumnValueWhere(oDoc, 0, 3, 3, 4, "◎", "合格", "")' in code, code


# --- 合計行はデータ行ではない（2026-08-28・Namakoo が請求書のデモで実測）------------

def test_a_total_row_is_not_marked(tmp_path):
    """★★ 実測: 「金額が10万以上の行に印を付けて」が**合計行にも印を付けた**。
       条件としては真だが、合計行は請求の行ではない ── 意味が違う。
       ★ 判定は既存の凍結規則（ailine_core.total_row）を借りる ── 新しい規則を書かない。
       ★ 外したことは必ず画面に出す（黙って行を外さない）。"""
    rows = [["商品", "売上", "原価", "チェック"], ["りんご", 1200, 700, None],
             ["みかん", 800, 300, None], ["合計", 2000, 1000, None]]
    meta = _meta(_book(tmp_path, rows))
    assert ailine.total_rows_in(meta, "売上") == [4]
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "チェック", "cond_col": "売上", "cmp": "gte"}, meta,
        task="売上が700以上の行のチェック列に「◎」を付けて")
    assert ok, err
    assert r["_match_rows"] == [2, 3], r["_match_rows"]      # 合計行(4)は入らない
    assert r["_skip_rows"] == [4] and "合計行" in r["_skip_label"], r.get("_skip_label")
    # ★ 外す行は Basic にも渡る（0 起点）── 書き手も同じ行を触らない
    assert 'SetColumnValueWhere(oDoc, 0, 3, 1, 0, 700.0, "◎", "3")' in         ailine.codegen_dsl("SET_WHERE", r, meta), ailine.codegen_dsl("SET_WHERE", r, meta)


def test_a_total_row_that_was_written_anyway_fails(tmp_path):
    """★ 恒真殺し: 外した行が黙って書き換わるのは、外していないのと同じくらい悪い。
       事後条件は「外した行は**変わっていない**」ことまで要求する。"""
    rows = [["商品", "売上", "原価", "チェック"], ["りんご", 1200, 700, None],
             ["みかん", 800, 300, None], ["合計", 2000, 1000, None]]
    before = _book(tmp_path, rows, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "チェック"],
                              ["りんご", 1200, 700, "◎"], ["みかん", 800, 300, "◎"],
                              ["合計", 2000, 1000, "◎"]])
    args = {"col": "チェック", "cond_col": "売上", "cmp": "gte", "cond_value": 700.0,
             "value": "◎", "_header_row": 1, "_skip_rows": [4]}
    status, reason = ailine.check_set_where(after, args, source_book=before)
    assert status == "fail" and "広がった疑い" in reason, reason
