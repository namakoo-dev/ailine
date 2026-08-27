# 列の追加（ADD_COLUMN）── 2026-08-27。Namakoo「列の追加はできないの？」
#
# ★ 実測でその通りだった: 列は**削除だけ**あって追加が無かった。行は空行(INSERT_ROWS)と
#   値つき(ADD_ROW)の 2 つがあるのに、列は片側だけ ── ADD_ROW を足した時と同じ形の欠け。
#   「備考という列を追加して」も「原価の右に列を追加して」も語彙外で断られていた
#   （Namakoo が GUI で実測・画面の「要望として記録します」）。
#
# 契約:
#   ① 位置は**機械が実表の見出しから**決める（LLM に数えさせない）
#   ② 位置の言い回しが解けないなら**断る**（黙って末尾に付けない）
#   ③ 名前は任意 ── 「列を追加して」しか言わない依頼が実在する。見出しは空のまま入れて、
#      そのことを画面に書く
#   ④ 同名の列が既に在れば断る（2 本目を黙って作らない）
#   ⑤ 事後条件: 列が 1 本増え、宣言した位置に宣言した名前が在り、**他の列は 1 セルも
#      変わらない**。ただし式は列が動けば参照が追随する（比べ方は 1 箇所に任せる）
#   ⑥ 挿した列のデータ行は空（勝手に埋めない）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

ROWS = [["商品", "売上", "原価"], ["りんご", 1200, 700],
         ["みかん", 800, 300], ["ぶどう", 1500, 900]]
META = {"sheets": ["売上"], "headers": {"売上": ["商品", "売上", "原価"]},
         "header_rows": {"売上": 1}}


def _book(tmp_path, rows=None, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for r in (rows or ROWS):
        ws.append(r)
    wb.save(p)
    return p


# --- ① 位置の解決 ---------------------------------------------------------------------

@pytest.mark.parametrize("task,at,shown", [
    ("原価の右に備考の列を追加して", 4, "『原価』（3列目）の右"),
    ("原価列の右に列を追加して", 4, "『原価』（3列目）の右"),      # 「〜列」と書かれても解ける
    ("商品の左にコードの列を追加して", 1, "『商品』（1列目）の左"),
    ("売上の前に区分の列を追加して", 2, "『売上』（2列目）の左"),
    ("原価と売上の右側に利益の列を追加して", 4, "『原価』と『売上』の右"),
])
def test_the_machine_resolves_the_position(task, at, shown):
    got, note = ailine.resolve_col_anchor(task, ["商品", "売上", "原価"])
    assert got == at, note
    assert shown in note, note


def test_no_position_words_means_no_answer_not_a_guess():
    """② 位置の言い回しが無い時は (None, None) ── 呼び側が「末尾」を選ぶ。
       ここで勝手に末尾を返すと、断れない場面との区別がつかなくなる。"""
    assert ailine.resolve_col_anchor("備考という列を追加して", ["商品", "売上"]) == (None, None)


def test_an_unknown_anchor_is_refused_by_name():
    got, note = ailine.resolve_col_anchor("すいかの右に列を追加して", ["商品", "売上"])
    assert got is None
    assert "『すいか』という列がありません" in note and "商品、売上" in note, note


# --- ③④ 関所 -------------------------------------------------------------------------

def test_a_named_column_at_a_relative_position(tmp_path):
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "ADD_COLUMN", {"name": "備考"}, META, task="原価の右に備考の列を追加して")
    assert ok, err
    assert resolved["_at_col"] == 4
    assert resolved["_name_label"] == "備考"


def test_a_column_without_a_name_is_allowed_and_disclosed(tmp_path):
    """③ 「列を追加して」しか言わない依頼 ── 入れるが、空であることを画面に書く。"""
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "ADD_COLUMN", {}, META, task="原価の右に列を追加して")
    assert ok, err
    assert resolved["name"] == ""
    assert "空のまま" in resolved["_name_label"], resolved["_name_label"]


def test_no_position_falls_back_to_the_end_with_a_stated_reason():
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "ADD_COLUMN", {"name": "備考"}, META, task="備考という列を追加して")
    assert ok, err
    assert resolved["_at_col"] == 4
    assert "末尾" in resolved["_at_basis"] and "位置の指定が無い" in resolved["_at_basis"]


@pytest.mark.parametrize("args,task,expect", [
    ({"name": "売上"}, "原価の右に売上の列を追加して", "既にあります"),
    ({"name": "備考"}, "すいかの右に備考の列を追加して", "という列がありません"),
])
def test_refusals(args, task, expect):
    ok, _r, _i, err = ailine.verify_dsl_args("ADD_COLUMN", dict(args), META, task=task)
    assert not ok, f"{args} を通した"
    assert expect in err, err


def test_the_request_is_recognised():
    assert ailine.task_asks_to_add_a_column("原価の右に列を追加して")
    assert ailine.task_asks_to_add_a_column("備考という列を足して")
    assert not ailine.task_asks_to_add_a_column("金額で降順に並べ替えて")
    assert not ailine.task_asks_to_add_a_column("原価の列を削除して")


# --- 生成 ------------------------------------------------------------------------------

def test_codegen_calls_the_helper_with_a_zero_based_index():
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "ADD_COLUMN", {"name": "備考"}, META, task="原価の右に備考の列を追加して")
    assert ok, err
    code = ailine.codegen_dsl("ADD_COLUMN", resolved, META)
    assert 'Call InsertColumnAt(oDoc, 3, "備考", 0)' in code, code


def test_the_helper_exists_in_the_bas():
    bas = (REPO / "src" / "ailine" / "helpers" / "AiLineHelpers.bas").read_text(encoding="utf-8")
    assert "Sub InsertColumnAt(" in bas
    assert "Columns.insertByIndex" in bas


# --- ⑤⑥ 事後条件 ----------------------------------------------------------------------

_F = [["商品", "売上", "原価", "利益"],
       ["りんご", 1200, 700, "=B2-C2"],
       ["みかん", 800, 300, "=B3-C3"]]


def _args(at=4, name="備考"):
    return {"name": name, "_at_col": at, "_headers": ["商品", "売上", "原価"], "_header_row": 1}


def test_a_correct_insert_passes(tmp_path):
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "備考"],
                              ["りんご", 1200, 700], ["みかん", 800, 300], ["ぶどう", 1500, 900]])
    status, reason = ailine.check_add_column(after, _args(), source_book=before)
    assert status == "pass", reason


def test_an_insert_in_the_middle_that_shifts_formulas_passes(tmp_path):
    """★ 列を挿すと右の列を指す式は参照が追随する（=B2-C2 → =B2-D2）。
       文字で比べると必ず食い違うので、**計算結果**で見る（追加・削除・入れ替えと同じ 1 箇所）。"""
    before = _book(tmp_path, _F, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "備考", "原価", "利益"],
                              ["りんご", 1200, None, 700, "=B2-D2"],
                              ["みかん", 800, None, 300, "=B3-D3"]])
    status, reason = ailine.check_add_column(
        after, {"name": "備考", "_at_col": 3, "_header_row": 1}, source_book=before)
    assert status == "pass", reason


def test_an_overwrite_instead_of_an_insert_fails(tmp_path):
    """★ 恒真殺し: 押し出さずに既存の列を潰したら落ちること。"""
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "備考"],
                              ["りんご", 1200, None], ["みかん", 800, None], ["ぶどう", 1500, None]])
    status, reason = ailine.check_add_column(
        after, {"name": "備考", "_at_col": 3, "_header_row": 1}, source_book=before)
    assert status == "fail", reason


def test_a_column_inserted_at_the_wrong_place_fails(tmp_path):
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "備考", "売上", "原価"],
                              ["りんご", None, 1200, 700], ["みかん", None, 800, 300],
                              ["ぶどう", None, 1500, 900]])
    status, reason = ailine.check_add_column(after, _args(at=4), source_book=before)
    assert status == "fail", f"宣言と違う位置を通した: {reason}"


def test_a_column_that_was_silently_filled_fails(tmp_path):
    """⑥ 空の列を作るはずが、勝手に値が入っていたら落ちる。"""
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "備考"],
                              ["りんご", 1200, 700, "?"], ["みかん", 800, 300, "?"],
                              ["ぶどう", 1500, 900, "?"]])
    status, reason = ailine.check_add_column(after, _args(), source_book=before)
    assert status == "fail" and "値が入っています" in reason, reason


def test_without_the_before_file_it_does_not_claim(tmp_path):
    after = _book(tmp_path, ROWS)
    status, _r = ailine.check_add_column(after, _args(), source_book=None)
    assert status == "warn"


def test_an_insert_that_disturbs_another_column_fails(tmp_path):
    """★ 恒真殺し（本命）: 列数も見出しも合っているのに、他の列の値が変わっていたら落ちる。
       ★ 上の overwrite の検体は**列数**で先に落ちるので、この形を別に置く
         （「落ちた」だけでは、どの番人が噛んだのかは分からない）。"""
    before = _book(tmp_path, ROWS, name="before.xlsx")
    after = _book(tmp_path, [["商品", "売上", "原価", "備考"],
                              ["りんご", 9999, 700], ["みかん", 800, 300], ["ぶどう", 1500, 900]])
    status, reason = ailine.check_add_column(after, _args(), source_book=before)
    assert status == "fail" and "他の列が変わっています" in reason, reason


def test_a_name_the_request_never_said_is_dropped():
    """★★ 実測（3 回中 1 回）: 名前を言っていない依頼に、LLM が「新しい列」という
       **依頼文に無い名前**を作って返した。A' 原則の違反 ── 値は LLM に確定させない。
       ★ 空欄は誤った名前より安い: 見出しが空なら △ になって人が気づく。もっともらしい
         名前が付くと、人は「自分がそう言った」と思ってしまう。"""
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "ADD_COLUMN", {"name": "新しい列"}, META, task="原価の右に列を追加して")
    assert ok, err
    assert resolved["name"] == "", resolved["name"]
    assert "採りませんでした" in resolved["_name_label"], resolved["_name_label"]


def test_a_name_the_request_did_say_is_kept():
    """★ 恒真殺し: 落とす側だけ強くして、正当な名前まで捨てていないこと。"""
    ok, resolved, _i, err = ailine.verify_dsl_args(
        "ADD_COLUMN", {"name": "備考"}, META, task="原価の右に備考の列を追加して")
    assert ok and resolved["name"] == "備考", (err, resolved.get("name"))


def test_the_trigger_does_not_eat_a_swap_request():
    """★★ 自分で開けた穴（実機の検体が捕まえた）: 「列を**入れ替え**て」の『入れ』に
       当たって、入れ替えの依頼が列追加として横取りされ、SWAP が動かなくなった。
       ★ 語の一部が別の語の一部でありうる ── 部分文字列の穴は、この repo で 2 度目。"""
    assert not ailine.task_asks_to_add_a_column("売上と原価の列を入れ替えて")
    assert ailine.task_asks_for_a_swap("売上と原価の列を入れ替えて")
    # ★ 恒真殺し: 塞いだ側で、正当な「入れて」まで殺していないこと
    assert ailine.task_asks_to_add_a_column("原価の右に空の列を入れて")
