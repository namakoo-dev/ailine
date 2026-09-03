# 新しい列の名前と、行×列のピンポイント指定 ── 2026-08-30。
# Namakoo「税込み金額が式で表示されるようになってしまった。あと「」で囲んだ操作は
# セルの値代入に統一したほうがいいかもしれないな。あと行と列による一意の指定も
# 出来た方がいい。ピンポイントに操作できるようになる」
#
# ★★ 実測（下書きに列が 2 本できていた）:
#     「金額の右に税込み金額を追加」を 2 回頼んで、見出しが
#       1 回目「税込金額」（**「み」が落ちた**）／2 回目「金額*1.1」（**式が名前になった**）
#   ★ 前者は道具が `f"税込{列名}"` と**作った**名前、後者は式そのもの。
#     人が書いた名前がそこに在るのに、機械が別の名前を発明していた（A' 原則が抜けていた）。
#   ★ しかも**解釈行に名前が出ていなかった**ので、気づく手がかりが無かった。
#
# ★★ もう 1 つ、もっと危ない形:
#     「A行G列を『税込み金額』に上書き」→ **一括書換 対象列:金額**（金額列を丸ごと文字で潰す）
#   ★ LLM が実在しない列『税込み金額』を返し、救済が**別の列『金額』を採用した**。
#     『金額』が依頼文に現れるのは**引用符の中（＝書き込む値）だけ**だった。
#   ★ 引用符の中は**値**であって、対象の名指しではない ── 証拠に使わない。

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from _product_source import window_around  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む

H = ["取引先", "項目", "件数", "単価", "金額", "締め日", "担当"]


# --- ① 新しい列の名前は依頼文から取る ---------------------------------------------------

@pytest.mark.parametrize("task, want", [
    ("金額の右に税込み金額を追加", "税込み金額"),
    ("金額の右に税込み金額列を追加", "税込み金額"),
    ("原価の右にチェックという列を追加して", "チェック"),
    ("単価と金額の右に粗利の列を作って", "粗利"),
])
def test_the_new_column_takes_its_name_from_the_request(task, want):
    assert ailine.new_column_name_from_task(task, H) == want


@pytest.mark.parametrize("task", [
    "金額の右に列を追加して",                    # 名前が書かれていない
    "A行G列を「税込み金額」に上書き",            # 引用符の中は値
    "金額の右に金額を追加",                      # 既にある列の名前
])
def test_it_does_not_invent_a_name(task):
    assert ailine.new_column_name_from_task(task, H) is None


def test_the_name_reaches_the_interpretation_line():
    """★ 名前が出ていなかったので、「み」が落ちても誰も気づけなかった。"""
    keys = {k for _label, k, _fn in ailine._CONFIRM_FIELDS["COMPUTE_COLUMN"]}
    assert "_new_col_label" in keys, "解釈行に『新しい列の名前』が無い"


def test_the_request_beats_the_invented_label():
    """★★ 変異試験: 依頼文の名前を、道具が作る「税込〜」より先に見ること。"""
    seg = window_around('if not resolved.get("target"):', after=900)
    j_asked = seg.index("new_column_name_from_task")
    j_tax = seg.index("_TAX_INCLUSIVE_RE")
    assert j_asked < j_tax, "作った名前のほうが先に当たっている"


# --- ② 引用符の中は値であって、対象の名指しではない --------------------------------------

def test_quoted_text_is_masked_before_looking_for_a_target():
    assert ailine._task_outside_quotes("A行G列を「税込み金額」に上書き") == "A行G列を       に上書き"


def test_a_column_is_not_chosen_from_inside_the_quotes():
    """★★ ここが緩むと、金額列を丸ごと文字列で潰す（実測でその一歩手前だった）。"""
    assert ailine._task_names_single_real_column("A行G列を「税込み金額」に上書き", H) is None


def test_a_column_named_outside_the_quotes_still_wins():
    """★ 黙りすぎていないこと: 引用符の外で名指しされた列は今までどおり拾う。"""
    assert ailine._task_names_single_real_column("担当を「佐藤」にして", H) == "担当"


# --- ③ 見出しの名前を、行×列で名指しして変えられる ---------------------------------------

def _meta(path):
    return {"sheets": ["請求"], "headers": {"請求": list(H)},
            "header_rows": {"請求": 1}, "path": str(path)}


@pytest.fixture()
def book(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(H)
    ws.append(["丸和物流", "配送", 12, 4800, 57600, "2026/08/31", "田中"])
    wb.save(p)
    return p


def test_the_header_row_can_be_written_when_named_by_coordinates(book):
    """★★ Namakoo「行と列による一意の指定も出来た方がいい。ピンポイントに操作できる」
       ── それまで見出し行は一律で断っていたので、**列の名前を直す手段が 1 つも
       無かった**（計算列の見出しが「金額*1.1」に化けた表を、人が直せない）。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", {"row_number": 1, "col": "F", "value": "税込金額"},
        _meta(book), task="1行F列を「税込金額」にして")
    assert ok, err
    assert r.get("_writes_header") is True
    assert r["_row_index"] == 1 and r["_col_index"] == 6
    assert "見出しの名前を変えます" in r.get("_at_basis", ""), r.get("_at_basis")


def test_the_header_is_not_written_on_a_guess(book):
    """★ 開いたのは**人が行番号で名指しした時だけ** ── LLM の推しでは開けない。"""
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", {"row_number": 1, "col": "F", "value": "税込金額"},
        _meta(book), task="締め日を税込金額にして")     # 依頼文に行番号が無い
    assert not ok
    assert "見出し行" in err, err


def test_the_column_letter_in_the_request_beats_the_llm(book):
    """★★ 実測: 第二段は col に**書き込む値**を入れてきた（col=『税込金額(10%)』）。
       依頼文が英字で列を名指ししているなら、それが正（行番号と同じ分担）。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "SET_CELL_VALUE", {"row_number": 1, "col": "税込金額(10%)", "value": "税込金額(10%)"},
        _meta(book), task="1行F列を「税込金額(10%)」にして")
    assert ok, err
    assert r["_col_index"] == 6, r.get("_col_index")


def test_the_check_looks_at_the_coordinate_not_the_old_name():
    """★ 見出しを書き換えると、その列は**元の名前で引けなくなる**。
       実測で「列『税込み金額』が見つからない」と落ちた（書き込みは成功していたのに）。"""
    seg = window_around("def check_set_cell_value(", after=3000)
    assert '_writes_header' in seg and '_col_index' in seg, "検算が名前で引いたまま"
    assert "_scan_from" in seg, "見出し行が「変わったセル」の数え上げに入っていない"
