# 名前で指した行の抜き出し ── 2026-08-30。Namakoo「特定条件の行や列の抜き出しができない」
#
# ★★ 実測（「丸和物流とみどり建設を抽出して」）── **抽出自体は成功していた**。
#   落ちたのは 2 つ:
#
#  ① 検算が**式と値を文字どおり比べていた**
#       × 出力1行目 F列が元と不一致（元 '=E2*1.1'（str） 出力 63360（int））
#     ★ 抽出が値を写すのは正しい ── 式をそのまま持っていけば新しいシートでは
#       違うセルを指す。比べる相手は**計算結果**でなければならない
#       （並べ替えの検算が既に取っている線と同じ）。
#     ★ 抽出と重複除去の**両方**が同じ形で比べていた ── 片方だけ直さない。
#
#  ② 一段目が**値ごとに 1 段ずつ**返していた
#       [EXTRACT value:丸和物流, EXTRACT value:みどり建設]
#     機械は各段で値を依頼文から取り直すので、解決後は同じ抽出が 2 段になり、
#     2 段目が連鎖の規則で**1 段目の出力を食って**落ちた（人は 1 回しか頼んでいない）。
#     ★ 「同じ仕事か」を見る時は、**機械が取り直す引数を外してから**比べる。
#     ★ 畳むのは**新しいシートを作る段**だけ ── 同じ並べ替えを 2 回のような段は
#       無害なので触らない（絞らずに入れたら既存の検体を壊しかけた）。
#
# ★ ついでにシート名: 『取引先丸和物流・みどり建設のどれか』は文になっておらず
#   会社名に見えた。一致（eq/in）は助詞が要る → 『取引先が丸和物流・みどり建設』。

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


# --- ① 検算は「見えている値」と比べる -----------------------------------------------------

def test_the_check_compares_what_is_shown_not_the_formula_text():
    """★ 変異試験: 元セルが式なら、比べる相手はキャッシュ値。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    assert src.count("def _row_as_shown(") == 1
    # 抽出と重複除去の両方がその関数を通ること（片配線を作らない）
    # 定義 1 + 呼び出し 2（抽出と重複除去）── 片方だけ直さない
    assert src.count("_row_as_shown(") == 3, "片方だけが計算結果と比べている"


def test_a_formula_row_is_read_as_its_value(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws.append(["名前", "金額", "税込"])
    ws.append(["丸和物流", 100, "=B2*1.1"])
    wb.save(p)
    # LibreOffice を通していないのでキャッシュ値は無い ── その時は None（黙る）
    from ailine_core.book_view import BookView
    with BookView(p) as bv:
        got = ailine._row_as_shown(bv, "S", 2, 3)
    assert got[0] == "丸和物流" and got[1] == 100
    assert got[2] != "=B2*1.1", "式の文字列をそのまま比べている"


# --- ② 同じ仕事の二重宣言を畳む ----------------------------------------------------------

def test_two_extracts_from_one_request_are_one_job():
    """★★ 一段目は値ごとに 1 段ずつ返す ── 機械が値を取り直すので同じ仕事になる。"""
    plan = [{"op": "EXTRACT", "args": {"col": "取引先", "cmp": "eq", "value": "丸和物流"}},
            {"op": "EXTRACT", "args": {"col": "取引先", "cmp": "eq", "value": "みどり建設"}}]
    folded, dropped = ailine.fold_identical_steps(plan)
    assert dropped == 1 and len(folded) == 1, folded


def test_a_different_condition_is_a_different_job():
    """★ 黙りすぎていないこと: 条件が違えば畳まない。"""
    plan = [{"op": "EXTRACT", "args": {"col": "取引先", "cmp": "eq", "value": "丸和物流"}},
            {"op": "EXTRACT", "args": {"col": "金額", "cmp": "gte", "value": 40000}}]
    _folded, dropped = ailine.fold_identical_steps(plan)
    assert dropped == 0


def test_steps_that_do_not_create_a_sheet_are_left_alone():
    """★★ 絞らずに入れたら、既存の検体（同じ並べ替えを 2 回）を壊しかけた。
       落ちたのは「新しいシートを作る段が、自分の出力を食う」形だけ。"""
    plan = [{"op": "SORT", "args": {"col": "金額", "order": "desc"}},
            {"op": "SORT", "args": {"col": "金額", "order": "desc"}}]
    folded, dropped = ailine.fold_identical_steps(plan)
    assert dropped == 0 and len(folded) == 2


def test_the_machine_derived_args_are_declared():
    """★ 「機械が取り直す引数」は表で宣言する（op が増えたら足せる形）。"""
    assert "value" in ailine.MACHINE_DERIVED_ARGS["EXTRACT"]


# --- ③ 出力シートの名前が日本語として読めること -------------------------------------------

@pytest.mark.parametrize("cmp, value, want", [
    ("eq", "丸和物流", "取引先が丸和物流"),
    ("in", ["丸和物流", "みどり建設"], "取引先が丸和物流・みどり建設"),
])
def test_a_match_needs_a_particle(cmp, value, want):
    assert ailine._extract_output_sheet_name("取引先", cmp, value) == want


def test_a_comparison_still_reads_as_before():
    """★ 大小比較は連結で日本語になっている ── そこは変えない。"""
    assert ailine._extract_output_sheet_name("金額", "gte", 40000) == "金額40000以上"
