# 2 列の計算の演算子は依頼文が決める（2026-10-01・依頼の項の台帳で D だった項目）。
#
# ★ 形: 「売上から原価を引いた利益の列を作って」に LLM が `+` を返すと、事後条件は
#   宣言どおりの式（売上+原価）を確かめて ✓ を出す。記号の読み手（arith.py）は `A÷B` の
#   形しか読まないので、語で書いた依頼は誰も見ていなかった。
#
# 契約:
#   ① 依頼文の語・記号から演算子が 1 つに読めて LLM と食い違えば、依頼文が勝つ
#   ② 置き換えたことは解釈行に出る（出典「依頼文: 『引い』」・警告ではない）
#   ③ 一致する時・読めない時（語が無い・2 種類の演算が読める）は何も変えない
#   ④ 列の名前の中の語（『値引き』の引）は数えない
#   ⑤ 引き算に直した回も、向きの関所（どちらから引くか）はそのまま効く
#   ⑥ 見出しが 1 行目でない表でも同じ
#   ⑦ 列 1 つ × 率（税込み）の枝には触らない

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import arith, argcheck  # noqa: E402

HEAD = ["商品", "売上", "原価", "数量", "単価", "値引き"]


def _book(tmp_path, header_row=1):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for i in range(header_row - 1):
        ws.append(["売上表"] if i == 0 else [])
    ws.append(HEAD)
    ws.append(["りんご", 1000, 600, 3, 100, 50])
    ws.append(["みかん", 2000, 1500, 5, 80, 0])
    wb.save(p)
    return p


def _meta(path, header_row=1):
    return {"sheets": ["売上"], "headers": {"売上": list(HEAD)},
            "header_rows": {"売上": header_row}, "path": str(path)}


def _compute(tmp_path, task, operands, operator, header_row=1, expect_ok=True):
    ok, r, _i, err = ailine.verify_dsl_args(
        "COMPUTE_COLUMN", {"operands": list(operands), "operator": operator},
        _meta(_book(tmp_path, header_row), header_row), task=task)
    assert ok is expect_ok, err
    return r if ok else err


# --- 読み手（純関数）-------------------------------------------------------------------

def test_the_reader_reads_words_and_signs():
    assert arith.read_operator("売上から原価を引いた利益の列", HEAD)[0] == "-"
    assert arith.read_operator("数量と単価を掛けた金額の列", HEAD)[0] == "*"
    assert arith.read_operator("売上と原価を足した列", HEAD)[0] == "+"
    assert arith.read_operator("売上を原価で割った列", HEAD)[0] == "/"
    assert arith.read_operator("数量×単価の列", HEAD) == ("*", "×")


def test_the_reader_is_silent_when_it_cannot_decide():
    for t in ("売上と原価から新しい列を作って",          # 語が無い
              "売上から原価を引いた利益率の列",          # - と /
              "在庫が不足している数量と単価の列",        # 不足の足は数えない
              "売上と原価を比較した列"):                  # 比較の比は数えない
        assert arith.read_operator(t, HEAD) is None, t


# --- ①② 食い違えば依頼文が勝つ ------------------------------------------------------------

def test_subtract_word_overrides_a_plus(tmp_path):
    r = _compute(tmp_path, "売上から原価を引いた利益の列を作って", ["売上", "原価"], "+")
    assert r["operator"] == "-" and r["operands"] == ["売上", "原価"]
    assert r["_sources"]["operator"] == "依頼文: 『引い』"


def test_multiply_word_overrides_a_plus(tmp_path):
    r = _compute(tmp_path, "数量と単価を掛けた金額の列を作って", ["数量", "単価"], "+")
    assert r["operator"] == "*"


def test_the_disclosure_reaches_the_interpretation_line(tmp_path):
    r = _compute(tmp_path, "数量と単価を掛けた金額の列を作って", ["数量", "単価"], "+")
    line = ailine.format_confirmation_line("COMPUTE_COLUMN", r, set())
    assert "演算子:*（依頼文: 『掛け』）" in line


# --- ③ 何も変えない ----------------------------------------------------------------------

def test_agreeing_answer_is_left_alone(tmp_path):
    r = _compute(tmp_path, "数量と単価を掛けた金額の列を作って", ["数量", "単価"], "*")
    assert r["operator"] == "*" and "operator" not in r.get("_sources", {})


def test_no_word_leaves_the_answer(tmp_path):
    r = _compute(tmp_path, "数量と単価から金額の列を作って", ["数量", "単価"], "+")
    assert r["operator"] == "+" and "operator" not in r.get("_sources", {})


def test_a_word_about_something_else_is_not_read(tmp_path):
    """★ 実走行から: 依頼文が 2 列を名指ししていない回は、語は 2 列の関係を言っていない
    （「合計を出して」→ 数量×単価、「残業時間を加算」→ 退勤−出勤）。読むと正しい計算を壊す。"""
    r = _compute(tmp_path, "合計を出して", ["数量", "単価"], "*")
    assert r["operator"] == "*" and "operator" not in r.get("_sources", {})


def test_two_operations_leave_the_answer(tmp_path):
    r = _compute(tmp_path, "数量と単価を掛けて合計した金額の列を作って", ["数量", "単価"], "*")
    assert r["operator"] == "*" and "operator" not in r.get("_sources", {})


# --- ④ 列の名前の中の語は数えない --------------------------------------------------------

def test_a_word_inside_a_column_name_is_not_read(tmp_path):
    """『値引き』の引を引き算と読まない（読むと - と + で決まらず、足し算の依頼が素通り）。"""
    r = _compute(tmp_path, "売上と値引きを足した列を作って", ["売上", "値引き"], "-")
    assert r["operator"] == "+"


# --- ⑤ 向きの関所は壊さない ---------------------------------------------------------------

def test_direction_gate_still_asks_after_the_fix(tmp_path):
    """「差」は引き算だが向きを言っていない ── 直した後も、どちらから引くかを問う。"""
    err = _compute(tmp_path, "売上と原価の差の列を作って", ["売上", "原価"], "+", expect_ok=False)
    assert "どちらから引く" in err


def test_direction_is_taken_from_the_request_after_the_fix(tmp_path):
    r = _compute(tmp_path, "原価を売上から引いた列を作って", ["原価", "売上"], "+")
    assert r["operator"] == "-" and r["operands"] == ["売上", "原価"]


# --- ⑥ 見出しが 1 行目でない表 -----------------------------------------------------------

def test_header_not_on_the_first_row(tmp_path):
    r = _compute(tmp_path, "売上から原価を引いた利益の列を作って", ["売上", "原価"], "*",
                 header_row=3)
    assert r["operator"] == "-"


# --- ⑦ 列 1 つの枝には触らない -------------------------------------------------------------

def test_single_column_rate_branch_is_untouched(tmp_path):
    r = _compute(tmp_path, "売上に消費税10%を掛けた税込みの列を作って", ["売上"], "*")
    assert r["operator"] == "*" and "operator" not in r.get("_sources", {})


# --- 変異: 配線を外すと直らない -----------------------------------------------------------

def test_without_the_wiring_the_llm_operator_would_win(tmp_path, monkeypatch):
    monkeypatch.setattr(argcheck.arith_request, "read_operator", lambda task, names=(): None)
    r = _compute(tmp_path, "売上から原価を引いた利益の列を作って", ["売上", "原価"], "+")
    assert r["operator"] == "+"      # ← 直す前の挙動（嘘の ✓ の材料）
