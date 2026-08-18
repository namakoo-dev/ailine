"""ailine_core/sum_identity.py 単体の検体（算術恒等の検算そのもの）。

★ ここで測るのは「語を読まないこと」と「位置を返すこと」の2点。
ヘルパは数値の並びしか受け取らないので、辞書・書式・行の型・言語のどれにも依存しない
―― その独立性を、番人テスト（下の test_module_reads_no_words / _is_portable）で機械化する。
"""
import ast
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from ailine_core.sum_identity import rows_matching_sum_above  # noqa: E402

MODULE_PATH = REPO_ROOT / "ailine_core" / "sum_identity.py"


def _rows(hits):
    return [h.row for h in hits]


# --- 恒等式そのもの ---------------------------------------------------------

def test_row_equal_to_the_sum_of_every_row_above_is_found():
    hits = rows_matching_sum_above([(2, 100), (3, 200), (4, 300)])
    assert _rows(hits) == [4]
    assert hits[0].value == 300
    assert hits[0].term_rows == (2, 3)
    assert hits[0].is_last is True


def test_hit_in_the_middle_is_reported_as_not_last():
    # 既存の合計 300 の下に、それを含めて足した 600 が来た形（＝二重計上）。
    hits = rows_matching_sum_above([(2, 100), (3, 200), (4, 300), (5, 600)])
    assert _rows(hits) == [4, 5]
    assert hits[0].is_last is False
    assert hits[1].is_last is True


def test_only_one_row_above_never_matches():
    """★ 上が1つだけなら「2 行目が 1 行目と等しい」だけで当たってしまう。F6 の線。"""
    assert rows_matching_sum_above([(2, 100), (3, 100)]) == []
    assert rows_matching_sum_above([(2, 100), (3, 100), (4, 200)])[0].row == 4


def test_all_zero_column_is_silent():
    """0 == 0+0 は恒真。合計とは呼べないので数えない。"""
    assert rows_matching_sum_above([(2, 0), (3, 0), (4, 0)]) == []


def test_non_numeric_rows_are_skipped_not_counted_as_zero():
    hits = rows_matching_sum_above([(2, 100), (3, "小計"), (4, 200), (5, None), (6, 300)])
    assert _rows(hits) == [6]
    assert hits[0].term_rows == (2, 4)          # 文字列と None は項に入らない
    assert hits[0].is_last is True


def test_booleans_are_not_numbers():
    assert rows_matching_sum_above([(2, True), (3, True), (4, 2)]) == []


def test_float_rounding_is_tolerated():
    hits = rows_matching_sum_above([(2, 0.1), (3, 0.2), (4, 0.30000000000000004)])
    assert _rows(hits) == [4]


def test_running_total_column_does_not_match():
    """F7: 累計列（100/300/450/750）は上の全部の和にはならない。"""
    assert rows_matching_sum_above([(2, 100), (3, 300), (4, 450), (5, 750)]) == []


def test_two_blocks_only_misfire_when_read_as_one():
    """F9 の対比: 続けて読むと当たる並びが、塊を切れば当たらない。"""
    merged = [(2, 1200), (3, 3400), (5, 4600), (6, 800)]
    assert _rows(rows_matching_sum_above(merged)) == [5]      # 続けて読むと誤爆する
    assert rows_matching_sum_above(merged[:2]) == []          # 上の塊だけなら黙る


def test_empty_and_single_row_inputs_are_silent():
    assert rows_matching_sum_above([]) == []
    assert rows_matching_sum_above([(2, 5)]) == []


# --- 番人（この関数が語も書式も読まないこと） --------------------------------

# --- 「足し込んだ範囲の最終行だけ」という絞り（_nested_total_reason の線） ----------
#   ヘルパは一致する行を全部返す。どれを鳴らすかを決めるのは呼び出し側の**位置の条件**。
#   ここはその条件そのものを凍結する（実測で 100/200/300 型の誤爆が出て絞った回）。

def _reason(values):
    import ailine
    return ailine._nested_total_reason(values, "Sheet", 2)


def test_hit_at_the_last_row_of_the_summed_range_fires():
    """既存の合計 300 の下に、それを含めて足した 600 が来た形（T6 と同じ並び）。"""
    reason = _reason([(2, 100), (3, 200), (4, 300), (5, 600)])
    assert reason is not None
    assert "B4" in reason and "300" in reason and "B2:B3" in reason


def test_hit_in_the_middle_of_the_summed_range_stays_silent():
    """★ demo/sales.xlsx（100,200,300,400,500,250）の 300 は開発部門のただの売上。

    恒等式としては当たるが、足し込んだ範囲の**真ん中**なので鳴らさない ―― 既存の合計は
    その塊の一番下に在る。ここが実測で誤爆した唯一の形（README の quickstart）。
    """
    values = [(r, v) for r, v in zip(range(2, 8), [100, 200, 300, 400, 500, 250])]
    assert _reason(values + [(8, 1750)]) is None


def test_a_total_that_is_not_the_last_row_of_the_range_is_missed_on_purpose():
    """★ 絞った代償を明示的に凍結する（取り逃がし側に倒したことを隠さない）。

    『本体 → 小計 → 本体 → 小計 → 合計』のように、既にある合計が範囲の最終行**でない**
    位置にある帳票は、二重に数えていても鳴らない。これは既知の穴であって偶然ではない。
    """
    # 100,200 → 小計300 / 400,500 → 小計900 / 全部足した 2400（300 と 900 が二重計上）
    values = [(2, 100), (3, 200), (4, 300), (5, 400), (6, 500), (7, 900), (8, 2400)]
    assert _reason(values) is None


def test_module_reads_no_words():
    """★ 辞書禁止の機械化: docstring 以外の文字列リテラルが1つも無いこと。

    語で行を見分ける実装は、必ずどこかで文字列と突き合わせる。突き合わせる相手が
    存在しない（説明文以外に文字列が無い）なら、この関数は語を読みようがない。
    """
    tree = ast.parse(MODULE_PATH.read_text(encoding="utf-8"))
    docstrings = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            doc = node.body[0] if node.body else None
            if isinstance(doc, ast.Expr) and isinstance(doc.value, ast.Constant) \
                    and isinstance(doc.value.value, str):
                docstrings.add(id(doc.value))
    literals = [n.value for n in ast.walk(tree)
                if isinstance(n, ast.Constant) and isinstance(n.value, str)
                and id(n) not in docstrings]
    assert literals == [], f"語と突き合わせる文字列が混ざっている: {literals}"


def test_module_is_portable_and_needs_no_spreadsheet_library():
    """★ 言語非依存＝表計算ライブラリにも依存しない（標準ライブラリだけで閉じる）。"""
    tree = ast.parse(MODULE_PATH.read_text(encoding="utf-8"))
    roots = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            roots.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom):
            if node.level == 0 and node.module:
                roots.add(node.module.split(".")[0])
    assert roots <= {"__future__", "dataclasses", "typing"}, roots
