# グラフ段（完成ロード: W10 の次）── 実装より先に凍結した赤い検体。
# 出典: spike 実測（2026-08-23・LO の壁ゼロ・chart XML は 3 種共通 XPath・
# tests/fixtures/charts/*.xlsx は spike が実 LO で生成した本物）。
#
# 契約:
#   ① kind の機械抽出（cmp と同じ作法）: 折れ線/推移→line・円/構成比/割合/内訳→pie・
#      棒→bar・手掛かりなし→None。LLM と食い違えば機械が勝つ+開示
#   ② 事後条件の恒真殺し: 「グラフ数 +1」だけでなく、chart XML の series 参照
#      （c:val/c:cat の c:f）が意図した値列/横軸列を指し、種別タグが kind と一致することまで
#      機械検証する。読むのは参照のみ（タイトル等の見た目要素に依存しない ── spike 事実 5）
#   ③ 検証は 3 種共通コード（spike 事実 4: XPath は kind 非依存）

import shutil
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402

FIXTURES = Path(__file__).parent / "fixtures" / "charts"

needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "extract_chart_kind_from_task"),
    reason="グラフ段 未実装（契約は凍結済み・実装が来たら自動実測化）",
    strict=True,
)


# --- ① kind の機械抽出 -----------------------------------------------------------

@needs_impl
@pytest.mark.parametrize("task,expected", [
    ("売上の推移を折れ線で見せて", "line"),
    ("月ごとの売上の推移をグラフにして", "line"),
    ("商品ごとの構成比を円グラフにして", "pie"),
    ("売上の内訳を円で見せて", "pie"),
    ("金額の割合をグラフにして", "pie"),
    ("商品ごとの金額を棒グラフにして", "bar"),
    ("金額をグラフにして", None),        # 手掛かりなし → 機械は断定しない（既定は呼び出し側）
])
def test_extract_chart_kind_from_task(task, expected):
    assert ailine.extract_chart_kind_from_task(task) == expected


# --- ② 事後条件の恒真殺し（実 LO 産の治具で縛る）--------------------------------

@needs_impl
@pytest.mark.parametrize("fixture,kind,value_col_letter", [
    ("bar.xlsx", "bar", "B"),
    ("line.xlsx", "line", "B"),
    ("pie.xlsx", "pie", "B"),
    ("noncontig.xlsx", "line", "C"),   # 横軸 A・値 C（B を飛ばす）── 非隣接も検証できる
])
def test_chart_series_verification_passes_on_real_lo_output(fixture, kind, value_col_letter):
    """実 LO が書いたブックで: 種別と値列の参照が一致 → pass。"""
    status, reason = ailine.check_chart_series(
        FIXTURES / fixture, kind=kind, value_col_letter=value_col_letter)
    assert status == "pass", reason


@needs_impl
def test_chart_series_verification_fails_on_wrong_column():
    """恒真殺し: グラフは在る（数は合う）が、意図した列（D）を描いていない → fail。
       旧 check_chart（数だけ）ならここで ✓ が出ていた。"""
    status, reason = ailine.check_chart_series(
        FIXTURES / "bar.xlsx", kind="bar", value_col_letter="D")
    assert status == "fail"
    assert "D" in reason or "参照" in reason


@needs_impl
def test_chart_series_verification_fails_on_wrong_kind():
    """種別違い: line を頼んだのに bar が刺さっている → fail（種別まで検証する）。"""
    status, reason = ailine.check_chart_series(
        FIXTURES / "bar.xlsx", kind="line", value_col_letter="B")
    assert status == "fail"


@needs_impl
def test_chart_series_verification_fails_when_no_chart(tmp_path):
    """グラフが無いブック → fail（読めない時に pass へ倒れない）。"""
    import openpyxl
    p = tmp_path / "plain.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["a", 1])
    wb.save(p)
    status, _reason = ailine.check_chart_series(p, kind="bar", value_col_letter="B")
    assert status == "fail"
