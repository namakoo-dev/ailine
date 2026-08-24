# 検算できなかった列を「黙って落とさない」（2026-08-24・第二波 M6）。
#
# ★ 実測（盲検の実データ耐性レビュー・俺も実物で再現）:
#   金額列が `=B2*C2`（キャッシュ値なし）の請求書を stack すると
#     Σ数量: 元 30 / 出力 30
#     Σ単価: 元 600 / 出力 600
#     （★ Σ金額 の行が**出ない**）
#   `#REF!` / `#N/A` も同じ。**absence が唯一の信号**で、
#   「金額が検算されていない」ことに誰も気づけない。
#
# ★ 根は今日ずっと出ている家系と同じ ── **「無いこと」で伝えようとしている**。
#   分母が消える・列が消える・Σ の行が消える。全部「出ないこと」が信号になっている。
#   ★ 出ないものは読めない。**検算できなかったなら、できなかったと書く。**
#
# 契約:
#   ① 数値列として扱えなかった列は、**名指しで「検算できていない」と言う**
#   ② 誤爆しない: 最初から数値でない列（品名など）は言わない
#   ③ 全部検算できたときは何も言わない

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from ailine_core.multifile import unverified_numeric_columns  # noqa: E402


def test_formula_column_without_cache_is_named():
    """① 数式でキャッシュ値が無い列を名指しする。"""
    headers = ["品名", "数量", "単価", "金額"]
    verified = ["数量", "単価"]
    # 金額列は「数式は在るが値が無い」＝ 検算できなかった側
    got = unverified_numeric_columns(headers, verified, formula_columns=["金額"])
    assert got == ["金額"], got


def test_plain_text_columns_are_not_named():
    """② 誤爆防止: 最初から数値でない列は言わない。"""
    got = unverified_numeric_columns(["品名", "数量"], ["数量"], formula_columns=[])
    assert got == [], got


def test_nothing_when_all_verified():
    """③ 全部検算できたら黙る。"""
    got = unverified_numeric_columns(["数量", "金額"], ["数量", "金額"], formula_columns=["金額"])
    assert got == [], got
