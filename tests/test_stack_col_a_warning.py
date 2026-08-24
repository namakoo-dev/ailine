# stack の「1列目 vs 表の範囲」の食い違い警告 ── 誤爆を止める（2026-08-24）。
#
# ★ 出所（盲検の査定・2 回目）: 小計行のある請求書 3 冊**すべて**に
#     ⚠ 0703_あかつき商事.xlsx: 1列目から数えると 2 行ですが、表の範囲は 3 行あります
#   が出た。原因は**自分が正しく除外した小計行**。小計行の無い 1 冊だけ無警告だったので
#   因果は確定している。日本の請求書は「小計」を金額の隣（右寄せ）に書き、1 列目は空にする
#   のが最も普通の形なので、**この警告は普通の請求書で必ず鳴る**。
#
# ★ この文言は今朝（同じ日）俺が「開発者用語が読めない」という指摘を受けて書き換えた所。
#   **文言だけ直して根を残した。** オオカミ少年防止を謳う道具が、自分でオオカミ少年になっていた。
#
# 契約:
#   ① 1列目の非空行数と表の範囲の差が、**合計行として除外した行で説明できる**なら黙る
#   ② 説明できない差（本物のデータ行が 1 列目空欄で落ちている等）は今までどおり言う

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from ailine_core.stack import col_a_mismatch_is_explained  # noqa: E402


def test_difference_explained_by_excluded_total_rows_is_silent():
    """① 3 行のうち 1 行が合計行（1列目が空）── 差 1 は除外 1 で説明できる。黙る。"""
    assert col_a_mismatch_is_explained(col_a_count=2, used_range_count=3,
                                        excluded_blank_label_rows=1) is True


def test_unexplained_difference_still_warns():
    """② 差 2 に対して除外が 1 しか無い ── 説明できない 1 行がある。言う。"""
    assert col_a_mismatch_is_explained(col_a_count=1, used_range_count=3,
                                        excluded_blank_label_rows=1) is False


def test_no_difference_is_silent():
    assert col_a_mismatch_is_explained(col_a_count=3, used_range_count=3,
                                        excluded_blank_label_rows=0) is True


def test_excluded_rows_that_had_a_label_do_not_explain_the_gap():
    """★ 恒真殺し: 1列目にラベルの在る合計行は col_a_count に数えられているので、
       差の説明にはならない（説明に使えるのは**1列目が空の**除外行だけ）。"""
    assert col_a_mismatch_is_explained(col_a_count=2, used_range_count=3,
                                        excluded_blank_label_rows=0) is False
