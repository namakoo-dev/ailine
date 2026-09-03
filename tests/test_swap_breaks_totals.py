# 列を入れ替えると合計式が二列にまたがる ── 2026-08-29。
# Namakoo「この操作が拒否されるのは正答か？」（デモ材料で「税込み金額と金額を入れ替えて」）
#
# ★★ 断りは**正しかった**。実測（LibreOffice で適用して読み戻した実物）:
#     入れ替え前  E9: =SUM(E2:INDEX(E:E,ROW()-1)) → 476,400
#                 F9: =SUM(F2:INDEX(F:F,ROW()-1)) → 524,040
#     入れ替え後  E9: =SUM(E2:INDEX(F:F,ROW()-1)) → 1,000,440   ★ E2 から F8 まで
#                 F9: =SUM(F2:INDEX(E:E,ROW()-1)) → 1,000,440   ★ 二列ぶん足している
#   データ側は正しく追従していた（税込み列は `=F2*1.1` になった）。**合計式だけ**、
#   範囲の片側しか動かなかった。画面には大きな数字が出るだけなので人は気づかない。
#
# ★ ただし**理由の文が行の話をしていた** ──「追加/削除した行を参照する合計式なら
#   正当ですが」。列の入れ替えでその説明は出てはいけない。
#   ★ 正しい判定を、間違った言葉で説明していた。理由は呼ぶ側が渡す形にした
#     （並べ方は _moved_rows_note の 1 箇所のまま ── 文面を写し取らない）。

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from _product_source import count_in_product, window_around  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む


def test_the_default_reason_still_talks_about_rows():
    """★ 行の挿入・削除で使う既定の理由は変えていない（そこでは行の話が正しい）。"""
    note = ailine._moved_rows_note(["8 行目の 5 列目（式の結果 1→2）"])
    assert "追加/削除した行" in note
    assert "8 行目の 5 列目" in note


def test_the_swap_reason_talks_about_columns_not_rows():
    """★★ 列の入れ替えなのに行の話をしていた ── そこを塞ぐ。"""
    note = ailine._moved_rows_note(["8 行目の 5 列目（式の結果 1→2）"],
                                    why="式が入れ替え先の列に付いていきませんでした")
    assert "追加/削除した行" not in note, note
    assert "列に付いていきません" in note, note


def test_the_swap_check_uses_the_column_reason():
    """★ 変異試験: 入れ替えの断り文に、行の理由が混ざらないこと。"""
    seg = window_around("を入れ替えたあと、式の計算結果が変わっています", after=500)
    assert "why=" in seg, "入れ替えが既定の（行の）理由をそのまま使っている"
    assert "二列にまたがる合計" in seg, seg[:300]


def test_the_note_is_still_assembled_in_one_place():
    """★ 理由を差し替えられるようにしたせいで、並べ方が写し取られていないこと。"""
    assert count_in_product("（ほか {len(disclosures) - 5} 件）") == 1
