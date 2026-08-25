# 盲検 2 回目 R1（2026-08-25）── **今日の俺の直しが入れた回帰**。
#
# ★ 実測: `_sort_rows_lost_their_identity`（7e0befe・同じ日の午後）は行の中身を
#   **生値の多重集合**で比べていた。相対参照の式は行が動けば `=B2*C2` → `=B5*C5` と
#   **変わるのが正しい**のに、それを「ちぎれた」と判定していた。
#   → **この製品の看板ユースケース**「金額列を作って金額順に並べる」が必ず落ちた。
#      出力は完全に正しい（1000/1000/800/300 の降順・式も各行に追随）のに、
#      検算だけが間違って exit 1・全ロールバック。
#   ★ 原因表示も「見出しの無い列が一緒に動かなかった可能性」と**誤診断**していた。
#
# 契約:
#   ① 式の列があっても、正しい並べ替えは通る（看板ユースケース）
#   ② 式を外したせいで**恒真にしない** ── 本当にちぎれた行は、式が在っても掴む
#   ③ 式が「行と一緒に動いた」ことは要求しない（動くのが正しい）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _mk(tmp_path, name, rows):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p


def test_the_flagship_case_passes(tmp_path):
    """① 金額列（式）を作って金額順に並べる ── これが落ちてはいけない。"""
    before = _mk(tmp_path, "before.xlsx",
                 [["商品", "数量", "単価", "金額"],
                  ["りんご", 3, 100, "=B2*C2"], ["みかん", 10, 80, "=B3*C3"],
                  ["ぶどう", 2, 500, "=B4*C4"], ["かき", 5, 200, "=B5*C5"]])
    after = _mk(tmp_path, "after.xlsx",
                [["商品", "数量", "単価", "金額"],
                 ["ぶどう", 2, 500, "=B2*C2"], ["かき", 5, 200, "=B3*C3"],
                 ["りんご", 3, 100, "=B4*C4"], ["みかん", 10, 80, "=B5*C5"]])
    status, reason = ailine.check_sort(after, {"col": "単価", "order": "desc"},
                                        source_book=before)
    assert status == "pass", f"正しい並べ替えを落とした（看板ユースケース）: {reason}"


def test_torn_rows_are_still_caught_when_formulas_exist(tmp_path):
    """② 恒真にしない ── 式が在っても、本当にちぎれた行は掴む。"""
    before = _mk(tmp_path, "before.xlsx",
                 [["商品", "単価", None, None],
                  ["りんご", 120, "=B2*2", "特売"], ["ぶどう", 500, "=B3*2", "高級"]])
    after = _mk(tmp_path, "after.xlsx",
                [["商品", "単価", None, None],
                 ["ぶどう", 500, "=B2*2", "特売"],   # 備考が置き去り（本当にちぎれた）
                 ["りんご", 120, "=B3*2", "高級"]])
    status, reason = ailine.check_sort(after, {"col": "単価", "order": "desc"},
                                        source_book=before)
    assert status == "fail", f"式のせいで見逃した（恒真）: {reason}"
    assert "ちぎれ" in reason


def test_formula_text_change_alone_is_not_a_tear(tmp_path):
    """③ 式が動いただけ（他は完全に一致）は、ちぎれではない。"""
    before = _mk(tmp_path, "before.xlsx",
                 [["商品", "単価", "計"], ["a", 100, "=B2*2"], ["b", 500, "=B3*2"]])
    after = _mk(tmp_path, "after.xlsx",
                [["商品", "単価", "計"], ["b", 500, "=B2*2"], ["a", 100, "=B3*2"]])
    status, reason = ailine.check_sort(after, {"col": "単価", "order": "desc"},
                                        source_book=before)
    assert status == "pass", reason
