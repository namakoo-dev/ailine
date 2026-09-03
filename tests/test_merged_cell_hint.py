# 土台固め（2026-08-24）── 事後条件が破れたときの「心当たり」。
#
# ★ 飾りの生存表（scripts/fidelity_matrix.py）を作っている最中に、**対照実験**で確定した:
#   表の範囲に結合セルがあると LibreOffice の並べ替えが黙って何もしない。
#   同じブックから結合セルだけ外すと ✓ になる（実測・exit 1 → exit 0）。
#   ailine は嘘の ✓ を出さずに落ちるので「壊さない」は守れているが、**理由を言わない**
#   ので使う側はそこで詰まる ── 今日ずっと直してきた形。
#
# 契約:
#   ① 結合セルが在れば、件数と位置を名指しして次の一手を言う
#   ② 断定しない（「原因はこれ」ではなく「心当たり」）── 結合セルが在っても効く操作はある
#   ③ 結合セルが無ければ 1 行も出さない（誤爆しない）
#   ④ 読めないブックでも例外にならない
#   ⑤ 3 つの失敗経路すべてが同じ関数を呼ぶ（書き写さない）

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from _product_source import count_in_product  # noqa: E402 ── ★ 番人は本体決め打ちでなく製品コード全体を読む


def _book(tmp_path, merge=None):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["取引先", "金額"])
    ws.append(["あかつき商事", 12000])
    ws.append(["みどり工業", 8500])
    if merge:
        ws.merge_cells(merge)
    wb.save(p)
    return p


def test_names_the_merged_cells(tmp_path):
    lines = ailine.likely_cause_of_no_change(_book(tmp_path, "A4:B4"), "売上")
    text = "\n".join(lines)
    assert "結合セル" in text and "A4:B4" in text, text
    assert "1 件" in text, text
    assert "解除" in text, f"次の一手が無い: {text}"


def test_hedges_instead_of_asserting_the_cause(tmp_path):
    """② 断定しない ── 結合セルが在っても効く操作はある。"""
    text = "\n".join(ailine.likely_cause_of_no_change(_book(tmp_path, "A4:B4"), "売上"))
    assert "心当たり" in text, text
    assert "原因は" not in text, f"断定している: {text}"


def test_silent_without_merged_cells(tmp_path):
    assert ailine.likely_cause_of_no_change(_book(tmp_path), "売上") == []


def test_unreadable_book_is_silent(tmp_path):
    p = tmp_path / "broken.xlsx"
    p.write_bytes(b"not a zip")
    assert ailine.likely_cause_of_no_change(p, "売上") == []


def test_all_failure_paths_share_one_implementation():
    """⑤ 3 つの失敗経路が同じ関数を呼ぶ（実測: この形は今日 3 回片配線を生んだ）。"""
    assert count_in_product("適用されたが事後条件を満たさない") == 3, "経路の数が変わった（検体の前提が古い）"
    assert count_in_product("likely_cause_of_no_change(") == 4, \
        "失敗経路の一部が心当たりを言わない（片配線）"
