"""陽性対照（すぐ消す）── この道具が「死んでいる番人」を報告できるかを確かめる検体。"""
from _product_source import count_in_product


def test_a_guard_that_does_not_actually_look():
    # ★ 常に真。文言が消えても緑のまま ── 道具はこれを「死んでいる」と言うべき。
    assert count_in_product("列全体は勝手に書き換えません") >= 0
