"""「見出し」の語の名簿は 1 本（2026-09-24・Namakoo 決裁 A: 見出し|ヘッダー|ヘッダ）。

★ なぜ在るか: 本体は『見出し』だけ、ailine_core/subject は『見出し|ヘッダー|ヘッダ』と、同じ概念で
  名簿を 2 つ持っていた。「ヘッダーを太字にして」は主語の照合では見出し行に届くのに、前段で作った列を
  太字にしかけている時の助言（_maybe_warn_header_col_mismatch）は黙っていた。
★ 1 本であることは tests/test_same_name_is_the_same_object.py が守る（本体の名前は subject の物と同じ実体）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import ailine  # noqa: E402
from ailine_core import subject  # noqa: E402


def _warn(task):
    return ailine._maybe_warn_header_col_mismatch(
        "BOLD", {"target": "col:数量*単価"}, ["数量*単価"], task)


def test_header_in_katakana_also_raises_the_advisory():
    for word in ("ヘッダー", "ヘッダ"):
        w = _warn(f"{word}を太字にして")
        assert w, f"「{word}」で助言が出ない（名簿が狭い方に戻っている）"
        assert f"「{word}」" in w, f"助言が依頼で使われた語を引用していない: {w}"


def test_the_original_word_still_works():
    w = _warn("見出しを太字にして")
    assert w and "「見出し」" in w, w


def test_no_header_word_no_advisory():
    assert _warn("数量*単価を太字にして") is None


def test_the_main_module_uses_the_one_list():
    assert ailine._HEADER_WORD_RE is subject._HEADER_WORD_RE
