"""「語が独立して出ているか」の判定と漢字の範囲は 1 本（2026-09-24）。

★ なぜ在るか: 同じ判定が argcheck と alias_store に 2 つあり、写しの方の漢字の範囲だけが、互換漢字の
  「豈」（U+F900）の代わりに見た目が同じ統合漢字の「豈」（U+8C48）で始まっていた ── ハングルや私用領域まで
  漢字と数え、その隣の語を「語の内部」と誤って読んだ。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from ailine_core import alias_store, argcheck, word_boundary  # noqa: E402

CASES = [
    ("売上", "売上を太字にして", True),
    ("売上", "総売上を太字にして", False),          # 漢字の内部
    ("売上", "売上高の列", False),
    ("売上", "売上" + chr(0xAC00) + "を", True),      # ハングルの隣は漢字の内部ではない
    ("売上", "売上" + chr(0xF900) + "を", False),     # 互換漢字は漢字
    ("売上", "売上" + chr(0xE000) + "を", True),      # 私用領域は漢字ではない
    ("売上", "", False),
]


def test_the_range_is_written_by_codepoint():
    want = "[" + "".join(chr(a) + "-" + chr(b) for a, b in
                         ((0x3400, 0x4DBF), (0x4E00, 0x9FFF), (0xF900, 0xFAFF))) + "]"
    assert word_boundary.CJK_KANJI_RE.pattern == want, [hex(ord(c)) for c in word_boundary.CJK_KANJI_RE.pattern]


def test_both_callers_give_the_same_answer():
    for phrase, task, want in CASES:
        assert word_boundary.stands_alone(phrase, task) is want, (phrase, task)
        assert argcheck._raw_target_not_embedded_in_task(phrase, task) is want, (phrase, task)
        assert alias_store.phrase_is_standalone_in_task(phrase, task) is want, (phrase, task)


def test_both_modules_use_the_one_range():
    assert argcheck._CJK_KANJI_RE is word_boundary.CJK_KANJI_RE
    assert alias_store._CJK_KANJI_RE is word_boundary.CJK_KANJI_RE
