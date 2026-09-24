"""sum_identity の _is_number は primitives.is_number の**意図した写し** ── 同じ答えを返し続ける（2026-09-24）。

★ なぜ写しなのか: sum_identity は「標準ライブラリだけで閉じる」契約を持つ（言語非依存）ので、primitives を
  import できない（2026-09-03 に畳もうとして契約を破って気づいた ── sum_identity の docstring）。
★ 写しは黙ってずれる（片配線の静かな側）。畳めないので、同じ入力で同じ答えを返すことを縛る。
"""
import sys
from decimal import Decimal
from fractions import Fraction
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from ailine_core import primitives, sum_identity  # noqa: E402

VALUES = [0, 1, -3, 2.5, float("nan"), float("inf"), True, False, None, "1", "", "１",
          Decimal("1.5"), Fraction(1, 2), complex(1, 0), [1], {"a": 1}, b"1"]


def test_the_copy_answers_the_same_as_the_original():
    diff = [(repr(v), primitives.is_number(v), sum_identity._is_number(v)) for v in VALUES
            if primitives.is_number(v) != sum_identity._is_number(v)]
    assert not diff, f"写しが元とずれた（値・元・写し）: {diff}"


def test_the_sample_has_both_answers():
    """★ 陽性対照: 検体に True と False の両方の答えが含まれる（片方だけなら比べても恒真）。"""
    answers = {primitives.is_number(v) for v in VALUES}
    assert answers == {True, False}
