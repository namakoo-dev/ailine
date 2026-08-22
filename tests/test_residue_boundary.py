# W10 便C2 S5 の追補: 自分で足す境界検体2本（ailine_core/residue.py の純ロジック）。
# ブリーフの縛り: 「オオカミ少年化（きれいな依頼への誤発火）を最も恐れよ」── 1本は
# 未消費のカタカナ内容語を確実に拾えること、もう1本は複数文字の解決値が丸ごと消費されて
# 断片が残らないこと（きれいな依頼側の追加保証）を確かめる。
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
from ailine_core import residue  # noqa: E402

# SORT の照合語彙（ailine.py の OP_META["SORT"] の label+synonyms+match_phrases と同じもの）。
_SORT_POOL = ["並べ替え", "ソート", "順に並べる", "順番", "昇順", "降順", "整列",
              "並び替える", "順位付け"]


def test_katakana_content_word_is_flagged_as_residue():
    """境界①: 漢字だけでなく、未消費のカタカナ内容語（例: グラフ）も残差として拾う。"""
    words = residue.find_unconsumed_words(
        "金額を並べ替えしてグラフも作って", {"col": "金額", "order": "desc"}, _SORT_POOL)
    assert "グラフ" in words, f"カタカナの残差語を拾えていない: {words}"


def test_multi_kanji_resolved_value_leaves_no_fragment():
    """境界②: 解決済み値が複数漢字の複合語（例: 税込金額）でも丸ごと1つの span として
       消費され、断片（例: 込/金額）が残差として誤って残らない（オオカミ少年の追加保証）。"""
    words = residue.find_unconsumed_words(
        "税込金額を並べ替えして", {"col": "税込金額", "order": "desc"}, _SORT_POOL)
    assert words == [], f"消費済みの列名から断片が残差として漏れた: {words}"
