"""operator9 ①②の実装に添える自分の境界検体（凍結済み tests/test_operator9_fixes.py は
   期待値を1文字も変更しない・こちらは追加の番人）。

   ① extract_cmp_from_task: 「以上」「以下」は文末定型（「以上です」等）の断片として
     現れやすいので、値の近傍にだけ絞る断片ガードが効いているかを見る。
   ② _op_has_task_grounding: 依頼文に語として在るかどうかの断片ガード（1文字の値は
     偶然一致しすぎるので証拠にしない・より長い漢字の内部に埋もれた出現は証拠にしない）
     が効いているかを見る。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402


# --- ① 比較語の機械抽出: 断片ガード ---------------------------------------------

def test_extract_cmp_ignores_trailing_ijou_desu_without_nearby_number():
    """「以上です」は文末の定型句であって比較語ではない ── 直前に数字が無ければ採用しない。"""
    assert ailine.extract_cmp_from_task("報告は以上です。ご確認ください。") is None


def test_extract_cmp_gte_still_works_with_fullwidth_digit_nearby():
    """全角数字（LLM/日本語入力でありうる表記）でも「以上」の直前なら比較語として採用する。"""
    assert ailine.extract_cmp_from_task("金額が５０００以上の行を抜き出して") == "gte"


def test_extract_cmp_rejects_ika_when_number_is_in_a_different_sentence():
    """数字と「以下」が別の文（句点を跨ぐ）にある場合は近傍とみなさない。"""
    assert ailine.extract_cmp_from_task("先月は5000円でした。以下のとおり修正します") is None


# --- ② 捏造段の検出: 断片ガード -------------------------------------------------

def test_op_has_task_grounding_true_when_task_is_empty():
    """判定材料（依頼文）が無ければ誤って★を出さない ── 根拠ありに倒す。"""
    assert ailine._op_has_task_grounding("SORT", {}, "") is True


def test_op_has_task_grounding_rejects_single_char_arg_value_as_evidence():
    """1文字の値は偶然一致しすぎるので、依頼文にその文字が現れても証拠にしない
       （単位B の _MIN_FRAGMENT=2 と同じ理由）。"""
    assert ailine._op_has_task_grounding("SORT", {"target": "計"}, "小計を確認して") is False


def test_op_has_task_grounding_rejects_value_embedded_in_longer_kanji_word():
    """2文字以上でも、依頼文中の全出現が「より長い漢字語の内部」でしかないなら証拠にしない
       （例:『集計』への出現しか無い『計算』は、集計の一部でしかないので証拠にならない）。"""
    assert ailine._op_has_task_grounding("SORT", {"target": "計算"}, "集計算出を確認して") is False


def test_op_has_task_grounding_true_via_pool_word_even_without_args_match():
    """照合プール句（label/synonyms/match_phrases）が依頼文に語として在ればそれだけで根拠あり。"""
    assert ailine._op_has_task_grounding("SORT", {"col": "存在しない列名です"}, "並べ替えして") is True
