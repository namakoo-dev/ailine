"""W10 便A（別名ストア）の自分の境界検体 ── 検体④（凍結: 「金額」⊂「税込金額」に
誤ヒットしない）が要求する2例は tests/test_alias_store.py 側で凍結済み。ここでは
brief（「この判定の実装は任せるが... 自分で境界検体を2〜3足すこと」）の指示に沿って、
`ailine_core/alias_store.phrase_is_standalone_in_task` を使う `lookup_alias` の境界を
もう少し広く踏む。

★ この検体は凍結しない（tests/test_alias_store.py と違い、実装の判断が変われば
   ここも一緒に見直してよい）。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_golden_transcripts import _isolate  # noqa: E402


def _use_tmp_aliases(monkeypatch, tmp_path):
    _isolate(monkeypatch, tmp_path)
    monkeypatch.setattr(ailine, "ALIASES_FILE", tmp_path / "aliases.json")


def test_alias_lookup_rejects_when_both_sides_are_kanji(tmp_path, monkeypatch):
    """断片ガードの対称性 ── 「金額」⊂「税込金額」（前だけ漢字）は凍結済み検体で確認済み。
       ここでは前後**両方**が漢字の場合（「字」⊂「太字体」）も同様に当てないことを確認する。
       前だけ確認して後ろを確認していないと、実装がうっかり片方向ガードに戻っても
       この検体だけは緑のままになりうる ── 単位B'（片方向の穴）の再演を防ぐ。"""
    _use_tmp_aliases(monkeypatch, tmp_path)
    ailine.save_alias("字", "BOLD")
    assert ailine.lookup_alias("太字体にして") is None, \
        "前後とも漢字に挟まれた断片（字 ⊂ 太字体）に誤ヒット"


def test_alias_lookup_rejects_prefix_of_longer_kanji_compound_even_if_alias_ends_in_kana():
    """断片問題は「別名が全部漢字」の場合だけでなく、別名の末尾が仮名でも、直後に漢字が
       続いて別の複合語を作っている場合に同様に起きる（「並べ」が「並べ替え」の一部に
       なっている）。alias_store.phrase_is_standalone_in_task はこのケースも拾う。"""
    assert ailine.alias_store.phrase_is_standalone_in_task("並べ", "降順に並べ替えて") is False, \
        "「並べ」が「並べ替え」という別の複合語の内部でしか出現しないのに当たった"
    # 同じ言い回しでも、複合語の内部でなく独立した語として出現する依頼文では当ててよい。
    assert ailine.alias_store.phrase_is_standalone_in_task("並べ", "逆に並べてほしい") is True, \
        "語としての独立した出現まで拒否した（過剰なガード）"


def test_alias_lookup_prefers_the_longest_standalone_match(tmp_path, monkeypatch):
    """⑤ の裏側: 複数の言い回しが同じ依頼文に同時に「語として」ヒットしうる場合
       （短い方が長い方の接頭辞になっている）、より具体的な（長い）方を勝たせる。
       これは brief の凍結5項目には無い、このタスクで足した独自の設計判断 ──
       もし別実装がヒットの中から任意の1件を返す形（例: 辞書の反復順）にしていたら、
       この検体はヒット対象の違いで検出できるはずのケースとして選んだ。"""
    _use_tmp_aliases(monkeypatch, tmp_path)
    ailine.save_alias("並べる", "SORT")
    ailine.save_alias("順に並べる", "AGGREGATE")
    hit = ailine.lookup_alias("順に並べるようにして")
    assert hit == "AGGREGATE", (
        f"短い言い回し『並べる』の op が勝った（{hit}）── "
        "より具体的な（長い）方『順に並べる』を優先すべき"
    )
