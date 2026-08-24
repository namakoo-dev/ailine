# 対象シートの決定 ── 明示マーカー付きの言及は、裸の言及に殺されない（2026-08-24）。
#
# ★ 実測した事故: 複合計画の連鎖の ⚠ が「シート名を依頼文に書いて実行し直してください」と
#   案内するので、その逃げ道を実物で確かめたら**効かなかった**。
#     ブック: ['売上','売上60以上','集計']（前の run が 集計 を作っている）
#     依頼:  「売上60以上シートを現場ごとに集計して」
#     結果:  対象シート = 『売上』(既定) ── 動詞「集計して」が**シート名『集計』と一致**して
#            言及が2つになり、曖昧と判断されて既定へ後退していた。
#   ★ つまり ⚠ が案内する逃げ道が嘘になっていた。番人の文言も契約のうち。
#
# 契約: 言及が複数あっても、**「シート」「タブ」を後ろに伴う言及がちょうど1つ**なら
#       それを採る（裸の言及は、動詞や一般語との偶然の一致でありうるため譲る）。

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
from ailine_core.target_sheet import resolve_target_sheet  # noqa: E402

SHEETS = ["売上", "売上60以上", "集計"]
HEADERS = {"売上": ["現場", "売上"], "売上60以上": ["現場", "売上"],
            "集計": ["現場", "合計 - 売上"]}


def test_marked_mention_wins_over_a_bare_verb_collision():
    name, source, err, conflict = resolve_target_sheet(
        "売上60以上シートを現場ごとに集計して", SHEETS, None, HEADERS)
    assert name == "売上60以上", f"明示指定が裸の言及に殺された（実測の再現）: {name} / {source}"
    assert source == "task"


def test_two_marked_mentions_are_still_ambiguous():
    """誤爆防止: マーカー付きが2つあるなら、今までどおり曖昧＝既定へ後退する
       （LOOKUP_FILL の『転記先シート』『参照元シート』を勝手に片方へ寄せない）。"""
    # ★ 治具の訂正（封印者ナギ・2026-08-24）: 初版は「売上シートから売上60以上シートへ」で、
    #   これは『売上』⊂『売上60以上』の部分文字列規則で 1 つに畳まれるため、そもそも
    #   「マーカー2つ」の検体になっていなかった（実測で確認）。重ならない2枚に差し替えた。
    #   assert の意図（マーカーが2つなら曖昧のまま既定へ）は不変。
    name, source, _e, _c = resolve_target_sheet(
        "集計シートから売上60以上シートへ転記して", SHEETS, None, HEADERS)
    assert (name, source) == ("売上", "default"), f"曖昧なのに片方を選んだ: {name} / {source}"


def test_single_bare_mention_behaviour_is_unchanged():
    """退行防止: 裸の言及が1つだけの従来ケースは今までどおり（列名と衝突しなければ採用）。"""
    name, source, _e, _c = resolve_target_sheet(
        "売上60以上を現場ごとにまとめて", SHEETS, None, HEADERS)
    assert (name, source) == ("売上60以上", "task")


# --- 同じ家系の片配線: 言及の抽出も部分文字列を畳む（2026-08-24）------------------
#
# ★ 実測: 「売上60以上シートを現場ごとに集計して」で対象シートは正しく『売上60以上』に
#   なったのに、言及の抽出は『売上』も数えていて「依頼で言及された『売上』は変更されて
#   いません」という ⚠ が出た。決裁③でその ⚠ が ✓ を △ に降格させる ── **正しくできた
#   仕事に、誤った ⚠ で傷が付く**。対象シートの決定は畳むのに言及の抽出は畳んでいない
#   ＝同じ規則の片配線。

def test_mentions_collapse_substring_sheet_names():
    import ailine
    m = ailine.extract_task_mentions("売上60以上シートを現場ごとに集計して",
                                      ["売上", "売上60以上", "集計"])
    assert "売上" not in m["sheets"],         f"『売上60以上』の一部でしかない『売上』を言及として数えた: {m['sheets']}"
    assert "売上60以上" in m["sheets"]


def test_mentions_keep_a_genuinely_named_shorter_sheet():
    """誤爆防止: 短い方が**独立に**書かれているなら、両方とも言及として数える。"""
    import ailine
    m = ailine.extract_task_mentions("売上シートと売上60以上シートを見比べて",
                                      ["売上", "売上60以上", "集計"])
    assert m["sheets"] >= {"売上", "売上60以上"}, f"独立の言及を落とした: {m['sheets']}"
