# 台帳の実文言に対する「もしかして提案」の到達率 ── 番人（2026-08-24）。
#
# ★ なぜ在るか: 3 op（REPORT_PER_ROW/FORMAT_MAP/CSV_EXPORT）を出荷した直後に、
#   台帳（MARKET-20260823-lancers.md）の実案件の言い回しで測ったら **1/7 しか
#   候補が出なかった**（偽陽性は 0）。op は在るのに、依頼者の言い方では扉が開かない。
#   「在っても鳴らない」の提案側の亜種 ── 出荷は翻訳経路にだけ配線され、提案の
#   照合語彙が置き去りになっていた（片配線）。
#
# ★ 測定器の出所を分ける: この検体は台帳（実案件）由来で、退行を守る
#   bench/w10_suggest_frozen_set.json（陽性対照/ノイズ床/感度）とは**別出所**。
#   同じ集合で語彙を足して同じ集合で測ると自己汚染になるため、両方を別々に回す。
#
# ★ 等号で縛る（率でなく件数）。増える方向に動かすときも、必ずこの数字を書き換えて
#   diff に出す。★ L01 は語彙ではなく**断片ガード**で落ちている既知の穴（下記）。

import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

HOLDOUT = json.loads((REPO / "bench" / "ledger_phrasing_holdout.json").read_text(encoding="utf-8"))


def test_ledger_phrasings_reach_the_ops_we_have():
    """当たるべき 7 件のうち 6 件で、期待した op が候補に含まれる（等号）。"""
    hit, missed = 0, []
    for e in HOLDOUT["should_hit"]:
        cands = ailine.suggest_ops(e["text"]) or []
        if e["expect_op"] in cands:
            hit += 1
        else:
            missed.append((e["id"], e["text"], cands))
    assert hit == 6, f"到達件数が 6 でない（実測 {hit}）。外れ: {missed}"


def test_the_known_hole_is_the_fragment_guard_not_the_vocabulary():
    """L01「複数列の一括入れ替え」が落ちる理由を名指しで固定する。

    照合語彙には「入れ替え」が在る。落ちるのは alias_store の断片ガードが
    『一括入れ替え』の内部に埋もれた「入れ替え」を独立した語と認めないため。
    ★ ここに「一括入れ替え」を足して 7/7 にするのは、検体 1 個への当てはめ
      （うま味調味料）なので**やらない**。断片ガードそのものを見直す時に、
      別出所の検体を集めてから測り直す。
    """
    from ailine_core.alias_store import phrase_is_standalone_in_task
    assert "入れ替え" in ailine.OP_META["FORMAT_MAP"]["match_phrases"]
    assert phrase_is_standalone_in_task("入れ替え", "列の入れ替え") is True
    assert phrase_is_standalone_in_task("入れ替え", "複数列の一括入れ替え") is False
    assert ailine.suggest_ops("複数列の一括入れ替え") == []


def test_no_false_suggestions_for_ops_we_do_not_have():
    """持っていない機能（横結合/セル分割/印刷/日付計算/ランダム割付/ツール納品）に
       候補を出さない ── 0/6 の等号。★ ここが 1 でも増えたら「できると言う嘘」。"""
    fired = [(e["id"], e["text"], ailine.suggest_ops(e["text"]))
              for e in HOLDOUT["should_not_hit"] if ailine.suggest_ops(e["text"])]
    assert fired == [], f"無い機能に候補を出した: {fired}"
