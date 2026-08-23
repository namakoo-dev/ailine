# W10 便C1: 候補生成 suggest_ops（文字マッチ+about・bge-m3 なし）── 実装より先に凍結。
# 凍結セット bench/w10_suggest_frozen_set.json の bars を機械の番人にする。
# in_vocab の recall は内部診断値（bench スクリプトが件数報告）── ここでは縛らない。
#
# 契約:
#   ① 陽性対照: OP_META の label そのものは recall@1 = 18/18（割れたら機構の故障）
#   ② ノイズ床: 無意味文字列 10 件に候補を 1 つも出さない
#   ③ 感度: プールから SORT を意図的に抜いた劣化版で、SORT の label が当たらなくなる
#      （下がらなければ測定器が壊れている）
#   ④ 候補は OP_META に実在する op 名のみ・最大 3 件（幻覚 op の構造的封鎖）

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "suggest_ops"),
    reason="候補生成 未実装（バーは凍結済み・実装が来たら自動で実測に切替）",
    strict=True,
)

FROZEN = json.loads((REPO / "bench" / "w10_suggest_frozen_set.json").read_text(encoding="utf-8"))


@needs_impl
def test_positive_control_labels_hit_at_rank1():
    """①: label そのもの → その op が第 1 候補。18/18 の等号（率でない）。"""
    hits = 0
    misses = []
    for e in FROZEN["positive_control"]:
        cands = ailine.suggest_ops(e["text"])
        if cands and cands[0] == e["expect_op"]:
            hits += 1
        else:
            misses.append((e["id"], e["text"], cands))
    assert hits == 18, f"陽性対照の取りこぼし（機構の故障）: {misses}"


@needs_impl
def test_noise_floor_yields_no_candidates():
    """②: 無意味文字列に候補を出さない。0/10 の等号。"""
    fired = [(e["id"], ailine.suggest_ops(e["text"]))
              for e in FROZEN["noise_floor"] if ailine.suggest_ops(e["text"])]
    assert fired == [], f"ノイズに候補が出た（閾値の故障）: {fired}"


@needs_impl
def test_sensitivity_removing_op_from_pool_removes_hits():
    """③: SORT をプールから抜くと SORT の label が当たらない（測定器の感度確認）。"""
    assert ailine.suggest_ops("並べ替え") and ailine.suggest_ops("並べ替え")[0] == "SORT"
    degraded = ailine.suggest_ops("並べ替え", exclude_ops={"SORT"})
    assert "SORT" not in (degraded or []), "抜いたはずの op が出る（感度ゼロ=測定器の故障）"


@needs_impl
def test_candidates_are_real_ops_and_capped_at_3():
    """④: どんな入力でも、候補は実在 op のみ・3 件以下。"""
    for e in FROZEN["in_vocab"] + FROZEN["true_out_of_vocab"] + FROZEN["slot_missing"]:
        cands = ailine.suggest_ops(e["text"]) or []
        assert len(cands) <= 3, f"{e['id']}: 候補が 3 件超"
        unknown = [c for c in cands if c not in ailine.OP_META]
        assert not unknown, f"{e['id']}: 実在しない op {unknown}"
