"""翻訳 prompt 3 定数の凍結番人（architect M2 レビューの指摘・2026-08-21）。

★ なぜ在るか: battery の凍結バー（op90%/slot80%）は「prompt が動いていない」前提での
数字だ。W9 の実測では few-shot を 1 例足しただけで別 op の誤断定が 27.3% に跳ねた ──
prompt は動くと壊れる種類の部品なのに、動いたことを知らせる番人が無かった
（「在っても鳴らない」の亜種: battery は在るが、走らせない限り黙っている）。

この試験は 3 定数の SHA-256 を凍結する。意図して変えた時は、battery を回して
凍結バーを確認してから、下のハッシュを更新して同じ commit に入れること
（更新だけの commit は「測らずに動かした」ことが git 履歴で見える）。
★ M2（run のフォルダ分岐）の設計判断「LLM は複数ファイルを知る必要がない ── 1 語も
足さない」も、この番人が機械で保証する。"""
import hashlib
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402

_FROZEN = {
    # ★ 2026-08-21: DEDUP op の語彙昇格（freeform 廃止バンドル前段）で OPS_DOC/
    #   TRANSLATION_FEWSHOT を更新。bench/translation_dsl_battery_run.py 実 7B
    #   （qwen2.5-coder:7b）で凍結バーを確認済み: op 96.2%→98.1%（合格線90%）・
    #   slot 98.6%→98.6%（合格線80%）・曖昧誤断定 0%→0%（合格線20%以下、変化なし）。
    #   DEDUP 専用 battery（items_v8・3件）は op/slot とも100%。
    "OPS_DOC": "2a87a1945bc6e29c",
    "TRANSLATION_SYSTEM": "8dd5cd3a43d833fe",
    "TRANSLATION_FEWSHOT": "c10a9b45e6cada35",
    # ★ 2026-08-22: W10 便B（二段目翻訳・op 固定で args だけ埋めさせる）で追加した第4の定数。
    #   OPS_DOC 全文を見せず、固定した op 1 つ分のスキーマだけを見せる（W9 の 27.3% 誤断定の
    #   教訓＝別名/few-shot を混ぜると壊れる部品）。frozen 対象は tests/test_fixed_op_translation.py
    #   が番人（test_fixed_op_prompt_is_frozen_constant）。
    "TRANSLATION_FIXED_OP_SYSTEM": "230c3e0e92d3aa20",
}


def _digest(value) -> str:
    return hashlib.sha256(repr(value).encode("utf-8")).hexdigest()[:16]


def test_translation_prompt_constants_are_frozen():
    for name, frozen in _FROZEN.items():
        actual = _digest(getattr(ailine, name))
        assert actual == frozen, (
            f"{name} が変わった（凍結 {frozen} / 実際 {actual}）。意図した変更なら "
            f"battery（bench/translation_dsl_battery_run.py）で凍結バーを確認してから、"
            f"この試験のハッシュを同じ commit で更新すること。"
        )
