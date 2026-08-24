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
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

_FROZEN = {
    # ★ 2026-08-21: DEDUP op の語彙昇格（freeform 廃止バンドル前段）で OPS_DOC/
    #   TRANSLATION_FEWSHOT を更新。bench/translation_dsl_battery_run.py 実 7B
    #   （qwen2.5-coder:7b）で凍結バーを確認済み: op 96.2%→98.1%（合格線90%）・
    #   slot 98.6%→98.6%（合格線80%）・曖昧誤断定 0%→0%（合格線20%以下、変化なし）。
    #   DEDUP 専用 battery（items_v8・3件）は op/slot とも100%。
    # ★ 2026-08-23: グラフ段（kind/category_col の語彙拡張）で OPS_DOC の CHART 行を更新。
    #   bench/translation_dsl_battery_run.py 実 7B（qwen2.5-coder:7b）で凍結バーを確認済み:
    #   op 98.1%→98.1%・slot 98.6%→98.6%・曖昧誤断定 0%→0%（合格線90%/80%/20%以下・変化なし）。
    # ★ 2026-08-24: 帳票段（REPORT_PER_ROW）の語彙昇格で OPS_DOC に op 説明を追加。
    #   bench/translation_dsl_battery_run.py 実 7B（qwen2.5-coder:7b・items v1・52件）で
    #   既存語彙への退行が無いことを確認済み: op 98.1%(51/52)・slot 98.6%(71/72)・
    #   曖昧誤断定 0%(0/11)（合格線90%/80%/20%以下・すべてクリア）。
    #   帳票段専用 battery（bench/translation_battery_run.py v9・items_v9・4件・暫定バー
    #   op75%/slot80%）は op 100%(4/4)・slot 100%(8/8)。
    # ★ 2026-08-24: 様式写像段（FORMAT_MAP）の語彙昇格で OPS_DOC に op 説明を追加。
    #   bench/translation_dsl_battery_run.py 実 7B（qwen2.5-coder:7b・items v1・52件）で
    #   既存語彙への退行が合格線内であることを確認済み: op 94.2%(49/52)・slot 98.5%(67/68)・
    #   曖昧誤断定 0%(0/11)（合格線90%/80%/20%以下・すべてクリア。LLM サンプリングの揺れで
    #   帳票段追加直後の98.1%からは下がったが合格線は割っていない）。
    #   様式写像段専用 battery（bench/translation_battery_run.py v10・items_v10・4件・
    #   暫定バー op75%/slot80%）は op 75%(3/4)・slot 100%(3/3)。
    # ★ 2026-08-24 の検分で退行を捕獲: 新 op（REPORT_PER_ROW/FORMAT_MAP）の記述を
    #   16 行書いたら v1 が 98.1%→94.2% に落ちた（FILL_COLOR/NUMBER_FORMAT が新たに崩れ、
    #   2 回再現・git stash で帰属も確認）。★ W9 の「few-shot 1 例で誤断定 27.3%」の再演。
    #   機械が必ず強制する規則（実在検証・印以外に触らない・出力名の決定）を OPS_DOC から
    #   外して 4 行に削り、98.1%/98.6% に復帰。代償は v9 が 4/4→3/4（合格線 75% ちょうど）。
    #   ★ これ以上は同じ検体で調整しない（自己汚染）── 実運用の誤分類が出たら測り直す。
    # ★ 2026-08-24: SPLIT_CELL の語彙昇格で OPS_DOC に op 説明を **1 行**追加。
    #   bench/translation_dsl_battery_run.py 実 7B（qwen2.5-coder:7b・items v1・52件）で
    #   実走して確認: op 96.2%(50/52)・slot 98.6%(69/70)・曖昧誤断定 0%(0/11)
    #   （合格線 90%/80%/20%以下 ── すべてクリア）。
    #   ★ 正直に記録する: 帳票段直後の 98.1% からは **1 件ぶん下がった**（不一致は
    #   #12 LOOKUP_FILL→FREEFORM と #39 NUMBER_FORMAT→CLARIFY）。1 行でも他 op を
    #   押しのけるという 08-24 の観測（16 行で 98.1%→94.2%）と同じ向きで、桁が小さいだけ。
    #   ★ この実測を根拠に、台帳の残り（DATE_CALC 2 件・PRINT/EXPORT_DOC 4 件）は
    #   **OPS_DOC を増やさない形**で実装した ── PRINT/EXPORT_DOC は `export-pdf`
    #   サブコマンド、日付の扱いは既存 op（EXTRACT/SORT）の穴埋めとして。
    "OPS_DOC": "f1d00fb654aed28a",
    "TRANSLATION_SYSTEM": "8dd5cd3a43d833fe",
    "TRANSLATION_FEWSHOT": "c10a9b45e6cada35",
    # ★ 2026-08-22: W10 便B（二段目翻訳・op 固定で args だけ埋めさせる）で追加した第4の定数。
    #   OPS_DOC 全文を見せず、固定した op 1 つ分のスキーマだけを見せる（W9 の 27.3% 誤断定の
    #   教訓＝別名/few-shot を混ぜると壊れる部品）。frozen 対象は tests/test_fixed_op_translation.py
    #   が番人（test_fixed_op_prompt_is_frozen_constant）。
    "TRANSLATION_FIXED_OP_SYSTEM": "230c3e0e92d3aa20",
    # ★ 2026-08-22: W10 便C2（もしかして提案の判定器・第2段）で追加した第5の定数。
    #   両面プロンプト（candidates+unsupported）── unsupported が非空なら候補を丸ごと
    #   捨てる。frozen 対象は tests/test_suggest_flow.py が番人
    #   （test_judge_prompt_is_fifth_frozen_constant）。
    "SUGGEST_JUDGE_SYSTEM": "27286fa0a6be6117",
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
