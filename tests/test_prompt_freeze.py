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
    "OPS_DOC": "1ad4870e37799d84",
    "TRANSLATION_SYSTEM": "8dd5cd3a43d833fe",
    "TRANSLATION_FEWSHOT": "d29158d01f236c2f",
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
