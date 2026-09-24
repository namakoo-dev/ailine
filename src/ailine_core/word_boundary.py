"""語の境界 ── 依頼文の中で、ある語が「より長い漢字の連なりの内部」でなく、独立した語として出ているか。

★★ 2026-09-24（同じ名前は同じ物の番人の棚卸しで見えた）: 同じ判定が 2 つあった ──
  argcheck._raw_target_not_embedded_in_task（元は本体）と alias_store.phrase_is_standalone_in_task。
  別名の照合を core に置いた時、core は本体を import できないので**写経した**（本体の注記にそう書いてあった）。
  その写しの漢字の範囲だけが、互換漢字の「豈」（U+F900）の代わりに**見た目が同じ別の字**・統合漢字の
  「豈」（U+8C48）で始まっていて、U+8C48〜U+FAFF（ハングル・彝文字・私用領域まで）を「漢字」と数えていた。
  ★ 判定が core に揃った今は写す理由が無いので、ここ 1 本にする。
★ 範囲は**コードポイントで**書く（字形で書くと、見た目が同じ別の字が紛れる）。
★ ailine を import しない（持ち出せる部品）。
"""
from __future__ import annotations

import re

#: 漢字: CJK 統合漢字拡張 A・CJK 統合漢字・CJK 互換漢字
CJK_KANJI_RE = re.compile("[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")


def stands_alone(phrase: str, task: str) -> bool:
    """phrase の task 中の出現のうち、少なくとも 1 つが「より長い連続した漢字の内部」ではない
    （＝独立した語としての出現がある）なら True。ひらがな/カタカナ/記号は語境界として扱う ──
    漢字が両隣にも続く場合だけ『内部』とみなす。出現が無ければ False（そもそも証拠が無い）。"""
    if not phrase or not task:
        return False
    at = task.find(phrase)
    if at < 0:
        return False
    n = len(phrase)
    while at >= 0:
        before_ok = at == 0 or not CJK_KANJI_RE.match(task[at - 1])
        after_ok = (at + n) >= len(task) or not CJK_KANJI_RE.match(task[at + n])
        if before_ok and after_ok:
            return True
        at = task.find(phrase, at + 1)
    return False
