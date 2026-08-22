"""alias_store — W10 便A: 別名ストア（言い回し → op 名）の純関数部分。

★ 移植性番人: ここは ailine を import しない（OP_META の実在チェック・ALIASES_FILE の
既定パス・件数上限の**値**そのものは ailine.py 側の責務のまま・ここには渡さない値を
持ち込まない）。置くのは3つだけ:
  ① 検疫（sanitize_phrase）— ailine.py の `_sanitize_vocab_term`（vocab.json 用）と
     同じ規則の写経。vocab とは別ファイルなので依存させず複製する。
  ② 照合（phrase_is_standalone_in_task）— ailine.py の `_raw_target_not_embedded_in_task`
     （単位B・列名照合の断片ガード）と同じ判定の写経。断片問題を3度目に踏まないための
     同型ガードを、別名照合でも独立に持つ（設計ノート③: 「金額」⊂「税込金額」に誤ヒットしない）。
  ③ 保存形式の読み書き（parse_aliases_json / build_aliases_payload）— 平文 JSON
     {"aliases": {言い回し: op名}, "order": [登録順]} を機械的に整える（壊れたファイルは
     クラッシュせず空として扱う＝load_vocab と同じ流儀）。op 名が実在するかどうかの判定
     （OP_META 参照）は呼び出し側が渡す述語関数（is_valid_op）に委ねる。
"""
from __future__ import annotations

import re

_CONTROL_CHAR_RE = re.compile(r"[\x00-\x1f\x7f]")
_CJK_KANJI_RE = re.compile(u"[㐀-䶿一-鿿豈-﫿]")


def sanitize_phrase(phrase, max_len: int) -> str | None:
    """言い回し（キー）が登録可能な形か。空・制御文字・長すぎるものは None（拒否）。"""
    s = str(phrase).strip()
    if not s or len(s) > max_len:
        return None
    if _CONTROL_CHAR_RE.search(s):
        return None
    return s


def phrase_is_standalone_in_task(phrase: str, task: str) -> bool:
    """phrase の task 中の出現のうち、少なくとも1つが「より長い連続した漢字の内部」
       ではない（＝独立した語としての出現がある）なら True。ひらがな/カタカナ/記号は
       語境界として扱う ── 漢字が両隣にも続く場合だけ『内部』とみなす。
       出現が無ければ False（そもそも証拠が無い）。"""
    if not phrase or not task:
        return False
    at = task.find(phrase)
    if at < 0:
        return False
    n = len(phrase)
    while at >= 0:
        before_ok = at == 0 or not _CJK_KANJI_RE.match(task[at - 1])
        after_ok = (at + n) >= len(task) or not _CJK_KANJI_RE.match(task[at + n])
        if before_ok and after_ok:
            return True
        at = task.find(phrase, at + 1)
    return False


def parse_aliases_json(raw, is_valid_op, max_entries: int, max_phrase_len: int) -> tuple:
    """壊れた/形が違う JSON でもクラッシュせず (aliases dict, order list) を返す
       （load_vocab と同じ流儀）。is_valid_op(op) -> bool は呼び出し側（ailine.py）が
       OP_META で判定する述語を渡す（ここでは op の実在は判定しない＝逆流回避）。
       件数が上限を超えた分は読み捨てる（先着順・ファイルの並び順に依存）。
       order は「aliases に実在する言い回しだけ・重複無し・元の並び優先」に正規化し、
       aliases 側にあって order に無かった分は末尾に足す（ファイルが手で壊された場合の保険）。"""
    if not isinstance(raw, dict):
        return {}, []
    raw_aliases = raw.get("aliases")
    if not isinstance(raw_aliases, dict):
        return {}, []
    aliases: dict = {}
    for phrase, op in raw_aliases.items():
        if len(aliases) >= max_entries:
            break
        clean = sanitize_phrase(phrase, max_phrase_len)
        if clean is None:
            continue
        if not isinstance(op, str) or not is_valid_op(op):
            continue
        aliases[clean] = op
    raw_order = raw.get("order")
    order: list = []
    if isinstance(raw_order, list):
        for phrase in raw_order:
            if isinstance(phrase, str) and phrase in aliases and phrase not in order:
                order.append(phrase)
    for phrase in aliases:
        if phrase not in order:
            order.append(phrase)
    return aliases, order


def build_aliases_payload(aliases: dict, order: list) -> dict:
    """保存する JSON の形（{"aliases": ..., "order": ...}）を一箇所で決める。"""
    return {"aliases": aliases, "order": order}
