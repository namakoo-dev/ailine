"""重複（DEDUP）の鍵の規則 ── 書く前の検査（argcheck）と事後条件（postconditions.move）が同じ規則を引く。

★ 2026-09-24 に postconditions/move.py から**移しただけ**（本文は移した後に docstring の 1 か所だけ直した）。
  引数の検査が事後条件の内側を import していた配置を解くため、中立な冊にした。move は同じ実体を再輸出する。
"""
from __future__ import annotations


def _dedup_normalize_key_part(v):
    """DEDUP のキー正規化: 前後空白除去のみ・型が違えば別キー。

       ★ 2026-09-24 に直した docstring: 以前は「match.normalize_key と同じ規則」と書いていたが、
         **None と空文字の扱いが違う**（match は鍵不明の None を返し、こちらは普通の値として比べる ──
         空の行どうしは重複とみなす）。同じ規則ではないので畳まない。"""
    if isinstance(v, str):
        return ("str", v.strip())
    return (type(v).__name__, v)
