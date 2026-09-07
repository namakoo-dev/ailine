# -*- coding: utf-8 -*-
"""依頼が名指しした**操作の種類**と、実行した操作が食い違っていないか。

★★ なぜ在るか（2026-09-07・外部の検品が最重の所見として拾った）:

    依頼   「ヤマノ食品の行を**削除して**」
    実行   操作:**抽出** → 新しいシートを作り、元の 6 行はそのまま
    出力   **✓ 機械検証済み**

  事後条件は「抽出として正しいか」を確かめるので通る ── **宣言と実体は一致していて、
  依頼だけが落ちている**。判定に要る三項のうち、また依頼が見られていなかった。

★ 既存の関所（`residue.unaccounted_request_words`）は**列名**しか見ないので、
  『削除』のような**動詞**は拾えない。ここはその隣を受け持つ。

★★ 何を見るか ── op 名の一致ではなく、**効果の種類**を見る:

      依頼文が「取り除く」系の op の語彙に当たっている
      かつ 実行した op が「取り除く」を書かない       → ✓ を出さない

  ★ op 名で見ると誤爆する。実測（4,538 件の実走行）:

      COMPUTE_COLUMN ← 「列を追加して」が ADD_COLUMN に当たる  x22  ← 上位下位の重なり
      DEDUP          ← 「行を消して」が DELETE_ROWS に当たる    x14  ← ★ 本物
      EXTRACT        ← 「削除して」                             x 2  ← ★ 本物

    効果の種類で見ると上の x22 は消え、**16 件 0.35% が残り、全部が同じ家系**だった
    （DEDUP も EXTRACT も新しいシートを作るだけで、元の行を取り除かない）。

★ 直さない・止めない ── ⚠ を出して ✓ を降ろすだけ（既存の関所と同じ作法）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations


def removal_asked_but_not_done(task: str, op: str, vocab_by_op: dict,
                               removes: dict) -> list:
    """「取り除く」を頼まれたのに取り除かなかったなら、当たった語を返す（無ければ空）。

    task        … 依頼文
    op          … 実行した操作
    vocab_by_op … {op: その op の照合語彙}（`_op_match_pool` の結果）
    removes     … {op: その op が「取り除く」を書くか}（書き込み様式の登録簿から）

    ★ 実行した op **自身の語彙**が依頼文に当たっているなら、食い違いとは言わない
      （「重複行を削除して」で DEDUP が選ばれた回に、DEDUP の語彙も当たっていれば
      利用者はその操作を名指ししている）。
    """
    text = task or ""
    if not text or removes.get(op):
        return []                       # 取り除く op を実行したなら、食い違わない
    if any(p for p in (vocab_by_op.get(op) or ()) if p and p in text):
        return []                       # 実行した op 自身が名指しされている
    hits = []
    for other, phrases in (vocab_by_op or {}).items():
        if other == op or not removes.get(other):
            continue
        hits += [p for p in (phrases or ()) if p and p in text]
    return list(dict.fromkeys(hits))
