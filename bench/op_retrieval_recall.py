#!/usr/bin/env python
"""op を「絞って見せる」案の go/no-go ── 正解が上位 k に入るかだけを測る。

★★ 2026-08-30（Namakoo「語彙の拡張については RAG を使うアイデアもあるな」）:
  今夜の実測は「**op を増やすと完遂率が落ちる**」だった（98.9% → 82.8%・壊した 0→9）。
  選択肢が増えてモデルが迷い、機械が持っていた確実な経路を奪ったため。
  ★ ここでの RAG は**知識を足す**話ではなく、**依頼ごとに選択肢を絞る**話 ──
    24 個ぜんぶ見せる代わりに、近い 4〜5 個だけ見せる。方向が逆。
  ★ 1B で段階的に聞いて効いたのと同じ梃子（選択肢を減らす）を、手で木を作る代わりに
    連続量でやる。未知の言い回しにも寄るので「列挙は漏れる」への答えにもなる。

★★ ただし新しい壊れ方が生まれる: **正解が上位 k に入らなければ、モデルは選べない。**
  絞ることは同時に取りこぼすこと ── しかも静かに。
  だから採否はこの 1 つの数字で決まる:
      recall@k = 正解の op が上位 k 個に入っていた依頼の割合

★ LLM は使わない（埋め込みだけ）。外部送信も無い（ollama のローカル埋め込み）。
"""

from __future__ import annotations

import json
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "bench"))

import ailine  # noqa: E402
from vocab_reach import CASES  # noqa: E402  ★ 同じ 48 件を使う（検体を増やさない）

EMBED = __import__("os").environ.get("AILINE_EMBED", "bge-m3:latest")
KS = (3, 5, 8)


def _embed(text: str) -> list:
    body = json.dumps({"model": EMBED, "prompt": text}).encode()
    req = urllib.request.Request("http://localhost:11434/api/embeddings", body,
                                  {"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=180) as r:
        return json.loads(r.read())["embedding"]


def _cos(a: list, b: list) -> float:
    s = sum(x * y for x, y in zip(a, b))
    na = sum(x * x for x in a) ** 0.5
    nb = sum(y * y for y in b) ** 0.5
    return s / (na * nb) if na and nb else 0.0


def _op_text(op: str) -> str:
    """op を表す文（★ 登録簿から作る ── 新しい辞書を書かない）。"""
    m = ailine.OP_META.get(op) or {}
    parts = [str(m.get("label") or op)]
    parts += [str(x) for x in (m.get("synonyms") or [])]
    parts += [str(x) for x in (m.get("match_phrases") or [])]
    return "。".join(parts)


def main() -> int:
    ops = sorted(ailine.OP_META)
    print(f"op {len(ops)} 個を埋め込み中（{EMBED}）…")
    vecs = {o: _embed(_op_text(o)) for o in ops}

    total = 0
    hits = {k: 0 for k in KS}
    misses = []
    for want, tasks in CASES:
        for t in tasks:
            total += 1
            v = _embed(t)
            ranked = sorted(ops, key=lambda o: -_cos(v, vecs[o]))
            pos = ranked.index(want) + 1 if want in ranked else 10 ** 6
            for k in KS:
                hits[k] += (pos <= k)
            if pos > max(KS):
                misses.append((want, t, ranked[:3], pos))
    print()
    for k in KS:
        print(f"  recall@{k}  {hits[k]}/{total} = {hits[k] / total * 100:.1f}%")
    print(f"  （op {len(ops)} 個から絞る・依頼 {total} 件・埋め込みのみ）")
    if misses:
        print()
        print(f"★ 上位 {max(KS)} 位にも入らなかったもの（ここが静かな取りこぼしになる）:")
        for w, t, top, pos in misses:
            print(f"    {w:18} {t}  → {pos}位（上位: {'・'.join(top)}）")
    print()
    print("★ 判断: recall@5 がほぼ 100% なら次へ（実際に絞って完遂率を測る）。")
    print("  90% 台前半なら**採用しない** ── 1 割の依頼が静かに到達不能になる。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
