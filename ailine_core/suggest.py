"""suggest — W10 便C1: もしかして提案の候補生成（文字マッチ+about・bge-m3 なし）。

★ なぜ文字マッチだけか（REVIEW-20260822-w10-architect.md 2-4 + Namakoo 決裁）:
architect 実測で bge-m3 は recall@3 89.6% だが呼び出し固定費 2.4s・依存追加。文字 bigram
Jaccard は ~0ms・依存ゼロで recall@1 66.2%/recall@3 68.8%（参考値）。bge-m3 は「文字マッチ+
about で不足が実測されたら」の発火条件つき後送 ── 先に安い方から確かめる。

★ 設計: 語を一切知らない照合器。op の名前も意味も持たず、呼び出し側が渡す
`pool`（op名 -> 照合フレーズの並び）と入力文字列の類似度だけで並べる。
- 文字 bigram の Jaccard（文全体で薄まる ── 短い言い回しが長い依頼文に埋め込まれている
  ケースを拾い損ねる）
- 部分一致（包含）── 短い側の bigram が長い側にどれだけ含まれるかの比率。
  「件名の行をボールドに」に「ボールド」が丸ごと入っている場合、文全体との Jaccard は
  薄まって低いままだが、包含はフレーズ側だけを分母にするので正しく高く出る。
2 つの大きい方をそのフレーズのスコアとする（素朴な OR）。

★ 自己汚染しない設計: この module は凍結セット（bench/w10_suggest_frozen_set.json）の
文言を一切知らない・importしない。閾値/pool は一般則（bigram の重なり）とオペ側が渡す
語彙だけで決まる。凍結セットは bench/run_w10_suggest_eval.py が「測定」に使うだけ。

★ 置き場所: ailine_core/（sum_identity.py と同じ理由）。標準ライブラリだけで閉じ、
ailine.py を import しない（移植可能性の番人 test_line_budget.py が機械で守る）。
"""
from __future__ import annotations

MAX_CANDIDATES = 3

# ノイズ床(無関係な文字列)に候補を出さないための下限。★ 一度決めたら凍結セットを見て
# 動かさない ── 動かした瞬間、この番人は物差しでなく物差しの真似になる
# （feedback_null_result_suspect_instrument と同じ規律）。
MATCH_THRESHOLD = 0.34


def _bigrams(s: str) -> set:
    """文字 2-gram の集合。2 文字未満はその文字列そのものを1要素として扱う
       （空文字列は空集合＝どのフレーズとも一致しない）。"""
    s = s.strip()
    if len(s) < 2:
        return {s} if s else set()
    return {s[i:i + 2] for i in range(len(s) - 1)}


def _phrase_score(text: str, phrase: str) -> float:
    """1 フレーズに対するスコア = 文字 bigram の Jaccard と部分一致(包含)の大きい方。
       包含 = 共通 bigram 数 / 短い側の bigram 数（短い側が長い側にどれだけ埋め込まれて
       いるかの比率・文全体の長さでは薄まらない）。"""
    bt, bp = _bigrams(text), _bigrams(phrase)
    if not bt or not bp:
        return 0.0
    inter = len(bt & bp)
    if not inter:
        return 0.0
    jaccard = inter / len(bt | bp)
    containment = inter / min(len(bt), len(bp))
    return max(jaccard, containment)


def _op_score(text: str, phrases) -> float:
    """op 1 つ分のスコア = そのプールの全フレーズのうち最大のもの。"""
    return max((_phrase_score(text, p) for p in phrases if p), default=0.0)


def suggest_ops(task: str, pool: dict, about: str | None = None, exclude_ops=None,
                 threshold: float = MATCH_THRESHOLD,
                 max_candidates: int = MAX_CANDIDATES) -> list:
    """task（渡されれば about も）を pool と文字マッチし、スコア降順で最大
       max_candidates 件の op 名を返す（閾値未満は候補に出さない）。

       pool:        {op名: [照合フレーズ, ...]}（呼び出し側が組む ── label/synonyms/
                    match_phrases の由来はここでは関知しない）。
       about:       渡された場合、task と about の両方でスコアを取り大きい方を採用する
                    （7B の一次翻訳が返す要約を候補生成にも使い回す・+0ms/+0依存）。
       exclude_ops: このプールから明示的に除く op 集合（測定器の感度確認用）。
       戻り値は実在の op 名のみ（pool のキー以外は絶対に出ない＝幻覚 op の構造的封鎖）。"""
    exclude_ops = exclude_ops or set()
    scored = []
    for op, phrases in pool.items():
        if op in exclude_ops or not phrases:
            continue
        score = _op_score(task, phrases)
        if about:
            score = max(score, _op_score(about, phrases))
        if score >= threshold:
            scored.append((score, op))
    scored.sort(key=lambda t: (-t[0], t[1]))   # op名は同点時の安定な順序付けのみに使う
    return [op for _, op in scored[:max_candidates]]
