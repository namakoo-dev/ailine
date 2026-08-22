"""suggest — W10 便C2: もしかして提案の候補生成（語としての厳格一致・bigram 廃止）。

★ なぜ変えたか（封印抜き打ち検体で 5/12 誤提示・Namakoo 決裁 2026-08-22）:
便C1 は文字 bigram の緩い類似度（Jaccard/包含）で候補を出していた。凍結セットの
true_out_of_vocab には効いたが、封印されていた別の12件（ページ番号/テーブル書式設定/
スパークライン/斜線 など）では「そこそこ似てる」が誤って複数の op を拾った。誤提示を
veto 語彙の増築で塞ごうとしても、非対応機能は列挙できない開集合であり収束しない
（黒側の後追いは過適合を繰り返すだけ）。そこで白側の証拠要求を厳格化する方針に転換した:
候補を出す条件を「pool のフレーズが task（または about）の中に**語として**現れる」
だけに絞る。「語として」の判定は ailine_core/alias_store.py の
phrase_is_standalone_in_task と同一の断片ガード（「金額」が「税込金額」の内部に
埋もれているだけなら独立した語と認めない）をそのまま再利用する ── 同じ判定を2つ
持つと将来どちらかだけ直されて食い違う事故になるため、別実装は書かずに import する。

★ 一致は真偽値（現れるか現れないか）。フレーズ長は「複数 op が同時に当たったときの
順位付け」にだけ使う代理指標であり、確信度の連続量ではない（bigram 時代の
0.34/0.8 のような閾値チューニングはもう存在しない）。

★ 自己汚染しない設計: この module は凍結セット（bench/w10_suggest_frozen_set.json）の
文言を一切知らない・importしない。一致規則は一般則（語としての厳格一致）とオペ側が
渡す語彙（pool/veto_phrases）だけで決まる。凍結セットは bench/run_w10_suggest_eval.py
が「測定」に使うだけ。

★ 置き場所: ailine_core/（sum_identity.py と同じ理由）。ailine.py を import しない
（移植可能性の番人 test_line_budget.py が機械で守る）。alias_store は同じ ailine_core/
内の兄弟モジュールなので import してよい（逆流ではない）。
"""
from __future__ import annotations

from ailine_core.alias_store import phrase_is_standalone_in_task

MAX_CANDIDATES = 3


def _op_hit_length(text: str, phrases) -> int:
    """op 1つ分の一致の強さ = task 中に「語として」現れたフレーズ（phrase_is_standalone_
       in_task が True を返すもの）のうち最長の文字数。無ければ 0（一致なし）。
       長さは「同時に複数 op が当たったときの並び順」にだけ使う ── 一致そのものは
       真偽値であり、長さで足切りはしない（1 フレーズでも語として現れれば白）。"""
    if not text:
        return 0
    lengths = [len(p) for p in phrases if p and phrase_is_standalone_in_task(p, text)]
    return max(lengths, default=0)


def suggest_ops(task: str, pool: dict, about: str | None = None, exclude_ops=None,
                 max_candidates: int = MAX_CANDIDATES,
                 veto_phrases=None) -> list:
    """task（渡されれば about も）の中に pool のフレーズが語として現れる op だけを、
       一致した最長フレーズの文字数の降順で最大 max_candidates 件返す。

       pool:        {op名: [照合フレーズ, ...]}（呼び出し側が組む ── label/synonyms/
                    match_phrases の由来はここでは関知しない）。
       about:       渡された場合、task と about のどちらかで語として現れれば拾う
                    （7B の一次翻訳が返す要約を候補生成にも使い回す・+0ms/+0依存）。
       exclude_ops: このプールから明示的に除く op 集合（測定器の感度確認用）。
       veto_phrases: pool とは別の「棄権フレーズ」の並び（任意）。task/about のどちらかに
                    このプールのフレーズが語として現れたら、pool 側の一致に関わらず
                    候補ゼロを返す。中身は呼び出し側の自由 ── ここでは「強く一致したら
                    候補を出さない」という一般則だけを持つ。
       戻り値は実在の op 名のみ（pool のキー以外は絶対に出ない＝幻覚 op の構造的封鎖）。"""
    exclude_ops = exclude_ops or set()
    if veto_phrases:
        veto_hit = _op_hit_length(task, veto_phrases)
        if not veto_hit and about:
            veto_hit = _op_hit_length(about, veto_phrases)
        if veto_hit:
            return []
    scored = []
    for op, phrases in pool.items():
        if op in exclude_ops or not phrases:
            continue
        score = _op_hit_length(task, phrases)
        if about:
            score = max(score, _op_hit_length(about, phrases))
        if score > 0:
            scored.append((score, op))
    scored.sort(key=lambda t: (-t[0], t[1]))   # op名は同点時の安定な順序付けのみに使う
    return [op for _, op in scored[:max_candidates]]
