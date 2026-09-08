# -*- coding: utf-8 -*-
"""「依頼文にこの語が明示された時だけ使う」を、散文でなく**機械**にする。

★★ なぜ在るか（2026-09-08・Namakoo「使い方の広い操作に出る不具合が怖い」から辿った）:

    依頼   「分類ごとの売上を出して」          ← 『ピボット』とは一言も言っていない
    実行   操作:**ピボット** 分類列:分類 集計列:売上
    出力   **✓ 機械検証済み**（数字は正しい・シート名は『ピボット』）

  ★ 禁止は**もう書いてあった** ── プロンプトに 3 行:
      「PIVOT: 依頼文に『ピボット』…が明示された時だけ使う」
      「★『ピボット』の語が無いグループ別集計はすべて AGGREGATE を使う」
    さらに few-shot に「部門別の金額合計がみたい → AGGREGATE」まで在る。
    それでも実物の CLI で **4/4 ピボット**だった。★ 指示は意図、保証は機械。

★★ これは「狭い op が広い op を奪う」形（同日の実測 3 件目）。集計は
  「〜ごとに」「まとめて」「別に」が全部流れ込む**広い口**で、そこを狭いピボットが取る。
  被害の重さは頻度に比例するので、広い側が奪われるのがいちばん高くつく。

★ 同じ契約が 3 つある（実測では今のところ壊れているのはピボットだけ・
  他の 2 つは曖昧な依頼をちゃんと聞き返している）。★ **今は鳴らなくても縛る** ──
  「在っても鳴らない」の逆で、鳴る条件が将来できたときに黙って通るのを防ぐ。

★ 引数の形が同じ op へは**読み直す**（ピボット→集計は group_col/value_col が同一）。
  同じものが無ければ**断る**（黙って別のことをしない）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations


def word_is_missing(task: str, words) -> bool:
    """依頼文に、その op を許す語が 1 つも無いか。"""
    return not any(w and w in (task or "") for w in (words or ()))


def enforce(plan, task: str, rules: dict) -> tuple:
    """計画の各段を見て、語の裏付けが無い op を**読み直すか断る**。

    rules: {op: (許す語のタプル, 読み替える op or None, その op の日本語名)}
    戻り値: (直した計画, 人に見せる行のリスト, 断るなら True)
    ★ 判定はここ 1 箇所 ── 呼び出し側は材料を渡すだけ（op ごとの if を書かない）。
    """
    out, lines, refuse = [], [], False
    for step in (plan or []):
        op = str((step or {}).get("op") or "")
        rule = rules.get(op)
        if not rule or not word_is_missing(task, rule[0]):
            out.append(step)
            continue
        words = "』『".join(rule[0])
        if rule[1]:
            lines.append(f"（『{rule[2]}』として読み直しました ── 依頼文に"
                         f"『{words}』の語がありません）")
            out.append({"op": rule[1], "args": dict((step or {}).get("args") or {})})
            continue
        lines.append(f"？ この操作は依頼文に『{words}』と書かれた時だけ使えます "
                     "── 頼みたいなら、その語を入れて言い直してください")
        refuse = True
        out.append(step)
    return out, lines, refuse
