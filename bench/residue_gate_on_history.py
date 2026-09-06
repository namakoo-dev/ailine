# -*- coding: utf-8 -*-
"""残差の関所を **本物の走行記録の上で** 測る（模型ではなく実物）。

★★ なぜこれが要るか（2026-09-06・Namakoo「残り 90 はどうする？」から出た）:
  `residue_gate_probe.py` は翻訳と解決を自前で呼ぶ**模型**で、246 件中 155 件しか
  届かなかった（実物の機械は 238 件を通す）。届かない 91 件を黙って捨てると、
  誤爆率は**楽な側に偏った分母**で出る。二段目を当てて埋めようとしたが、検体の
  第 1 要素は op ではなく**検査の分類名**（`row_add` 等）で、33 種中 6 種しか
  op に対応しない ── 手で対応表を書けば俺の推測が測定器に混じる。

  ★ そこで模型を継ぎ足すのをやめ、**本物の上で測る**ことにした。材料は既に在った:
    `~/.ailine/history.jsonl` が全実行の「依頼(task)」と「解釈行(command)」を残している。

★★ ここで設計が 1 つ変わった ── **宣言は解釈行の方**:
    引数の dict を宣言だと思って測っていたが、人に見せている宣言はもっと厚い。

        解釈: 操作:行削除 削除位置:3 位置の根拠:『鈴木』の行＝3行目 行数:1

    行を値で指した『鈴木』は、引数では `row=3` という数字に化けて消えるが、
    **解釈行には根拠として残っている**。だから解釈行を宣言に採ると、
    「値で行を指す依頼は必ず語が余る」という誤爆の family がまるごと消える。

★★ 測った結果（2026-09-02〜09-06 の実走行・成功した回だけ）:

      3179 件中 **28 件が鳴った（0.88%）／ そのすべてが本物の欠陥**

    しかも自然実験になっていた。同じ依頼「所属が営業の行のメモに「○」を付けて」で:

      09-04 06:26 〜 09-05 03:34   op=ADD_ROW    ← 条件つき書換を**行追加**と誤読
      09-05 06:01 以降             op=SET_WHERE  ← 正しい

    ★ **事後条件は前後どちらも `pass`** と言っている ── 実体は宣言どおりだから。
      見えなかったのが事後条件で、見えたのが残差。これがこの関所の存在理由そのもの。

★ 家系の外（鳴らないと分かっているもの）も、同じ記録の中に実例が在った:

      09-05 06:06  op=SET_WHERE  比べ方:**等しい**   ← 依頼は「営業以外」（否定の反転）

    本物の誤りだが**語は 1 つも落ちていない**ので関所は黙る。捕まえるのは
    「依頼に在る**列名**が宣言に出ない」形だけ ── 広げようとして失敗した記録は
    `residue_gate_sensitivity.py` の冒頭にある（`--values` は誤爆 21%）。

使い方:
    python bench/residue_gate_on_history.py                # 09-02 以降
    python bench/residue_gate_on_history.py --since 2026-09-05T06:00
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "bench"))

import ailine  # noqa: E402
from ailine_core import residue  # noqa: E402

import basic_ops_matrix as M  # noqa: E402

HISTORY = Path.home() / ".ailine" / "history.jsonl"


def _task_to_table() -> dict:
    """依頼文 → その検体が属する表。★ 2 つ以上の表で使い回している依頼は**除く**
       （列名の集合が決まらないため ── 合わせると広がって誤爆側に倒れる）。"""
    m = defaultdict(set)
    for key in M.TABLES:
        for _slug, task, _check in M._cases_for(key):
            m[task].add(key)
    return {t: next(iter(ks)) for t, ks in m.items() if len(ks) == 1}


def fired_words(task: str, op: str, command: str, headers: set) -> list:
    """関所が鳴った語（＝依頼に在り・解釈行に出ず・実表の列名であるもの）。"""
    left = residue.find_unconsumed_words(task, {}, ailine._op_match_pool(op or ""))
    return [w for w in left if w not in (command or "") and w in headers]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", default="2026-09-02")
    ap.add_argument("--until", default="9999")
    a = ap.parse_args()

    if not HISTORY.is_file():
        print("履歴が無い:", HISTORY)
        return 0
    t2k = _task_to_table()
    heads = {k: {str(h) for h in M.TABLES[k]["headers"] if h} for k in M.TABLES}

    n = skipped = 0
    fired = Counter()
    for line in HISTORY.read_text(encoding="utf-8", errors="replace").splitlines():
        try:
            e = json.loads(line)
        except Exception:
            continue
        ts = str(e.get("ts") or "")
        if not (a.since <= ts < a.until):
            continue
        if e.get("ok") is not True or not e.get("command"):
            continue
        task = e.get("task")
        if task not in t2k:
            skipped += 1
            continue
        n += 1
        hit = fired_words(task, str(e.get("op") or ""), str(e["command"]), heads[t2k[task]])
        if hit:
            fired[(task, tuple(hit), str(e.get("op") or ""))] += 1

    tot = sum(fired.values())
    print(f"窓 {a.since} 〜 {a.until}")
    print(f"成功した本物の走行 {n} 件（検体に無い/表が定まらない依頼 {skipped} 件は除外）")
    print(f"★ 関所が鳴った: {tot} 件 = {tot * 100 / max(1, n):.2f}%")
    for (task, words, op), c in fired.most_common(12):
        print(f"    x{c:4d} {list(words)}  op={op}  ← {task[:40]}")
    print()
    print("★ 鳴った回は**一件ずつ中身を見ること** ── 誤爆か本物かは数では決まらない。")
    print("★ 事後条件が `pass` でも鳴りうる（それがこの関所の存在理由）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
