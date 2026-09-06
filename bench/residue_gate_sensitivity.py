# -*- coding: utf-8 -*-
"""残差の関所の **感度**（本物の事故で鳴るか）を測る ── 誤爆率の相方。

★★ なぜ要るか（2026-09-06）: 誤爆の測定（`residue_gate_probe.py`）が
  **156 件中 0 件**という綺麗な数字を出した。だが「一度も鳴らない関所」と
  「壊れた関所」は、誤爆だけを見ていても**見分けがつかない**。
  ★ 記憶の規範:「差が出ない時はまず測定器を疑う ── 陽性対照が通らない回は採用しない」。

  ★ とくに疑うべき理由が在る: 誤爆を 17 → 0 にしたのは絞り（リストの平坦化・
    `col:` 等の接頭辞落とし）で、これは**消費される語を増やす**方向の変更だ。
    増やしすぎれば本物の事故の語まで消費して黙る。ここはその裏返しを測る。

★ 検体は**手で組む**（LLM を起こさない）── 事故は「依頼に在った語が宣言から落ちる」形なので、
  落ちた宣言（当時の resolved）を直接書けば再現できる。★ 実物の LLM を通すと、
  いまは直っているので事故が再現しない（＝感度を測れない）。

★ この道具が測るのは **1 つの家系だけ**:
      「依頼に在る語（それも実表の**列名**）が、解決済みの宣言に現れない」

★★ 初回の測定（2026-09-06）で、俺の**分類が 2 つとも実測に負けた** ── そのまま書き残す:
  ① 陽性のつもりだった「取引先が東西商事の行を消して」（値だけ落ちる）は**黙った**。
     列名（取引先）は宣言に在り、落ちたのは値。関所は列名しか見ないので**原理的に見えない**。
     ★ これは検体の誤りではなく**関所の穴**。陰性側（BLIND）へ移した。
  ② 見逃すと思っていた「1 文字の値（『主』）」は**鳴った**。値そのものは拾えないが、
     列名『担当』が残差に出るため。★ 「1 文字は 2 組目に採れない」という朝の限界は、
     関所を置けば**黙って半分やることだけは防げる**。陽性側へ移した。

使い方:
    python bench/residue_gate_sensitivity.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import ailine  # noqa: E402
from ailine_core import residue  # noqa: E402

#: 実表の列（関所は「残差のうち実表の列名であるもの」だけで鳴る）
HEADERS = {"品名", "数量", "単価", "金額", "所属", "担当", "メモ", "氏名", "取引先"}


def _fires(task: str, op: str, resolved: dict) -> list:
    """関所が鳴った語を返す（本番の絞りと同じ規則）。"""
    flat = {}
    for k, v in resolved.items():
        if isinstance(v, str):
            flat[k] = v
        elif isinstance(v, (list, tuple, set)):
            for i, one in enumerate(v):
                if isinstance(one, str):
                    flat[f"{k}{i}"] = one
    left = residue.find_unconsumed_words(task, flat, ailine._op_match_pool(op))
    return [w for w in left if w in HEADERS]


#: 陽性対照 ── **鳴らなければ関所は無意味**（今日直した事故の当時の宣言を再現）
POSITIVE = [
    ("2条件の片落ち",
     "所属が営業で担当が主任の行のメモに「○」を付けて",
     "SET_WHERE",
     {"col": "メモ", "cond_col": "所属", "cmp": "eq", "cond_value": "営業", "value": "○"}),
    ("式が落ちて名前だけ",
     "数量と単価をかけた金額の列を作って",
     "ADD_COLUMN",
     {"name": "金額"}),
    ("行と列の取り違え",
     "金額の列を追加して",
     "INSERT_ROWS",
     {"at": "2", "count": "1"}),
    ("1文字の値は採れないが、採り損ねたことは見える",
     "所属が営業で担当が主の行のメモに「○」を付けて",
     "SET_WHERE",
     {"col": "メモ", "cond_col": "所属", "cmp": "eq", "cond_value": "営業", "value": "○"}),
]

#: 陰性対照 ── **鳴ってはいけない**（正しく解決できている回）
NEGATIVE = [
    ("1条件・正しい",
     "所属が営業の行のメモに「○」を付けて",
     "SET_WHERE",
     {"col": "メモ", "cond_col": "所属", "cmp": "eq", "cond_value": "営業", "value": "○"}),
    ("2条件・正しい（いまの版）",
     "所属が営業で担当が主任の行のメモに「○」を付けて",
     "SET_WHERE",
     {"col": "メモ", "cond_col": "所属", "cmp": "eq", "cond_value": "営業",
      "cond2_col": "担当", "cond2_value": "主任", "value": "○"}),
    ("式つきで正しい",
     "数量と単価をかけた金額の列を作って",
     "ADD_COLUMN",
     {"name": "金額", "operands": ["数量", "単価"], "formula_kind": "mul"}),
    ("並べ替え",
     "金額の大きい順に並べ替えて",
     "SORT",
     {"col": "金額", "order": "desc"}),
]

# ★★ 「値だけ落ちた」穴は**塞げないと測って決めた**（2026-09-06）:
#   関所を「列名」から「列名 ∪ 実表に在る値」へ広げた版を 246 件で測ったところ、
#   誤爆が 0/156 → **35/160（21%）**に跳ねた。事前に置いた線は「4 件以上なら関所に
#   しない」だったので却下。★ 誤爆の中身が理由をそのまま語っていた:
#
#       [在庫] ['ナット']  ← ナットの行を削除して
#       [名簿] ['鈴木']    ← 鈴木の所属を「東棟」にして
#
#   **行を値で指す依頼は必ず語が余る** ── 「ナットの行」は解決すると `row=3` という
#   数字になるので、『ナット』はどこにも消費されない。絞りの工夫で直る類ではなく、
#   *値で行を指す* という言い方そのものの性質。★ だから穴は**開けたまま開示して持つ**。
#   （再現: `python bench/residue_gate_probe.py --values`）

#: ★ 捕まえられないと**分かっている**もの（表に載せて自覚する）
BLIND = [
    ("並べ替えで式が値に潰れる ── 依頼の語は 1 つも落ちていない",
     "金額の大きい順に並べ替えて", "SORT", {"col": "金額", "order": "desc"}),
    ("落ちたのが**値だけ**の時 ── 関所は列名しか見ない",
     "取引先が東西商事の行を消して", "DELETE_ROWS", {"cond_col": "取引先"}),
]


def main() -> int:
    bad = 0
    print("★ 陽性対照（鳴らなければ関所は無意味）")
    for name, task, op, res in POSITIVE:
        got = _fires(task, op, res)
        mark = "鳴った" if got else "★黙った"
        if not got:
            bad += 1
        print(f"  [{mark}] {name}: {got}")
    print()
    print("★ 陰性対照（鳴ってはいけない）")
    for name, task, op, res in NEGATIVE:
        got = _fires(task, op, res)
        mark = "静か" if not got else "★誤爆"
        if got:
            bad += 1
        print(f"  [{mark}] {name}: {got}")
    print()
    print("★ 捕まえられないと分かっているもの（穴として持つ・失格にはしない）")
    for name, task, op, res in BLIND:
        print(f"  [{'鳴った' if _fires(task, op, res) else '見逃す'}] {name}")
    print()
    print(f"判定: {'★不合格 ' + str(bad) + ' 件' if bad else '合格（陽性で鳴り陰性で黙る）'}")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
