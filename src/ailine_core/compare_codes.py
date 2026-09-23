"""比較の語彙（EXTRACT / SET_WHERE の cmp）── 名簿・人に見せる名・Basic の番号の 3 つ組。

★ 2026-09-23 に src/ailine/__init__.py から**移しただけ**。3 つは一緒に直す約束なので同じ所に置く
  （コード生成だけが番号を使うが、番号だけを連れ出すと組が割れる）。
"""
from __future__ import annotations



# ★ EXTRACT: 比較の語彙（設計書どおり6種）。gte/lte/gt/lt は数値比較・eq は値の型に応じて
#   数値/文字列どちらでも・contains は常に文字列の部分一致。
# ★ 2026-08-27: "in"（どれか）を足した。値は**一覧**（複数の名前）。
#   意味論は 3 箇所が同時に持つ: ここ / Basic の RowMatches Case 6 /
#   Python の _extract_predicate。凍結した真理値表 tests/test_predicate_truth_table.py
#   が 3 者の一致を縛る ── 変える時は必ず一緒に直すこと。
#: ★ 2026-09-05: "nin"（どれでもない）が名簿から漏れていた ── Basic の Case 7 も
#:   Python の別実装も凍結した真理表も既に持っているのに、**許可リストだけ古かった**。
#:   そのため EXTRACT は特別扱いの経路で通り、SET_WHERE は「その比較はありません」で
#:   断られていた（兄弟間の片配線）。
_EXTRACT_CMPS = ("gte", "lte", "gt", "lt", "eq", "contains", "in", "nin")


_EXTRACT_CMP_LABELS = {"gte": "以上", "lte": "以下", "gt": "超", "lt": "未満",
                        "eq": "等しい", "contains": "を含む", "in": "のどれか",
                        "nin": "のどれでもない"}


_EXTRACT_CMP_CODE = {"gte": 0, "lte": 1, "gt": 2, "lt": 3, "eq": 4, "contains": 5, "in": 6, "nin": 7}
