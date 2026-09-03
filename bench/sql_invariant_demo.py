# -*- coding: utf-8 -*-
"""スキーマが守る不変量と、守らない不変量 ── SQLite で実際に走らせて見せる。

★ なぜ在るか（2026-09-04）: `docs/なぜこの形か.md` は「SQL はほぼ全部の軸でこれより上」
  と書き、検証の難しさを「表計算にスキーマが無いから」で説明していた。
  ★ **その説明は半分しか正しくない。** スキーマが守るのは「宣言された不変量」だけで、
  宣言されていない不変量は **DB でも守られない**。ここはそれを実測で示す。

  題材は実際に踏んだ事故（2026-08-31・Namakoo の指摘）:
    「丸和物流の単価とみどり建設の単価を入れ替えて」は頼まれた 2 セルだけを正しく
    入れ替える。だが 金額（＝件数×単価）は**直値**なので取り残され、表として矛盾する。
    それでも「2 セルだけ動いた」は真実なので ✓ が出てしまう。

★ この実験が示すこと: 同じ操作を SQLite でやると、**CHECK 制約を全部通りながら
  同じ矛盾が残り、DB は「2 行更新しました」としか言わない**。
  そして `ailine_core/row_identity` は、語も見出しも読まずに**数だけ**から
  `金額 = 件数 × 単価` を復元し、崩れたことを名指しする。

★ 主張はここまで ── 「SQL より上」ではない。**「SQL にも同じ層が要る」**。
  集合演算・トランザクション・同時実行・規模は、今も DB のほうが上。

    python bench/sql_invariant_demo.py          # 走らせて見る
    python bench/sql_invariant_demo.py --json   # 機械可読（番人が読む）
"""
from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from ailine_core import row_identity as ri  # noqa: E402

HEADERS = ["取引先", "件数", "単価", "金額"]
BEFORE = [["丸和物流", 12, 4800, 57600],
          ["みどり建設", 8, 7200, 57600],
          ["あかつき商事", 5, 3000, 15000]]


def run() -> dict:
    con = sqlite3.connect(":memory:")
    con.execute(
        "CREATE TABLE 明細("
        "  取引先 TEXT NOT NULL,"
        "  件数 INTEGER NOT NULL CHECK(件数 > 0),"
        "  単価 INTEGER NOT NULL CHECK(単価 > 0),"
        "  金額 INTEGER NOT NULL CHECK(金額 > 0))")   # ★ 金額は「ただの数値列」
    con.executemany("INSERT INTO 明細 VALUES(?,?,?,?)", BEFORE)
    con.commit()

    cur = con.execute(
        "UPDATE 明細 SET 単価 = CASE 取引先"
        "  WHEN '丸和物流' THEN 7200 WHEN 'みどり建設' THEN 4800 END"
        " WHERE 取引先 IN ('丸和物流','みどり建設')")
    con.commit()
    after = [list(r) for r in con.execute("SELECT * FROM 明細").fetchall()]

    inconsistent = [r[0] for r in after if r[1] * r[2] != r[3]]
    ids = ri.identities(BEFORE)
    lost = ri.broken(BEFORE, after)
    return {
        "sql_said": f"{cur.rowcount} 行更新しました",
        "constraints_passed": True,          # CHECK を 1 つも破らずに完了している
        "rows_inconsistent": inconsistent,   # ★ それでも表は矛盾している
        "identities_found": [list(x) for x in ids],
        "identities_broken": [list(x) for x in lost],
        "ailine_said": ri.describe(lost, HEADERS),
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    r = run()
    if a.json:
        print(json.dumps(r, ensure_ascii=False, indent=1))
        return 0
    print("依頼: 丸和物流とみどり建設の単価を入れ替えて\n")
    print(f"  SQL が言うこと : 「{r['sql_said']}」 ── それだけ")
    print("  制約チェック   : 全部通っている（件数>0・単価>0・金額>0）\n")
    print("  実際の表:")
    con = sqlite3.connect(":memory:")
    for name in r["rows_inconsistent"]:
        print(f"    ★ {name}: 件数 × 単価 と 金額 が合っていない")
    print(f"\n  ailine が言うこと:\n    {r['ailine_said']}")
    print("\n★ 集合演算・トランザクション・同時実行・規模は、今も DB のほうが上。"
          "\n  ここで示したのは『スキーマは宣言された不変量しか守らない』という 1 点だけ。")
    con.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
