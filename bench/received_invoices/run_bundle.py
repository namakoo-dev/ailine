# -*- coding: utf-8 -*-
"""束で疑う（forms_suspect）を束の検体にかけて採点する（2026-09-11）。

    python bench/received_invoices/run_bundle.py

★ 検体は生成物（gitignore）。先に `mk_received_bundle.py` を走らせること。
★ 凍結時の予測と実測（docs/DESIGN-20260910 §4.4）:
  正 5（重複・秋月の再発行・年の誤り・桁違い・竹村の空欄）／
  取り逃し 3（大和田の再発行と磯部の番号重なり＝inv21 の骨で日付も番号も読めない、
  番号がとぶ＝意図して入れない）／ 偽の疑い 1（檜山 06 の令和が読めず「空欄」に見える）
"""
from __future__ import annotations

import copy
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(REPO / "src"))

import openpyxl                                   # noqa: E402
import score_bundle                               # noqa: E402
from bundle_common import load_bundles            # noqa: E402
from ailine_core.forms_suspect import suspect     # noqa: E402
from ailine_core.field_record import value        # noqa: E402
from ailine_core.form_read import read_book       # noqa: E402

ANSWER, BOOKS = "答え_received_bundle.json", "received_bundle"
_TABLE: dict = {}    # 束名 → {冊: {項目: 値}}（読むのは 1 回）


def table_for(folder: Path) -> dict:
    if folder.name not in _TABLE:
        t = {}
        for p in sorted(folder.glob("*.xlsx")):
            wb = openpyxl.load_workbook(p, data_only=True)
            wbf = openpyxl.load_workbook(p, data_only=False)
            try:
                recs = read_book(wb, wb_formula=wbf)
            finally:
                wb.close(); wbf.close()
            t[p.name] = {f: value(r) for f, r in recs.items()}
        _TABLE[folder.name] = t
    return _TABLE[folder.name]


def tool(folder: Path) -> list:
    return suspect(table_for(folder))


if __name__ == "__main__":
    if not (HERE / ANSWER).exists():
        print(f"★ 検体がありません（{ANSWER}）。先に mk_received_bundle.py を走らせてください。")
        sys.exit(2)
    if score_bundle.self_test(HERE, ANSWER, BOOKS) != 0:
        sys.exit(1)
    res = score_bundle.score(tool, HERE, ANSWER, BOOKS)
    print("\n=== いまの器官 + 束で疑う ===")
    score_bundle.report(res, show=40)
    print("\n出した所見:")
    for b in load_bundles(HERE / ANSWER):
        for s in tool(HERE / BOOKS / b["束"]):
            print(f"  [{b['束']}] {s['種類']} {s['冊']}\n      {s['理由']}")

    # ── 変異（束の論理だけを動かす: 器官が出した値の表を 1 か所いじる） ──
    t8 = _TABLE["2026年8月受領分"]
    dup_a, dup_b = "深田金属株式会社_2026-08.xlsx", "深田金属株式会社_2026-08 (2).xlsx"
    kinds0 = {(s["種類"], tuple(sorted(s["冊"]))) for s in suspect(t8)}
    m1 = copy.deepcopy(t8); m1[dup_b]["請求額"] = m1[dup_b]["請求額"] + 1
    kinds1 = {(s["種類"], tuple(sorted(s["冊"]))) for s in suspect(m1)}
    m2 = copy.deepcopy(t8); del m2[dup_b]
    kinds2 = {(s["種類"], tuple(sorted(s["冊"]))) for s in suspect(m2)}
    pair = tuple(sorted((dup_a, dup_b)))
    print("\n変異:")
    print("  金額を 1 円ずらす → 重複が消えて 訂正再発行 になる:",
          "○" if ("重複", pair) in kinds0 and ("重複", pair) not in kinds1 and ("訂正再発行", pair) in kinds1 else "★")
    print("  片方の冊を消す     → 重複が消える:",
          "○" if ("重複", pair) not in kinds2 and not any(dup_b in k[1] for k in kinds2) else "★")
