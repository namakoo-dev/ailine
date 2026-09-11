# -*- coding: utf-8 -*-
"""帳票を読む器官を検体にかけて採点する（2026-09-11）。

    python bench/received_invoices/run_organ.py                 # 開発群（基礎 36 冊）
    python bench/received_invoices/run_organ.py --groups '*'    # 全 87 冊

★★ 既定が「基礎 36 冊」なのはわざと。規則はこの群だけを見て書いた。
  残り 51 冊（敵対・税率・混入）は**凍結してから**開いたもので、
  そこで出た数字（正 43/63・誤報 6）が一般化についての唯一の証拠だった。
  ★ その後 51 冊を見て直したので、**いまの 171/171 は一般化の証拠ではない**。
    一般化を言うには新しい未見の検体が要る（docs/開発手法.md・tuning-audit）。

★ 検体は生成物（gitignore）。先に `mk_received_v2.py` を走らせること。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(REPO / "src"))

import score_v2                                        # noqa: E402
from ailine_core.field_record import GRADES_WITH_VALUE, describe, grade, value  # noqa: E402
from ailine_core.form_read import read_book                  # noqa: E402


def make_extractor():
    """検体 1 冊 → 採点器の契約（項目 → 区分・値・理由）。

    ★★ 引数を持たないのはわざと。2026-09-11 まで、どのシートを読むかを
      **検体の答えから受け取っていた**（`sheet_of`）。つまり「どのシートが
      請求書か」を一度も解かずに満点を出していた ── 買い手が持っていない
      手がかりで測っていた。渡せる口を残すと、また渡す。**口ごと消す。**
    """
    def extract(path: Path) -> dict:
        wb = openpyxl.load_workbook(path, data_only=True)
        wbf = openpyxl.load_workbook(path, data_only=False)
        try:
            recs = read_book(wb, wb_formula=wbf)
        finally:
            wb.close(); wbf.close()
        out = {}
        for fld, rec in recs.items():
            g = grade(rec)
            out[fld] = {"区分": g,
                        "値": value(rec) if g in GRADES_WITH_VALUE else None,
                        "理由": rec.blank_reason,
                        "説明": describe(rec)}
        return out
    return extract


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--groups", default="基礎",
                    help="測る群（基礎 / 敵対 / 税率 / 混入 / '*' で全部）")
    ap.add_argument("--floor", type=int, default=60,
                    help="分母の床。★ 空集合に○を付けないため")
    ap.add_argument("--show", type=int, default=40)
    ap.add_argument("--answer", default=score_v2.ANSWER_FILE,
                    help="答えの JSON（別の検体の束を測る時に差し替える）")
    ap.add_argument("--books", default=score_v2.BOOKS_DIR, help="冊の置き場")
    a = ap.parse_args()
    score_v2.use_bundle(a.answer, a.books)

    if a.groups and a.groups != "*":
        score_v2._ONLY_GROUPS = tuple(a.groups.split(","))
    score_v2.MIN_EXPECTATIONS = a.floor

    corpus = score_v2.DEFAULT_CORPUS
    if not (corpus / score_v2.ANSWER_FILE).exists():
        print(f"★ 検体がありません（{score_v2.ANSWER_FILE}）。先に生成器を走らせてください。")
        sys.exit(2)

    books = score_v2.load_books(corpus)
    print(f"★ 測る群: {a.groups}（{len(books)} 冊）")

    # ★ 採点器そのものを先に確かめる（この群でも対照と変異を通す）
    rc = score_v2.self_test(corpus)
    if rc != 0:
        print("★ 採点器が対照を通らない。点数は見ない。")
        sys.exit(rc)
    print()

    res = score_v2.score(make_extractor(), corpus)
    score_v2.report(res, show=a.show)
