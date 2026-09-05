# -*- coding: utf-8 -*-
"""属性が「語を接地できる先」をどれだけ覆っているかを測る（2026-09-05）。

★★ なぜスクリプトにして残すか（Namakoo「実測は大事だな。以前測ったところが後に効いて
  くる場面は何度もあった」を受けて、同じ日に自分の測定を見直した）:

  この測定の結果（29 語中 17 が当たった・当たらない 12 の内訳）は commit と試験の
  docstring に書いた。**だが数字だけだった** ── どの表にどの依頼文を当てたのかが
  残っていないので、**半年後の自分が再現できない**し、条件が変われば意味が変わる。

  ★ 今日はまさにその形で 2 回誤読した:
      「語彙外は 2.6%」    → 発火率だと読んだ（実は天井・実際は約 1%）
      「残差に 7 件の穴」  → 語彙の穴だと読んだ（時間で切ったらどれも閉じていた）
    効いた測定（述語の真理表・codegen のゴールデン）は**条件ごと凍って**いた。
    誤読した測定は**条件を持たない生の数字**だった。

★ だからここに検体（表と依頼文）を置く。数字は走らせれば出る ── 記憶にも文書にも
  数字は書かない（[[reference_ailine_shelf]] の掟と同じ）。

使い方:
    python bench/attribute_coverage_probe.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import openpyxl  # noqa: E402

import ailine  # noqa: E402
from ailine_core import attributes as A  # noqa: E402
from ailine_core import residue  # noqa: E402
from ailine_core.op_axes import AXES  # noqa: E402

#: ★ 検体の表（この形でないと数字の意味が変わる ── 列名が依頼文の語と噛み合うこと）
SHEET = "売上一覧"
HEADERS = ["納品日", "得意先", "部門", "品番", "品名", "数量", "単価", "金額", "備考"]
ROWS = [
    ["2026-01-05", "山田商事", "営業1課", "B-100", "ボルト", 10, 50, 500, ""],
    ["2026-01-06", "鈴木工業", "営業2課", "N-200", "ナット", 5, 20, 100, "確認済"],
    ["2026-01-07", "山田商事", "営業1課", "B-100", "ボルト", 3, 50, 150, ""],
    ["2026-02-01", "佐藤製作所", "購買", "W-300", "ワッシャ", 8, 15, 120, ""],
    ["", "", "", "", "", "", "合計", 870, ""],
]

#: ★ 依頼文（実務で出る形を並べる ── 通る/通らないを混ぜる）
TASKS = [
    "売上シートの平均を出して",
    "単価の平均値を一番下に追加して",
    "ボルトの行だけ抜き出して",
    "原価の列を消して",
    "消費税込みの金額を出して",
    "部門ごとに金額をまとめて",
    "得意先の一覧をシートにして",
    "小計の行を太字にして",
    "備考を全部「確認済」にして",
    "税率を8%にして",
    "去年のファイルと突き合わせて",
    "品番の重複を除いて",
    "数量が10以上の行に○を付けて",
    "金額を円マーク付きにして",
    "納品日の古い順に並べ替えて",
]


def build_book(dirpath: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r)
    p = dirpath / "probe.xlsx"
    wb.save(p)
    return p


def main() -> int:
    d = Path(tempfile.mkdtemp())
    book = build_book(d)
    bm = ailine.build_book_meta(book)
    sheets = bm.get("sheets") or []
    headers = bm.get("headers") or {}
    samples = A.sample_columns(book, headers, bm.get("header_rows"))
    lacks = {}
    for ax in AXES.values():
        lacks.update(getattr(ax, "lacks", {}) or {})
    pool = [p for op in ailine.OP_META for p in ailine._op_match_pool(op) if p]

    total = hit = 0
    misses = {}
    print(f"検体: {SHEET}（{len(HEADERS)} 列 × {len(ROWS)} 行）／ 依頼文 {len(TASKS)} 本")
    print(f"照合語彙 {len(pool)} 語 ／ 属性 {len(A.KINDS)} 種: {', '.join(A.KINDS)}")
    print()
    for task in TASKS:
        marks = []
        for word in residue.find_unconsumed_words(task, {}, pool):
            cands = A.candidates_for(word, sheets=sheets, headers=headers,
                                     samples=samples, lacks=lacks)
            total += 1
            if cands:
                hit += 1
                marks.append(f"{word}={len(cands)}")
            else:
                misses[word] = misses.get(word, 0) + 1
                marks.append(f"{word}=×")
        print(f"  {' '.join(marks):<44} | {task}")
    print()
    print(f"残差の語 {total} 件中、属性に当たった {hit} 件"
          f"（{hit * 100 // max(1, total)}%）")
    print("当たらなかった語:", ", ".join(
        f"{w}×{n}" if n > 1 else w for w, n in sorted(misses.items(), key=lambda x: -x[1])))
    print()
    print("★ 数字だけを持ち出さないこと ── 上の表と依頼文が変われば、この率は意味を変える。")
    print("★ 当たらない語の内訳は『属性が足りない』とは限らない（2026-09-05 の実測では")
    print("   12 件中 7 件が op の照合語彙の穴で、属性を増やしても解けなかった）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
