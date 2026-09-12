# -*- coding: utf-8 -*-
"""PDF 化した検体を、Excel と**同じ答え**で採点する（2026-09-12・設計 D10）。

    python bench/received_invoices/mk_pdf_corpus.py        # 先に PDF を作る
    python bench/received_invoices/score_pdf.py            # v2 87 冊
    python bench/received_invoices/score_pdf.py --self-test

★ 判定は既存の `score_v2.judge` をそのまま使う（物差しを作り替えない）。
★ 抽出は製品の経路そのもの: `form_read.read_pdf_book`（ページ選びも自分で決める）。

★★ 「PDF では持ち得ない情報」の台帳（`PDF_CANNOT_CARRY`）:
  凍結時の実測（設計 §1.3）で、残った外れのうち 4 件は**直せるバグではなく PDF の限界**だった。
  台帳に載せないと永久に赤く残り、「追いかけて正しい実装を壊す」方へ手が動く。
  ★ 台帳は**上限つき・1 件ずつ理由つき**。増やす時は「PDF が持たない情報は何か」を書けること。
  ★ 台帳の冊は**別枠で数える**（正に混ぜない・消さない）── 見えたまま、追いかけないだけ。
"""
from __future__ import annotations

import argparse
import collections
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(REPO / "src"))

import score_v2                                                     # noqa: E402
from ailine_core.field_record import GRADES_WITH_VALUE, describe, grade, value  # noqa: E402
from ailine_core.form_read import read_pdf_book                     # noqa: E402
from ailine_core.pdf_grid import NoTextLayer                        # noqa: E402

PDFS = HERE / "received_pdf" / "v2"
ANSWER = HERE / "答え_received_v2.json"

#: ★ PDF が物理的に持っていない情報で外れる冊（凍結 2026-09-12・上限 6）。
#:   形: (冊の id, 項目): 失う情報
PDF_CANNOT_CARRY = {
    ("T21", "請求額"): "数値 vs その通りの文字 ── Excel は『金額の欄が文字（¥1,320,000-）』で拒否する。PDF では区別が付かず、数に復元して値を出す",
    # ★ T34（全角の数字 `１，３２０，０００`）は当初ここに載せたが、部首だけを正規化する形にしたら
    #   全角の数字が文字のまま残り、Excel と同じ「金額の欄が文字」の拒否ができて**正**になった。
    #   --self-test が「正なのに台帳に載っている」と鳴らしたので理由ごと外した（2026-09-12・同日）。
    ("T33", "請求額"): "エラー状態 ── Excel は『D22 が #REF!（式が壊れている）』で拒否。PDF には見えず、もっともらしい値を出す",
    ("T23", "請求額"): "表示で丸められた精度 ── 真の値 13579.5 を 13,580 と描く。PDF から 0.5 は戻せない",
}
LEDGER_CAP = 6


def extract(path: Path) -> dict:
    recs = read_pdf_book(path)
    out = {}
    for fld, rec in recs.items():
        g = grade(rec)
        out[fld] = {"区分": g, "値": value(rec) if g in GRADES_WITH_VALUE else None,
                    "理由": rec.blank_reason, "説明": describe(rec)}
    return out


def score(show: int = 40) -> int:
    assert len(PDF_CANNOT_CARRY) <= LEDGER_CAP, "★ 台帳が上限を超えた ── 本当に PDF の限界か"
    books = [b for b in json.loads(ANSWER.read_text(encoding="utf-8")) if b.get("採用")]
    tally = collections.Counter()
    ledger_hits, misses, n_expect, no_text = [], [], 0, []
    for b in books:
        p = PDFS / (Path(b["file"]).stem + ".pdf")
        if not p.exists():
            print(f"★ 無い: {p.name}（先に mk_pdf_corpus.py）"); return 2
        try:
            got = extract(p)
        except NoTextLayer as e:
            got, _ = {}, no_text.append((b["id"], str(e)))
        for fld, want in (b.get("期待") or {}).items():
            n_expect += 1
            v, why = score_v2.judge(got.get(fld), want)
            if (b["id"], fld) in PDF_CANNOT_CARRY:
                ledger_hits.append((b["id"], fld, v, why[:80]))
                tally["台帳（PDF の限界）"] += 1
                continue
            tally[v] += 1
            if v != "正":
                misses.append((b["id"], fld, v, why))
    n_scored = n_expect - len(ledger_hits)
    print(f"PDF 化した {len(books)} 冊 / 期待 {n_expect} 件（うち台帳 {len(ledger_hits)} 件は別枠）"
          f"  ★ Excel では 171/171")
    for k in ("正", "誤報", "安全", "不親切", "誤区分", "無回答"):
        x = tally.get(k, 0)
        print(f"  {k:5s} {x:4d}  ({x / n_scored:5.1%})")
    print(f"  台帳   {len(ledger_hits):4d}  ← PDF が持ち得ない情報（別枠・追いかけない）")
    for bid, fld, v, why in ledger_hits:
        print(f"      {bid} {fld} [{v}] {PDF_CANNOT_CARRY[(bid, fld)][:60]}")
    if no_text:
        print(f"  テキスト層なし {len(no_text)} 冊: {[x[0] for x in no_text]}")
    if misses:
        print(f"\n外れ {len(misses)} 件（先頭 {show}）:")
        for bid, fld, v, why in misses[:show]:
            print(f"  {bid:>4s} {fld:5s} [{v}] {why[:100]}")
    return 1 if tally.get("誤報", 0) else 0


def self_test() -> int:
    """★ 台帳の番人: 台帳に載っている冊は、載せた理由どおりに**今も**外れていること。
    直ったのに台帳に残っていたら、それは台帳の腐り（消す時は理由ごと消す）。"""
    books = {b["id"]: b for b in json.loads(ANSWER.read_text(encoding="utf-8")) if b.get("採用")}
    bad = []
    for (bid, fld), why in PDF_CANNOT_CARRY.items():
        p = PDFS / (Path(books[bid]["file"]).stem + ".pdf")
        v, _ = score_v2.judge(extract(p).get(fld), books[bid]["期待"][fld])
        print(f"  台帳 {bid} {fld}: いま [{v}]  ── {why[:50]}")
        if v == "正":
            bad.append((bid, fld))
    if bad:
        print(f"★ 台帳の腐り: 正になっているのに載っている {bad} ── 理由ごと消すこと")
        return 1
    print("○ 台帳の冊は、載せた理由どおりに外れている")
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--show", type=int, default=40)
    a = ap.parse_args()
    if not PDFS.exists():
        print(f"★ 検体がありません（{PDFS}）。先に mk_pdf_corpus.py を走らせてください。")
        sys.exit(2)
    sys.exit(self_test() if a.self_test else score(a.show))
