# -*- coding: utf-8 -*-
"""束の検体を、いまの器官にかけて採点する（2026-09-11）。

    python bench/received_invoices/run_bundle.py            # ①②③
    python bench/received_invoices/run_bundle.py --flat     # ① だけ

  ① 冊ごとの 5 項目 ── 束の答えの「冊」を平らにして既存の `score_v2` で採点
  ② 束の疑い ── `forms_suspect` を `score_bundle` で採点（正／偽の疑い／取り逃し／不親切）
  ③ 材料の一覧 ── 束ごとに 請求元/請求日/請求番号/請求額 の区分と値（設計の目で見る用）

★ 検体は生成物（gitignore）。先に `mk_received_bundle.py` を走らせること。
★ 凍結時の予測と実測（docs/DESIGN-20260910 §4.4・§4.8）。
"""
from __future__ import annotations

import argparse
import copy
import json
import shutil
import sys
import tempfile
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(REPO / "src"))

import openpyxl                                          # noqa: E402
import score_bundle                                      # noqa: E402
import score_v2                                          # noqa: E402
from bundle_common import answer_value, load_bundles, norm_date   # noqa: E402
from run_organ import make_extractor                     # noqa: E402
from ailine_core.field_record import grade, value        # noqa: E402
from ailine_core.form_read import read_book              # noqa: E402
from ailine_core.forms_suspect import suspect            # noqa: E402

ANSWER, BOOKS = "答え_received_bundle.json", "received_bundle"
FIELDS = ("請求元", "宛先", "請求額", "請求日", "請求番号")
_TABLE: dict = {}    # 束名 → {冊: {項目: 値}}（読むのは 1 回）


# ── ① 冊ごとの 5 項目 ──────────────────────────────────────
def flatten(corpus: Path, workdir: Path) -> tuple:
    """束の答えの 冊 を、`score_v2` が読む平らな形にして作業場へ写す。

    ★ 既存の採点器は冊を **ファイル名**で引くので、束の階層を `束__冊名` に畳む。
    ★ 書き手は区分を宣言していない（答えは値と番地だけ）── 値が在れば「値を出す区分なら
      どれでも」、無ければ「無（理由つき）」。**値の正誤だけ**を測る。
    """
    flat = []
    (workdir / "books").mkdir(parents=True, exist_ok=True)
    for b in load_bundles(corpus / ANSWER):
        for x in b.get("冊") or []:
            y = dict(x)
            y["file"] = f"{b['束']}__{x['file']}"
            y["id"] = f"{b['束']}/{x.get('id') or Path(x['file']).stem}"
            y["採用"] = True
            shutil.copyfile(corpus / BOOKS / b["束"] / x["file"], workdir / "books" / y["file"])
            if not y.get("期待"):
                exp = {}
                for f in FIELDS:
                    a = x.get(f)
                    if not isinstance(a, dict):
                        continue
                    v = answer_value(f, a.get("値"))
                    if f == "請求日":
                        v = norm_date(v)
                    exp[f] = ({"区分": ["無"], "値": None, "名指し": []} if v in (None, "")
                              else {"区分": list(score_v2.VALUE_GRADES), "値": v})
                y["期待"] = exp
            flat.append(y)
    (workdir / ANSWER).write_text(json.dumps(flat, ensure_ascii=False, indent=1), encoding="utf-8")
    return len(flat)


def score_fields(corpus: Path, show: int) -> int:
    with tempfile.TemporaryDirectory(prefix="ailine_bundle_") as d:
        work = Path(d)
        n = flatten(corpus, work)
        print(f"① 冊ごとの項目 ── 平らにした冊 {n}")
        score_v2.use_bundle(ANSWER, "books")
        if score_v2.self_test(work) != 0:
            print("★ 採点器の対照が通らない ── 点数を見ない")
            return 1
        _ex = make_extractor()

        def extract(path):
            out = _ex(path)
            if out and out.get("請求日"):
                # ★ 答えの側と同じ形に均すだけ（値は変えない）
                out["請求日"]["値"] = norm_date(out["請求日"]["値"])
            return out

        score_v2.report(score_v2.score(extract, work), show=show)
    return 0


# ── ②③ 束の疑いと材料 ─────────────────────────────────────
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


def materials(corpus: Path) -> None:
    for b in load_bundles(corpus / ANSWER):
        print(f"\n=== 束 {b['束']}  冊 {len(b.get('冊') or [])}  仕込み {len(b.get('疑い') or [])}"
              f"  怪しくない {len(b.get('怪しくない') or [])} ===")
        for w in b.get("疑い") or []:
            print(f"   仕込み: {w['種類']} {w['冊']} 名指し={w.get('名指し')}")
        folder = corpus / BOOKS / b["束"]
        for name in sorted(table_for(folder)):
            wb = openpyxl.load_workbook(folder / name, data_only=True)
            wbf = openpyxl.load_workbook(folder / name, data_only=False)
            try:
                recs = read_book(wb, wb_formula=wbf)
            finally:
                wb.close(); wbf.close()
            cells = [f"{f}={grade(recs[f])}:{value(recs[f])!r}"
                     for f in ("請求元", "請求日", "請求番号", "請求額") if f in recs]
            print(f"   {name:34s} " + "  ".join(cells))


def mutations(corpus: Path) -> None:
    """★ 束の論理だけを動かす（器官が出した値の表を 1 か所いじる）。"""
    t8 = table_for(corpus / BOOKS / "2026年8月受領分")
    dup_a, dup_b = "深田金属株式会社_2026-08.xlsx", "深田金属株式会社_2026-08 (2).xlsx"
    kinds = lambda t: {(s["種類"], tuple(sorted(s["冊"]))) for s in suspect(t)}   # noqa: E731
    k0 = kinds(t8)
    m1 = copy.deepcopy(t8); m1[dup_b]["請求額"] = m1[dup_b]["請求額"] + 1
    k1 = kinds(m1)
    m2 = copy.deepcopy(t8); del m2[dup_b]
    k2 = kinds(m2)
    pair = tuple(sorted((dup_a, dup_b)))
    print("\n変異:")
    print("  金額を 1 円ずらす → 重複が消えて 訂正再発行 になる:",
          "○" if ("重複", pair) in k0 and ("重複", pair) not in k1 and ("訂正再発行", pair) in k1 else "★")
    print("  片方の冊を消す     → 重複が消える:",
          "○" if ("重複", pair) not in k2 and not any(dup_b in k[1] for k in k2) else "★")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=HERE)
    ap.add_argument("--flat", action="store_true", help="① だけ")
    ap.add_argument("--show", type=int, default=40)
    a = ap.parse_args()
    if not (a.corpus / ANSWER).exists():
        print(f"★ 検体がありません（{ANSWER}）。先に mk_received_bundle.py を走らせてください。")
        sys.exit(2)

    if score_fields(a.corpus, a.show) != 0:
        sys.exit(1)
    if a.flat:
        sys.exit(0)

    print("\n② 束の疑い")
    if score_bundle.self_test(a.corpus, ANSWER, BOOKS) != 0:
        sys.exit(1)
    score_bundle.report(score_bundle.score(tool, a.corpus, ANSWER, BOOKS), show=a.show)
    print("\n出した所見:")
    for b in load_bundles(a.corpus / ANSWER):
        for s in tool(a.corpus / BOOKS / b["束"]):
            print(f"  [{b['束']}] {s['種類']} {s['冊']}\n      {s['理由']}")
    mutations(a.corpus)

    print("\n③ 材料")
    materials(a.corpus)
