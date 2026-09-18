# -*- coding: utf-8 -*-
"""記法の盤 ── 依頼文の「どこを指すか」を、実物の器官に通して表にする（2026-09-18）。

★★ なぜ要るか（Namakoo 2026-09-18:「三項を満たしているかをチェックする機構はあるか？」）:
  三項（依頼・宣言・実体）のうち、**宣言**の項には既に分母がある ──
  plan_writes_beyond_one_cell は OP_WRITE_TARGET から導き、
  test_every_value_writing_op_is_covered_by_the_gate が漏れを赤にする。
  ところが**依頼**の項には分母が無かった。task_* な読み手は 16 個あるのに、
  それを束ねる名簿も、どの記法を読めるかの表も無い。

★★ 依頼の項が欠けても鳴らないのが致命だった:
  task_points_at_one_row が None を返した時、それは 2 通りの意味を持つ ──
    (a) 依頼は本当に 1 か所を指していない   → 一括書換でよい
    (b) 指しているが、読めない記法だった     → 一括書換は事故
  画面上この 2 つは区別がつかない。2026-09-18 の盲検 4 体目 ④ は (b) で、
  空の列を 6 行潰して exit 0 になった（空だから上書きの関所も鳴らない）。
  ★ 棚の線「出ないことは信号でない」そのもの。

★ この盤がやること: 名簿の検体を**実物の関数**に通し、宣言と実測を突き合わせる。
  ずれたら番人（tests/test_notation_board.py）が赤くなる。
  ★ 名簿の status を手で書き換えても、実物が変わらなければ赤のままになる。

★ なぜ tests/ に在るか: 素の環境の番人（scripts/_ci_parity_blocker.py）は
  requirements-dev.txt に無い import を全部止める。scripts/ に置いて import すると
  **自分の repo の道具なのに弾かれる**（2026-09-16 と 2026-09-18 に実測）。
  中身はここ、scripts/notation_board.py は薄い入口だけにする。
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import openpyxl  # noqa: E402

import ailine  # noqa: E402

REGISTER = REPO / "tests" / "notation_register.json"

#: 盤が使う検体の冊。★ 固定する ── 冊が変わると「行の名前」の実測が動く。
#:   同名 2 行（丸山重工）を必ず含める: 「指してはいるが決められない」を測る唯一の形。
SPECIMEN_ROWS = (("大東金属", 100), ("みどり商事", 200),
                  ("丸山重工", 300), ("丸山重工", 400))
SPECIMEN_HEADERS = ("取引先", "金額", "担当")

STATUS_MARK = {
    "reads": "○ 読む",
    "gap": "★ 在庫（1 か所を指すのに読めない）",
    "narrower_than_column": "★ 在庫（列より狭いのに読めない）",
    "by_design": "─ 列全体（読まなくてよい）",
}


def load_register() -> dict:
    return json.loads(REGISTER.read_bytes().decode("utf-8"))


def _make_specimen_book(d: Path) -> Path:
    p = d / "帳簿.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "仕訳"
    ws.append(list(SPECIMEN_HEADERS))
    for name, amt in SPECIMEN_ROWS:
        ws.append([name, amt, None])
    wb.save(p)
    wb.close()
    return p


def measure() -> dict:
    """検体 → 実物の task_points_at_one_row が 1 か所と認めたか（と、その理由の文）。

    ★ 測るのは**製品の関数**であって、名簿の写しではない（恒真殺し）。
    """
    reg = load_register()
    out = {}
    with tempfile.TemporaryDirectory() as td:
        book = _make_specimen_book(Path(td))
        meta = ailine.build_book_meta(book)
        sheet = meta["sheets"][0]
        for key, rec in reg["notations"].items():
            rows = []
            for sp in rec["specimens"]:
                said = ailine.task_points_at_one_row(sp["task"], meta, sheet)
                rows.append({"task": sp["task"], "declared": bool(sp["reads"]),
                              "measured": said is not None, "said": said or "",
                              # ★★ 2026-09-18: 「読めたか」だけを測ると、**どこを指したか**が
                              #   ずれる退行（表の端で _last+1 を返す等）が緑のまま通る。
                              #   says を書いた検体は、理由の文にその語が在ることまで見る。
                              "says": sp.get("says", ""),
                              "note": sp.get("note", "")})
            out[key] = rows
    return out


def survey() -> dict:
    reg = load_register()
    got = measure()
    notations = []
    for key, rec in reg["notations"].items():
        rows = got[key]
        notations.append({
            "key": key,
            "label": rec["label"],
            "points_at": rec["points_at"],
            "status": rec["status"],
            "organ": rec.get("organ", ""),
            "unlock": rec.get("unlock", ""),
            "note": rec.get("note", ""),
            "specimens": rows,
            "reads_any": any(r["measured"] for r in rows),
        })
    return {
        "commit": subprocess.run(["git", "log", "-1", "--format=%h %ad %s", "--date=short"],
                                  cwd=str(REPO), capture_output=True, text=True,
                                  encoding="utf-8").stdout.strip(),
        "notations": notations,
    }


def mismatches(data: dict) -> list:
    """宣言と実測のずれ。★ 3 種類あり、どれも別の事故を意味する。"""
    out = []
    for n in data["notations"]:
        for sp in n["specimens"]:
            if sp["declared"] != sp["measured"]:
                kind = ("読めるはずが読めない" if sp["declared"]
                        else "読まないはずが読んだ")
                out.append({"key": n["key"], "task": sp["task"], "kind": kind,
                             "declared": sp["declared"], "measured": sp["measured"]})
            elif sp["says"] and sp["says"] not in sp["said"]:
                # ★ 読めてはいるが**別の場所**を指した（行が 1 つずれる退行が在りうる）。
                out.append({"key": n["key"], "task": sp["task"],
                             "kind": "指した場所が違う", "declared": sp["says"],
                             "measured": sp["said"]})
        # ★ status は検体の実測から**導き直せる**。名簿の手書きと突き合わせる。
        derived = _derive_status(n)
        if derived != n["status"]:
            out.append({"key": n["key"], "task": "", "kind": "status が実測と違う",
                         "declared": n["status"], "measured": derived})
    return out


#: 読めなかった時、それが「在庫」か「仕様」かを分けるのは points_at だけ。
UNREAD_STATUS = {"one_place": "gap",
                  "narrower_than_column": "narrower_than_column",
                  "whole_column": "by_design"}


def _derive_status(n: dict) -> str:
    """points_at（人の事実）と実測だけから status を導く。

    ★★ 初版は最後の分岐で名簿の status 自身を読んでいた ── **恒真**だった
      （名簿を写して名簿と比べていた）。人が決める事実を points_at の 1 つに畳み、
      status は完全に導出だけで出す形へ直した。
    ★ points_at は機械に導けない（「A1:C5 は列より狭い」は意味の判断）。そこだけ名簿が正。
    """
    if n["reads_any"]:
        return "reads"
    return UNREAD_STATUS[n["points_at"]]


def render(data: dict) -> str:
    lines = [f"記法の盤  {data['commit']}", ""]
    w = max(len(n["label"]) for n in data["notations"])
    for n in data["notations"]:
        lines.append(f"{n['label']:<{w}}  {STATUS_MARK.get(n['status'], n['status'])}")
        for sp in n["specimens"]:
            mark = "YES" if sp["measured"] else "no "
            warn = "" if sp["declared"] == sp["measured"] else "   ★ 宣言とちがう"
            lines.append(f"    {mark}  {sp['task']}{warn}")
            if sp["said"]:
                lines.append(f"         └ {sp['said']}")
        if n["unlock"]:
            lines.append(f"    unlock: {n['unlock']}")
        lines.append("")
    stock = [n for n in data["notations"]
             if n["status"] in ("gap", "narrower_than_column")]
    lines.append(f"★ 在庫（読めない記法）: {len(stock)} / {len(data['notations'])}")
    for n in stock:
        lines.append(f"    {n['label']}  ── {n['organ']}")
    bad = mismatches(data)
    lines.append("")
    lines.append(f"宣言と実測の食い違い: {len(bad)}")
    for b in bad:
        lines.append(f"    ★ {b['key']}  {b['kind']}  {b['task']}")
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="依頼文の記法の盤（宣言 × 実測）")
    ap.add_argument("--json", action="store_true", help="機械が読む形で出す")
    a = ap.parse_args(argv)
    data = survey()
    if a.json:
        print(json.dumps(data, ensure_ascii=False, indent=1))
    else:
        print(render(data))
    return 1 if mismatches(data) else 0


if __name__ == "__main__":
    raise SystemExit(main())
