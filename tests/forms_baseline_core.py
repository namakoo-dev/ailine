# -*- coding: utf-8 -*-
"""請求書の読みの**基準線** ── 触る前に凍らせ、触った後に差を見る（2026-09-20）。

★★ なぜ在るか（`docs/PENDING-20260918-請求元と宛先の取り違え.md` の「必ず守ること」）:

    ★ **触る前に 305 冊の基準線を取る。**2026-09-18 はそれが在ったから、225 冊と 16 冊の
    退行をどちらも commit 前に止められた。無ければ「買い手の冊は直った」で出荷していた。

  その基準線は scratchpad に在って**残っていない**。同じ物を repo の道具として作り直す
  ── 次に誰かが読みの規則へ触る時、最初に打つのがこれになるように。

★★ 測るのは**製品と同じ入口**（`form_read.read_book`）── シートの選び方まで含めて
  製品がやることをやる。別の読み方で測ると、**製品ではないもの**の基準線になる
  （この repo が何度も踏んだ「測定器が別物」）。

★ 凍らせるのは項目ごとの **(区分, 値)** だけ。根拠の番地まで入れると、関係ない直しで
  差分が埋まって読めなくなる（見たいのは「人に出る答えが変わったか」）。

★ なぜ tests/ に在るか: 素の環境の番人は `scripts/` 同士の import を弾く（3 度踏んだ）。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import openpyxl  # noqa: E402

from ailine_core import form_read  # noqa: E402
from ailine_core.field_record import grade, value  # noqa: E402

#: ★ 基準線を取る冊の束（305 冊）。★ 固定する ── 束が変われば数の意味が変わる。
CORPUS = REPO / "bench" / "received_invoices"
#: ★ 凍らせた姿の置き場（repo に置く ── scratchpad に置いて消えたのが 2026-09-18 の反省）。
FROZEN = REPO / "bench" / "received_invoices" / "基準線.json"

#: 見る項目。★ 宛先と請求元が本題だが、巻き添えは**他の項目にも出る**ので全部見る。
FIELDS = ("宛先", "請求元", "請求額", "請求日", "請求番号")


def read_one(path: Path) -> dict:
    """1 冊を**製品と同じ入口**で読み、項目 → (区分, 値) にする。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:                      # noqa: BLE001 ── 読めない冊も記録に残す
        return {"_読めず": f"{type(e).__name__}: {e}"}
    try:
        try:
            wbf = openpyxl.load_workbook(path, data_only=False)
        except Exception:                       # noqa: BLE001
            wbf = None
        recs = form_read.read_book(wb, wbf)
    finally:
        wb.close()
    out = {}
    for f in FIELDS:
        r = recs.get(f)
        if r is None:
            out[f] = {"区分": "無", "値": ""}
            continue
        out[f] = {"区分": grade(r), "値": str(value(r) if value(r) is not None else "")}
    return out


def survey(folder: Path | None = None) -> dict:
    """束ぜんぶを読む。戻り値は「束からの相対パス → 項目 → (区分, 値)」。"""
    root = folder or CORPUS
    out = {}
    for p in sorted(root.rglob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        out[str(p.relative_to(root)).replace("\\", "/")] = read_one(p)
    return out


def counts(got: dict) -> dict:
    """人が読む 1 行 ── 「宛先が出た n / 請求元が出た n / 両方 n / 読めず n」。"""
    def filled(rec, f):
        return bool(rec.get(f, {}).get("値"))
    a = sum(1 for r in got.values() if "_読めず" not in r and filled(r, "宛先"))
    b = sum(1 for r in got.values() if "_読めず" not in r and filled(r, "請求元"))
    both = sum(1 for r in got.values()
               if "_読めず" not in r and filled(r, "宛先") and filled(r, "請求元"))
    bad = sum(1 for r in got.values() if "_読めず" in r)
    return {"冊": len(got), "宛先が出た": a, "請求元が出た": b, "両方": both, "読めず": bad}


def freeze(got: dict, path: Path | None = None) -> Path:
    p = path or FROZEN
    p.write_bytes((json.dumps(got, ensure_ascii=False, indent=1) + "\n").encode("utf-8"))
    return p


def load(path: Path | None = None) -> dict:
    p = path or FROZEN
    return json.loads(p.read_bytes().decode("utf-8")) if p.exists() else {}


def diff(before: dict, after: dict) -> list:
    """動いた冊だけを返す。★ 項目ごとに (前, 後) を並べる ── 数だけでは読めない。"""
    out = []
    for name in sorted(set(before) | set(after)):
        b, a = before.get(name, {}), after.get(name, {})
        if b == a:
            continue
        moved = {}
        for f in set(b) | set(a):
            if b.get(f) != a.get(f):
                moved[f] = (b.get(f), a.get(f))
        out.append({"冊": name, "動いた項目": moved})
    return out


def render(before: dict, after: dict, limit: int = 12) -> str:
    changed = diff(before, after)
    lines = [f"基準線: {counts(before)}", f"いま　: {counts(after)}", ""]
    lines.append(f"動いた冊: {len(changed)}")
    for row in changed[:limit]:
        lines.append(f"  {row['冊']}")
        for f, (b, a) in sorted(row["動いた項目"].items()):
            lines.append(f"    {f}: {b} → {a}")
    if len(changed) > limit:
        lines.append(f"  …ほか {len(changed) - limit} 冊")
    return "\n".join(lines)
