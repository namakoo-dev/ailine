# -*- coding: utf-8 -*-
"""実際に打たれた依頼と、その時の冊の**列の顔ぶれ**を履歴から復元する（2026-09-19）。

★★ なぜ要るか（Namakoo「1でいこう」＝買い手の実依頼で揺れを測る）:
  翻訳は**実表の列名に接地**するので、冊が違えば別の測定になる。ところが:

  ・買い手 4 体の依頼文は**残っていない**（`AILINE_HOME` を切って走らせ、
    `AILINE_TRACE` を渡し損ねた ── `tests/blind_runs/004/README.md` が自白している）
  ・履歴に在る依頼の冊は**全部一時ファイル**で、1 つも現存しない（24/24 で消えていた）

  ★ 手の届く範囲でいちばん偏りの少ない標本は「**履歴に実際に打たれた依頼**のうち
    battery に無いもの」── そこから**種を固定して無作為抽出**する
    （揺れたかどうかで選ばない ── B 群でやった選び方の偏りを繰り返さない）。

★★ 冊は解釈行から復元する。解釈行は「その時どの列に解決したか」を持っている:

    解釈: 操作:1セル書換 対象の行:ワッシャー 対象列:数量 書き込む値:999
    → 列に『数量』が在った

★ **新しく作った列は入れない**（`新しい列の名前` ── 当時は存在しなかった）。
★ 復元できた列数を必ず併記する ── 少ない冊は接地が痩せていて、測定の意味が変わる。
  「復元した」と「当時の冊と同じ」は別物なので、そこを数字で見せる。

使い方:
    python bench/rebuild_books_from_history.py --n 24 --out tasks.json
"""
from __future__ import annotations

import argparse
import json
import pathlib
import random
import re

ROOT = pathlib.Path(__file__).resolve().parent.parent
HIST = pathlib.Path.home() / ".ailine" / "history.jsonl"

#: 解釈行のうち「**既に在った列**」を指すキー。★ 新しく作る列のキーは入れない。
COL_KEYS = ("対象列", "分類列", "集計列", "条件列", "キー列", "合計する列",
             "並べ替える列", "削除する列", "移動する列", "対象の列")
#: 「演算対象:在庫数 と 単価」のように複数を並べる形。
MULTI_KEYS = ("演算対象",)

_re_pair = re.compile(r"([^\s:：]+)[:：]([^\s]+)")
_re_sheet = re.compile(r"[＊＋][^:：]*[:：]\s*([^\s\[\]']+)")


def _cols_from(command: str) -> list:
    out = []
    for k, v in _re_pair.findall(str(command or "")):
        if k in COL_KEYS:
            out.append(v)
        elif k in MULTI_KEYS:
            out += [x for x in re.split(r"\s*と\s*", v) if x]
    return out


def _sheet_from(changes) -> str | None:
    for c in (changes or []):
        m = _re_sheet.match(str(c))
        if m and not m.group(1).startswith("["):
            return m.group(1)
    return None


def load_rows() -> list:
    rows = []
    for ln in HIST.read_bytes().decode("utf-8", "replace").splitlines():
        try:
            rows.append(json.loads(ln))
        except Exception:
            pass
    return rows


def battery_texts() -> set:
    b = json.loads((ROOT / "bench" / "translation_battery.json")
                    .read_bytes().decode("utf-8"))
    out = set()
    for k, v in b.items():
        if k.startswith("items") and isinstance(v, list):
            for it in v:
                if isinstance(it, dict) and it.get("text"):
                    out.add(str(it["text"]).strip())
    return out


def build(n: int, since: str, seed: int) -> list:
    rows = load_rows()
    canon = battery_texts()
    by_task: dict = {}
    for r in rows:
        t = str(r.get("task") or "").strip()
        if not t or str(r.get("ts", "")) < since or t in canon:
            continue
        by_task.setdefault(t, []).append(r)
    # ★ 種を固定して無作為に採る（選び直して都合の良い標本を作らない）
    rnd = random.Random(seed)
    picks = rnd.sample(sorted(by_task), min(n, len(by_task)))
    out = []
    for t in picks:
        cols, sheet = [], None
        for r in by_task[t]:
            cols += _cols_from(r.get("command"))
            sheet = sheet or _sheet_from(r.get("changes"))
        cols = list(dict.fromkeys(c for c in cols if c and len(c) <= 20))
        out.append({"task": t, "sheet": sheet or "Sheet", "headers": cols,
                     "recovered_columns": len(cols), "runs_in_history": len(by_task[t])})
    # ★★ 列を 1 つも復元できなかった依頼に、**共通の代替の冊**を当てる。
    #   ★ 空の冊で測ると接地がゼロになり、翻訳が別物になる（測っているものが変わる）。
    #   ★ 代替は俺が発明しない ── **復元できた列の実物**を寄せ集めて作る。
    #   ★ 印（grounded）を残して、群を分けて読めるようにする（混ぜて平均しない）。
    fallback = list(dict.fromkeys(c for it in out for c in it["headers"]))
    for it in out:
        it["grounded"] = bool(it["headers"])
        if not it["headers"]:
            it["headers"] = fallback
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="履歴から依頼と冊を復元する")
    ap.add_argument("--n", type=int, default=24)
    ap.add_argument("--since", default="2026-09-01")
    ap.add_argument("--seed", type=int, default=20260919)
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    items = build(a.n, a.since, a.seed)
    thin = sum(1 for it in items if it["recovered_columns"] == 0)
    print(f"復元: {len(items)} 件（種 {a.seed}・{a.since} 以降・battery の外から無作為）")
    print(f"★ 列を 1 つも復元できなかった依頼: {thin} 件 ── 接地が無いので別扱いにすること\n")
    for it in items:
        print(f"  列{it['recovered_columns']:>2}  {it['task'][:44]:<44} {it['headers']}")
    pathlib.Path(a.out).write_bytes(
        json.dumps(items, ensure_ascii=False, indent=1).encode("utf-8"))
    print(f"\n書き出し: {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
