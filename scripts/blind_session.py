# -*- coding: utf-8 -*-
"""盲検の 1 体を残して合格率で採る ── **薄い入口**（中身は tests/blind_session_core.py）。

★ なぜ中身が tests/ に在るか: 素の環境の番人は `scripts/` 同士の import を弾く（3 度踏んだ）。

使い方:
    python scripts/blind_session.py prepare 5体目
    （買い手が走る）
    python scripts/blind_session.py freeze  5体目
    python scripts/blind_session.py replay  5体目            # 既定 30 回
    python scripts/blind_session.py report  5体目 --to-shelf # 棚に 1 節を積む
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

import blind_session_core as core  # noqa: E402

#: ★ ブルーストロベリーの棚（Namakoo「ブルーストロベリーに記録を蓄積したい」）。
#:   ★ 棚は memory と別立て ── 索引は棚の MEMORY.md、実体は nodes/ に置く既存の作法。
SHELF = Path.home() / ".claude" / "projects" / "C--Windows-system32" / "library-blind"


def _load(name: str) -> tuple:
    d = core.CORPUS / name
    data = json.loads((d / "requests.json").read_bytes().decode("utf-8"))
    reqs = [r for r in data["requests"] if r.get("book") and r.get("task")]
    return d, data, reqs


def cmd_prepare(a) -> int:
    for line in core.prepare(a.name, a.force):
        print(line)
    return 0


def cmd_freeze(a) -> int:
    for line in core.freeze(a.name):
        print(line)
    return 0


def _replay(a) -> tuple:
    d, data, reqs = _load(a.name)
    # ★ 同じ文・同じ冊は 1 つにまとめる（同じものを 30 回 × 重複ぶん振らない）
    seen, uniq = set(), []
    for r in reqs:
        key = (r["task"], r["book"])
        if key not in seen:
            seen.add(key)
            uniq.append(r)
    print(f"再生: {len(uniq)} 件（重複を畳む前 {len(reqs)}）× {a.runs} 回 "
          f"= {len(uniq) * a.runs} 走行", flush=True)
    results = []
    for i, req in enumerate(uniq, 1):
        res = core.replay_one(d, req, a.runs)
        results.append(res)
        print(f"  [{i}/{len(uniq)}] {core.rate_line(res['reached'], res['runs'])} "
              f"「{req['task'][:36]}」", flush=True)
    (d / "replay.json").write_bytes(
        json.dumps(results, ensure_ascii=False, indent=1).encode("utf-8"))
    return d, data, results


def cmd_replay(a) -> int:
    d, data, results = _replay(a)
    v = core.verdicts(results)
    print("")
    print(core.render(a.name, results, v, data))
    # ★ 落とすのは「嘘の到達の影」が在る回だけ ── 到達の率は**記録**であって閾値ではない
    #   （導通率を合格条件にしなかったのと同じ線。閾値は人が読む側が持つ）。
    return 1 if v["failed_by_two_answers"] else 0


def cmd_report(a) -> int:
    d, data, _ = _load(a.name)
    results = json.loads((d / "replay.json").read_bytes().decode("utf-8"))
    v = core.verdicts(results)
    node = core.shelf_node(a.name, results, v, data)
    if not a.to_shelf:
        print(node)
        print("")
        print(f"★ 棚に積むなら --to-shelf（置き先: {SHELF}）")
        return 0
    (SHELF / "nodes").mkdir(parents=True, exist_ok=True)
    p = SHELF / "nodes" / f"{a.name}.md"
    p.write_bytes(node.encode("utf-8"))
    index = SHELF / "MEMORY.md"
    line = (f"- [[{a.name}]] — 到達 {core.rate_line(v['reached'], v['runs'])}"
            + ("・★ 嘘の到達の影あり" if v["failed_by_two_answers"] else ""))
    if not index.exists():
        index.write_bytes(("# library-blind — 盲検の記録\n\n"
                           "盲検 1 体につき 1 節。数と出来事だけを置き、生の依頼文と冊は\n"
                           "`C:\\Dev\\ailine\\bench\\blind\\` に残す（索引と実体を分ける）。\n\n"
                           "## 索引（1 体 1 行）\n\n").encode("utf-8"))
    body = index.read_bytes().decode("utf-8")
    if f"[[{a.name}]]" not in body:
        index.write_bytes((body.rstrip() + "\n" + line + "\n").encode("utf-8"))
    print(f"✓ 棚に積みました: {p}")
    print(f"  索引: {index}")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="盲検の 1 体を残して合格率で採る")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("prepare", help="買い手専用の AILINE_HOME を用意する")
    p.add_argument("name")
    p.add_argument("--force", action="store_true")
    p.set_defaults(fn=cmd_prepare)
    p = sub.add_parser("freeze", help="セッション後に依頼文と冊を固める")
    p.add_argument("name")
    p.set_defaults(fn=cmd_freeze)
    p = sub.add_parser("replay", help="固めた依頼を N 回ずつ当て直して合格率を出す")
    p.add_argument("name")
    p.add_argument("--runs", type=int, default=core.DEFAULT_RUNS)
    p.set_defaults(fn=cmd_replay)
    p = sub.add_parser("report", help="棚に積む 1 節を書き出す")
    p.add_argument("name")
    p.add_argument("--to-shelf", action="store_true")
    p.set_defaults(fn=cmd_report)
    a = ap.parse_args(argv)
    return a.fn(a)


if __name__ == "__main__":
    raise SystemExit(main())
