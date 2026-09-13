# -*- coding: utf-8 -*-
"""検体「月末の束」を、本物の道具（`python -m ailine forms`）にかけて採点する（2026-09-13）。

    python bench/received_invoices/run_monthend.py

★ 道具は subprocess で呼ぶだけ ── この採点の経路は `ailine_core` を 1 行も import しない
  （score_monthend.py の独立性をここでも保つ）。答えは道具に渡さない。
★ 検体は生成物（gitignore）。先に `mk_monthend.py` を走らせること。
★ LibreOffice は起動しない（`ailine forms` 自身が読むだけの器官で LO を起動しない）。
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
sys.path.insert(0, str(HERE))

import score_monthend                                  # noqa: E402
from bundle_common import load_bundles                  # noqa: E402

ANSWER, BOOKS = "答え_monthend.json", "monthend_books"


def _run_cli(folder: Path, out: Path) -> dict:
    """`python -m ailine forms <folder> --out <out> --json` を呼んで結果の dict を返す。"""
    env = dict(os.environ)
    env["PYTHONPATH"] = str(REPO / "src") + os.pathsep + env.get("PYTHONPATH", "")
    r = subprocess.run(
        [sys.executable, "-m", "ailine", "forms", str(folder), "--out", str(out), "--json"],
        cwd=str(REPO), env=env, capture_output=True, text=True, encoding="utf-8", timeout=300)
    if r.returncode != 0:
        print(f"★ ailine forms が異常終了しました（{folder.name}）: exit={r.returncode}")
        print(r.stdout[-2000:])
        print(r.stderr[-2000:])
        raise SystemExit(1)
    try:
        return json.loads(r.stdout)
    except json.JSONDecodeError:
        print(f"★ --json の出力が JSON でない（{folder.name}）:")
        print(r.stdout[-2000:])
        raise


def tool_factory(workdir: Path):
    """束（フォルダ）→ suspect() 相当の出力。1 回呼ぶごとに実プロセスを 1 回起動する。"""
    def suspect(folder: Path) -> list:
        out = workdir / f"{folder.name}__out.xlsx"
        result = _run_cli(folder, out)
        return result.get("suspicions") or []
    return suspect


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=HERE)
    ap.add_argument("--show", type=int, default=60)
    a = ap.parse_args()

    if not (a.corpus / ANSWER).exists():
        print(f"★ 検体がありません（{ANSWER}）。先に mk_monthend.py を走らせてください。")
        return 2

    print("★ 採点器の自己診断（分布の床・対照・変異）")
    if score_monthend.self_test(a.corpus, ANSWER, BOOKS) != 0:
        print("★ 採点器が壊れている ── 実物の点数は見ない")
        return 1

    print("\n★ いまの製品（python -m ailine forms）を通した実測")
    with tempfile.TemporaryDirectory(prefix="ailine_monthend_") as d:
        tool = tool_factory(Path(d))
        res = score_monthend.score(tool, a.corpus, ANSWER, BOOKS)
        score_monthend.report(res, show=a.show)

    print("\n束ごとの出した所見（生データ）:")
    with tempfile.TemporaryDirectory(prefix="ailine_monthend_") as d:
        tool = tool_factory(Path(d))
        for b in load_bundles(a.corpus / ANSWER):
            found = tool(a.corpus / BOOKS / b["束"])
            print(f"\n=== {b['束']}  冊 {len(b.get('冊') or [])}  "
                  f"仕込み {len(b.get('疑い') or [])}  怪しくない {len(b.get('怪しくない') or [])} ===")
            for s in found:
                print(f"  [{s['種類']}] {s['冊']}\n      {s['理由']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
