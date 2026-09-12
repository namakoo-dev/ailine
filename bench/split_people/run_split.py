# -*- coding: utf-8 -*-
"""検体 6 冊に `ailine split` を掛けて、凍結済みの採点器で裁く（2026-09-12）。

    python bench/split_people/run_split.py            # 6 冊を走らせて内訳を出す
    python bench/split_people/run_split.py --keep <dir>  # 配った冊を残して目で見る

★ 道具は **subprocess で本物の CLI** を呼ぶ（import して内側を突かない ── 買い手が触るのは
  CLI であって関数ではない）。採点は `score_split.py` がそのまま行い、この runner は
  `--json` を採点器の契約（parts/blank/lookalike/excluded/refused/proof）へ写すだけ。
★ `--by`（どの列で分けるか）は**人が打つ入力**なので、答えが宣言している担当者の列の
  見出しをそのまま使う。列が決められない冊（宣言が無い冊）は、人が打つであろう『担当』を
  渡す ── そこで断れるかどうかが、その冊の試験そのものだから。
  ★ 規則（分ける・疑う・証明の中身）は答えを 1 文字も見ていない。答えは採点にだけ使う。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
sys.path.insert(0, str(HERE))

import score_split   # noqa: E402 ── 採点器（ailine_core を import しない側）

#: 列が決められない冊で人が打つであろう見出し（答えの「理由に含む語」と同じ文字）。
FALLBACK_BY = "担当"
#: 金額の列の見出し（人が `--amount` に打つ文字）。
AMOUNT = "金額"


def run_one(book: Path, by: str, out_dir: Path) -> dict:
    """`ailine split` を 1 冊に掛けて `--json` を辞書で返す（落ちたら中身を見せて上げる）。"""
    # ★ wheel を install していない手元でも src を見せる（tests/conftest.py と同じ作法）。
    full_env = dict(os.environ)
    full_env["PYTHONPATH"] = str(REPO / "src") + os.pathsep + full_env.get("PYTHONPATH", "")
    r = subprocess.run(
        [sys.executable, "-m", "ailine", "split", str(book), "--by", by,
         "--amount", AMOUNT, "--out", str(out_dir), "--json"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=300, cwd=str(REPO), env=full_env)
    out = (r.stdout or "").strip().splitlines()
    payload = None
    for line in reversed(out):
        try:
            payload = json.loads(line)
            break
        except ValueError:
            continue
    assert payload is not None, (
        f"--json が読めない（exit={r.returncode}）:\n{r.stdout[-1500:]}\n{r.stderr[-800:]}")
    payload["_exit"] = r.returncode
    return payload


def to_contract(payload: dict) -> dict:
    """`--json` → 採点器の契約。★ 数字は 1 つも作り直さない（写すだけ）。"""
    return {
        "parts": {name: {"rows": part.get("rows") or [], "amount": part.get("amount")}
                  for name, part in (payload.get("parts") or {}).items()},
        "blank": payload.get("blank") or [],
        "lookalike": payload.get("lookalike") or [],
        "excluded": payload.get("excluded") or [],
        "refused": payload.get("refused"),
        "proof": {"rows": {"parts": ((payload.get("proof") or {}).get("rows") or {}).get("parts")},
                  "amount": {"parts": ((payload.get("proof") or {}).get("amount") or {}).get("parts")}},
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=HERE)
    ap.add_argument("--answer", default="答え_split.json")
    ap.add_argument("--books", default="split_books")
    ap.add_argument("--keep", type=Path, default=None,
                    help="配った冊を残す先（既定は一時フォルダに作って消す）")
    a = ap.parse_args()

    workroot = Path(a.keep) if a.keep else Path(tempfile.mkdtemp(prefix="split_bench_"))
    workroot.mkdir(parents=True, exist_ok=True)
    exits = {}

    def split(path: Path, ans: dict) -> dict:
        col = ans.get("担当者の列") or {}
        by = col.get("見出し") or FALLBACK_BY
        out_dir = workroot / ans["id"]
        if out_dir.exists():
            shutil.rmtree(out_dir, ignore_errors=True)
        payload = run_one(Path(path), by, out_dir)
        exits[ans["id"]] = (payload.pop("_exit"), by)
        return to_contract(payload)

    try:
        res = score_split.score(split, a.corpus, a.answer, a.books)
        score_split.report(res, show=40)
        print("\n終了コードと --by（配った先: " + str(workroot) + "）")
        for book_id in sorted(exits):
            code, by = exits[book_id]
            print(f"  {book_id}  exit={code}  --by {by}")
    finally:
        if not a.keep:
            shutil.rmtree(workroot, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
