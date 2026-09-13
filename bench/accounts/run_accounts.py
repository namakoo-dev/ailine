# -*- coding: utf-8 -*-
"""検体に `ailine accounts` を掛けて、凍結済みの採点器で裁く（需要③・2026-09-13）。

    python bench/accounts/run_accounts.py               # 全冊を走らせて内訳を出す
    python bench/accounts/run_accounts.py --show 40     # 外れの明細も出す（答えが見える）
    python bench/accounts/run_accounts.py --keep <dir>  # 候補の冊を残して目で見る

★ 道具は **subprocess で本物の CLI** を呼ぶ（import して内側を突かない ── 買い手が触るのは
  CLI であって関数ではない）。採点は `score_accounts.py` がそのまま行い、この runner は
  `--json` を採点器の契約（rows / refused / 原本が変わった）へ**写すだけ**。
★★ `ans`（答え）は**無視する**（採点器の docstring が名指しした縛り）── 答えが道具に入る扉を
  開けない。渡す入力は「今回の仕訳」と「過去の仕訳」のパスだけ。
★ 既定の `--show 0` は**外れの明細を出さない** ── 実装した本人が答えを読むと、次の直しが
  「一般化」ではなく「当てはめ」になる（docs/開発手法.md・tuning-audit と同じ線）。
  明細を見るのは、規則を凍結した人の仕事。
★ 検体がまだ 1 冊も無い状態でも **import はできる**（採点器も runner も落ちない）。
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

import score_accounts   # noqa: E402 ── 採点器（ailine_core を import しない側）

#: 出力の冊の名前（★ 1 ケース 1 冊・答えの id でフォルダを分ける）。
OUT_NAME = "候補.xlsx"


def run_one(today: Path, past_paths: list, out: Path) -> dict:
    """`ailine accounts` を 1 ケースに掛けて `--json` を辞書で返す。

    ★ 落ちても中身を見せて上げる（黙って空の辞書を返すと、道具の不調が「漏れ」に化ける）。
    """
    # ★ wheel を install していない手元でも src を見せる（tests/conftest.py と同じ作法）。
    full_env = dict(os.environ)
    full_env["PYTHONPATH"] = str(REPO / "src") + os.pathsep + full_env.get("PYTHONPATH", "")
    out.parent.mkdir(parents=True, exist_ok=True)
    argv = [sys.executable, "-m", "ailine", "accounts", str(today)]
    argv += ["--past"] + [str(p) for p in past_paths]
    argv += ["--out", str(out), "--json"]
    r = subprocess.run(argv, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=600, cwd=str(REPO), env=full_env)
    payload = None
    for line in reversed((r.stdout or "").strip().splitlines()):
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
        "rows": {int(row): {"account": (data or {}).get("account"),
                            "grade": (data or {}).get("grade"),
                            "reason": (data or {}).get("reason") or ""}
                 for row, data in (payload.get("rows") or {}).items()},
        "refused": payload.get("refused"),
        "原本が変わった": bool(payload.get("原本が変わった")),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=HERE)
    ap.add_argument("--answer", default="答え_accounts.json")
    ap.add_argument("--books", default="accounts_books")
    ap.add_argument("--show", type=int, default=0,
                    help="外れの明細を何件出すか（既定 0 ── 実装者が答えを読まないため）")
    ap.add_argument("--keep", type=Path, default=None,
                    help="候補の冊を残す先（既定は一時フォルダに作って消す）")
    a = ap.parse_args()

    answer_path = a.corpus / a.answer
    if not answer_path.is_file():
        print(f"★ 検体がありません（{answer_path}）── 測っていません")
        return 1

    workroot = Path(a.keep) if a.keep else Path(tempfile.mkdtemp(prefix="accounts_bench_"))
    workroot.mkdir(parents=True, exist_ok=True)
    exits: dict = {}

    def accounts(today_path, past_paths, ans) -> dict:
        """★ `ans` は受け取るが**使わない**（採点器の契約どおり・答えは道具に入れない）。"""
        case_id = str((ans or {}).get("id") or Path(today_path).stem)
        out_dir = workroot / case_id
        if out_dir.exists():
            shutil.rmtree(out_dir, ignore_errors=True)
        payload = run_one(Path(today_path), [Path(p) for p in past_paths],
                          out_dir / OUT_NAME)
        exits[case_id] = payload.pop("_exit")
        return to_contract(payload)

    try:
        res = score_accounts.score(accounts, a.corpus, a.answer, a.books)
        score_accounts.report(res, show=a.show)
        print(f"\n終了コード（候補の冊: {workroot}）")
        for case_id in sorted(exits):
            print(f"  {case_id}  exit={exits[case_id]}")
    finally:
        if not a.keep:
            shutil.rmtree(workroot, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
