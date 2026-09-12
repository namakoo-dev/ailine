# -*- coding: utf-8 -*-
"""実機の走行を 1 本に縛る（2026-09-12）。

    python scripts/machine_lock.py -- <走らせるもの…>

★★ なぜ在るか（同じ日に 3 回・実測）:
  実機テスト（`pytest -m local`）と `ailine run` の基盤は LibreOffice を
  **固定ポート 2002・単一プロファイル**で起こす。2 本同時に走ると、1 本目がポートを握り、
  2 本目以降の LO は掴めず**終了もせず**溜まる。実測:

      単独走行        soffice は **ちょうど 1 個**のまま（20 秒ごと 6 回）・84 passed（34 分）
      2 本同時        **49 個**まで積み上がり、F が 17 → 20。本物の赤は 1 件も無かった

  ★ 原因は「テストごとの漏れ」ではなく**走行の同時実行だけ**。だから掃除ではなく**鍵**で直す。
  ★ 「気をつける」では止まらなかった（同じ家系を 1 日 4 回踏んだ）── **手順を機械にする**。

★ 断り方は**即座**（待たない）。待つと「詰まっているのか動いているのか」が人に分からない。
★ stale は**時間で奪わない** ── 実機テストは 75 分かかることがある。時間で奪えば今日の事故が
  そのまま再発する。奪うのは「握った PID が生きていない」ときだけ。
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path

#: 鍵の置き場。★ `ailine` 本体の `~/.ailine/run.lock`（製品の直列化）とは**別の鍵**。
#:   あちらは 1 回の `ailine run` を守る。こちらは**走行そのもの**（テスト一式・押し）を守る。
LOCK_NAME = "machine.lock"


def lock_path() -> Path:
    home = Path(os.environ.get("AILINE_HOME") or (Path.home() / ".ailine"))
    home.mkdir(parents=True, exist_ok=True)
    return home / LOCK_NAME


def _alive(pid: int) -> bool:
    """その PID が生きているか。★ 生きていない鍵だけを奪う（時間では奪わない）。"""
    if pid <= 0:
        return False
    if os.name == "nt":
        out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"],
                             capture_output=True, text=True, errors="replace")
        return str(pid) in (out.stdout or "")
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def read_holder(path: Path) -> dict | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def describe(holder: dict | None) -> str:
    """人に見せる 1 行 ── **誰が握っているかを名指しする**（「使用中」だけでは動けない）。"""
    if not holder:
        return "（鍵の中身が読めません）"
    started = holder.get("started")
    when = (time.strftime("%H:%M:%S", time.localtime(started))
            if isinstance(started, (int, float)) else "?")
    mins = (time.time() - started) / 60 if isinstance(started, (int, float)) else -1
    return (f"PID {holder.get('pid')}（{when} に開始・{mins:.0f} 分経過"
            f"・{holder.get('what', '?')}）")


def acquire(path: Path, what: str) -> tuple:
    """鍵を取る。戻り値 (取れたか, 説明)。★ O_EXCL ── 「見てから作る」にしない。"""
    payload = json.dumps({"pid": os.getpid(), "started": time.time(), "what": what},
                         ensure_ascii=False)
    for _ in range(2):
        try:
            fd = os.open(str(path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            holder = read_holder(path)
            pid = int((holder or {}).get("pid") or 0)
            if _alive(pid):
                return False, describe(holder)
            # ★ 握ったプロセスが死んでいる ── そのときだけ奪う
            try:
                path.unlink()
            except OSError:
                return False, describe(holder)
            continue
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(payload)
        return True, describe(json.loads(payload))
    return False, "（鍵を取れませんでした）"


def release(path: Path) -> None:
    holder = read_holder(path)
    if holder and int(holder.get("pid") or 0) != os.getpid():
        return                      # ★ 自分のものでない鍵は外さない
    try:
        path.unlink()
    except OSError:
        pass


def main(argv: list) -> int:
    ap = argparse.ArgumentParser(description="実機の走行を 1 本に縛る")
    ap.add_argument("--what", default="", help="何を走らせるか（断りの文面に出る）")
    ap.add_argument("--status", action="store_true", help="いま誰が握っているかだけ見る")
    ap.add_argument("cmd", nargs=argparse.REMAINDER)
    a = ap.parse_args(argv)
    path = lock_path()

    if a.status:
        holder = read_holder(path)
        if holder and _alive(int(holder.get("pid") or 0)):
            print(f"実機の走行が 1 本あります: {describe(holder)}")
            return 1
        print("実機の走行はありません")
        return 0

    cmd = [x for x in a.cmd if x != "--"]
    if not cmd:
        ap.error("走らせるものを `--` の後に書いてください")
    what = a.what or " ".join(cmd[:4])
    got, who = acquire(path, what)
    if not got:
        print("✗ 実機の走行が**すでに 1 本**あります: " + who, file=sys.stderr)
        print("  実機（LibreOffice・ollama）は固定ポートで 1 本しか動けません ── "
              "2 本同時に走らせると、片方が資源を握ってもう片方が大量に落ちます。",
              file=sys.stderr)
        print("  ★ 終わるのを待ってから、もう一度。いま走っているものを止めたい場合は "
              "その PID を落としてください（名前一括で落とすと他の道具まで撃ちます）。",
              file=sys.stderr)
        return 2
    try:
        return subprocess.run(cmd).returncode
    finally:
        release(path)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
