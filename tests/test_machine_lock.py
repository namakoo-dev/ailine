# 実機の走行を 1 本に縛る鍵の番人 ── 2026-09-12
#
# ★★ なぜ在るか（同じ日に 3 回・実測）: 実機（LibreOffice）は固定ポート 2002・単一プロファイルで
#   起こるので、2 本同時に走ると 1 本目が資源を握り、2 本目以降は掴めず**終了もせず**溜まる。
#
#       単独走行   soffice ちょうど 1 個のまま・84 passed（34 分）
#       2 本同時   49 個まで積み上がり F が 20 件 ── ★ 本物の赤は 1 件も無かった
#
#   ★ 原因は「テストごとの漏れ」ではなく**走行の同時実行だけ**。だから掃除ではなく鍵で直す。
#   ★ 「気をつける」では止まらなかった（同じ家系を 1 日 4 回）── 手順を機械にした。
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
SCRIPT = REPO / "scripts" / "machine_lock.py"


def _run(args, home: Path, timeout: int = 60):
    env = {**os.environ, "AILINE_HOME": str(home), "PYTHONPATH": str(REPO / "src")}
    return subprocess.run([sys.executable, str(SCRIPT), *args], capture_output=True,
                          text=True, encoding="utf-8", errors="replace", env=env,
                          cwd=str(REPO), timeout=timeout)


def test_one_run_goes_through_and_the_key_is_returned(tmp_path):
    r = _run(["--what", "試し", "--", sys.executable, "-c", "print('走った')"], tmp_path)
    assert r.returncode == 0 and "走った" in r.stdout, (r.returncode, r.stdout, r.stderr)
    assert not (tmp_path / "machine.lock").exists(), "★ 鍵を外していない（次の走行が永久に断られる）"


def test_a_second_run_is_refused_at_once_and_names_who_holds_it(tmp_path):
    """★★ 本番: 2 本目は**待たずに**断る。待つと「詰まっているのか動いているのか」が分からない。
    そして**誰が握っているかを名指しする** ── 「使用中」だけでは人は動けない。"""
    inner = [sys.executable, str(SCRIPT), "--what", "内側", "--",
             sys.executable, "-c", "print('二重で通ってしまった')"]
    started = time.time()
    r = _run(["--what", "外側", "--", *inner], tmp_path)
    assert time.time() - started < 50, "★ 断るのに待っている"
    assert r.returncode == 2, (r.returncode, r.stdout, r.stderr)
    assert "二重で通ってしまった" not in r.stdout, "★ 2 本目が通った"
    assert "すでに 1 本" in r.stderr and "PID" in r.stderr and "外側" in r.stderr, r.stderr


def test_a_key_left_by_a_dead_process_is_taken_over(tmp_path):
    """★ 握ったプロセスが死んでいたら次が取れる（＝落ちた走行が永久に塞がない）。"""
    (tmp_path / "machine.lock").write_text(
        json.dumps({"pid": 999999, "started": time.time(), "what": "死んだ走行"}),
        encoding="utf-8")
    r = _run(["--what", "次", "--", sys.executable, "-c", "print('取れた')"], tmp_path)
    assert r.returncode == 0 and "取れた" in r.stdout, (r.returncode, r.stdout, r.stderr)


def test_a_key_held_by_a_living_process_is_never_taken_over_by_age(tmp_path):
    """★★ 時間では奪わない ── 実機テストは 75 分かかることがある。時間で奪えば今日の事故が再発する。

    生きている PID（この試験のプロセス自身）が 3 時間前から握っている形にして、それでも断ること。
    """
    (tmp_path / "machine.lock").write_text(
        json.dumps({"pid": os.getpid(), "started": time.time() - 3 * 3600, "what": "長い走行"}),
        encoding="utf-8")
    r = _run(["--what", "割り込み", "--", sys.executable, "-c", "print('奪ってしまった')"], tmp_path)
    assert r.returncode == 2, (r.returncode, r.stdout, r.stderr)
    assert "奪ってしまった" not in r.stdout
    assert "180 分経過" in r.stderr or "分経過" in r.stderr, r.stderr


def test_status_says_whether_a_run_is_active(tmp_path):
    assert _run(["--status"], tmp_path).returncode == 0
    (tmp_path / "machine.lock").write_text(
        json.dumps({"pid": os.getpid(), "started": time.time(), "what": "走行中"}), encoding="utf-8")
    r = _run(["--status"], tmp_path)
    assert r.returncode == 1 and "走行" in r.stdout, (r.returncode, r.stdout)


def test_the_prepush_script_runs_the_machine_tests_through_the_key():
    """★★ 「在っても鳴らない」を防ぐ ── 鍵が pre-push に**配線されている**ことを縛る。

    ★ 実測で 3 回踏んだのは pre-push が鍵を取らなかったからで、鍵を書くだけでは同じことが起きる。
    """
    text = (REPO / "scripts" / "pre-push.sh").read_text(encoding="utf-8")
    # ★ 物差しの失敗を 1 件: 最初は本文全体の出現位置で測り、冒頭のコメントの「-m local」を
    #   拾って赤くなった。**起動している行だけ**を見る（説明文は起動ではない）。
    runs = [ln for ln in text.splitlines()
            if not ln.lstrip().startswith("#") and "pytest" in ln and "-m local" in ln]
    assert runs, "★ pre-push が実機テストを走らせていない"
    for ln in runs:
        assert "machine_lock.py" in ln, f"★ 鍵を通さずに実機テストを起動している: {ln.strip()}"
    assert "rc -eq 2" in text, "★ 鍵が取れなかったときに push を止めていない"


@pytest.mark.parametrize("bad", [[], ["--what", "x"]])
def test_it_refuses_to_run_nothing(tmp_path, bad):
    """★ 陰性対照 ── 走らせるものが無いのに鍵を取って帰らない（鍵が残る）。"""
    r = _run(bad, tmp_path)
    assert r.returncode != 0
    assert not (tmp_path / "machine.lock").exists()
