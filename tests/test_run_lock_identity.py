# run.lock の持ち主判定（2026-08-24・**同じ日に 2 回作り直した**）。
#
# ★ 経緯を残す。これは「直したつもりが足りなかった」の実物の記録:
#
#   19:46  第 1 版 — PID の生死で判定していたのを、「誰のロックか（実行ファイル名）」を
#          焼く三項化にした。検体も陽性対照つきで書き、変異試験も通した。
#   22:44  盲検の使い手が **2 回**踏んだ ── 死んだ PID のロックで exit 6。直っていなかった。
#   23:30  真因: 記録した名前は `python.exe` で、それはこの機械で走る**あらゆる python**に
#          一致する。テストで python を大量に起動するので PID の使い回しが日常的に起き、
#          死んだ ailine のロックが「生きている」と誤判定されていた。
#          → 三項にしたつもりで、第三項が**粗すぎた**。
#
# ★ 根治: 生死を**推測しない**。OS の排他ロックに持たせる。
#   プロセスが死ねば OS が必ず解放するので、居座りが原理的に起きない。
#   ファイルの中身（pid/ts/image）は今も書くが、**人へ見せる説明専用**で判定には使わない。
#
# 契約:
#   ① 別プロセスが持っている間は取れない（そして待てと言う）
#   ② 持ち主が**死んだら**取れる（居座らない）── 推測でなく OS が保証する
#   ③ 判定に PID の生死を使っていない（構造で縛る）
#   ④ ロックの中身は説明として書かれている（人が「誰が持っているか」を読める）

import json
import os
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _hold_in_child(lock_path):
    parts = [
        "import sys",
        "sys.path.insert(0, r%r)" % str(REPO / "src"),
        "import ailine, time",
        "from pathlib import Path",
        "ok, _ = ailine.acquire_run_lock(Path(r%r))" % str(lock_path),
        "print(chr(79) + chr(75) if ok else chr(78) + chr(71), flush=True)",
        "time.sleep(120)",
    ]
    proc = subprocess.Popen([sys.executable, "-c", chr(10).join(parts)],
                             stdout=subprocess.PIPE, text=True, encoding="utf-8")
    assert proc.stdout.readline().strip() == "OK", "子がロックを取れていない"
    return proc


def test_held_by_a_live_process_blocks_and_says_what_to_do(tmp_path):
    lock = tmp_path / "run.lock"
    child = _hold_in_child(lock)
    try:
        ok, msg = ailine.acquire_run_lock(lock)
        assert ok is False
        assert "実行中" in msg
        assert "待って" in msg, f"次の一手が無い: {msg}"
    finally:
        child.kill(); child.wait(timeout=10)


def test_a_dead_holder_does_not_squat(tmp_path):
    """★ これが盲検で 2 回踏まれた事故そのもの。PID の使い回しでも居座らない。"""
    lock = tmp_path / "run.lock"
    child = _hold_in_child(lock)
    child.kill(); child.wait(timeout=10)
    time.sleep(0.5)
    ok, msg = ailine.acquire_run_lock(lock)
    assert ok is True, f"死んだ持ち主のロックが居座った: {msg}"
    ailine.release_run_lock(lock)


def test_decision_does_not_consult_pid_liveness():
    """③ 構造で縛る: acquire_run_lock が PID の生死を見ていないこと。

    ★ 実行時の検体（上の 2 本）だけだと、将来また「PID が生きているか」で
      判定する実装に戻っても気づけない ── 一度そこで間違えているので、形も縛る。
    """
    import inspect
    src = inspect.getsource(ailine.acquire_run_lock)
    assert "_pid_alive" not in src, "PID の生死で判定に戻っている（推測をやめたはず）"
    assert "_try_os_lock" in src, "OS の排他ロックを使っていない"


def test_lock_file_explains_who_holds_it(tmp_path):
    """④ 中身は**説明**として在る（判定には使わないが、人は読む）。"""
    lock = tmp_path / "run.lock"
    assert ailine.acquire_run_lock(lock)[0] is True
    try:
        # ★ 人が読む経路（画面に「誰が持っているか」を出す所）と同じ関数で読む。
        #   ファイル全体を読むと末尾の OS ロックを跨いで PermissionError になる ──
        #   実測で画面に pid=?・? と出ていた事故そのもの。
        info = ailine._read_lock_info(lock)
        assert info is not None, "説明が読めない（画面に pid=? と出る状態）"
        assert info["pid"] == os.getpid()
        assert info.get("image"), "誰のロックかが書かれていない"
    finally:
        ailine.release_run_lock(lock)


def test_a_different_lock_path_is_not_mine(tmp_path):
    """★ 実装の穴だった所: 「既に何か持っている」だけで通すと、**別の鍵**でも
       開いてしまう（検体を並べて走らせて発覚した）。"""
    a, b = tmp_path / "a.lock", tmp_path / "b.lock"
    assert ailine.acquire_run_lock(a)[0] is True
    try:
        ok, msg = ailine.acquire_run_lock(b)
        assert ok is False, "別の鍵のロックを自分の物として通した"
    finally:
        ailine.release_run_lock(a)
