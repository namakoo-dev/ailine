# 復元の重大8（2026-08-24 の盲検）── undo/restore が実行ロックを取らない。
#
# ★ 実測: 生きた PID を焼いた run.lock を置いて
#     run  → exit=6  × 別の ailine が実行中です（pid=…）
#     undo → exit=0  ✓ … から復元した        ← 素通り
#   run 実行中の undo は、run 末尾の atomic_replace_inplace に上書きされるうえ、
#   その run が「復元したばかりの内容」を世代として積む
#   ── **致命1（undo の振動）の引き金を自分で引く**。
#
# 契約:
#   ① 別の ailine が走っている間、undo は exit 6 で止まる
#   ② 止めたら原本は 1 バイトも変わらない
#   ③ restore も同じ（委譲しているので自動的に効くはずだが、構造で縛る）
#   ④ `undo --list`（読むだけ）はロックを取らない ── 見るだけの操作を止めない
#   ⑤ run と undo が**同じ器官**を通る（書き写さない）

import argparse
import subprocess
import sys
import time
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


@pytest.fixture(autouse=True)
def _isolated(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")
    monkeypatch.setattr(ailine, "MISCLASS_FILE", tmp_path / "misclass.jsonl")
    monkeypatch.setattr(ailine, "RUN_LOCK_FILE", tmp_path / "run.lock")


def _prepared(tmp_path):
    book = tmp_path / "売上.xlsx"
    book.write_bytes(b"v0")
    ailine.make_backup(book)
    book.write_bytes(b"v1")
    return book


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
    assert proc.stdout.readline().strip() == "OK"
    return proc


def test_undo_waits_for_a_running_ailine(tmp_path, capsys):
    """①② 走っている間は止まる・原本は無傷。"""
    book = _prepared(tmp_path)
    before = book.read_bytes()
    child = _hold_in_child(ailine.RUN_LOCK_FILE)
    try:
        rc = ailine.cmd_undo(argparse.Namespace(book=str(book), list=False))
        out = capsys.readouterr().out
        assert rc == 6, f"undo が素通りした: exit={rc} / {out}"
        assert book.read_bytes() == before, "止めたのに原本が変わった"
        assert "実行中" in out, out
    finally:
        child.kill(); child.wait(timeout=10)


def test_restore_is_guarded_too(tmp_path, capsys):
    """③ restore も同じ（委譲しているので効く）。"""
    book = _prepared(tmp_path)
    child = _hold_in_child(ailine.RUN_LOCK_FILE)
    try:
        rc = ailine.cmd_restore(argparse.Namespace(book=str(book), list=False))
        assert rc == 6, capsys.readouterr().out
    finally:
        child.kill(); child.wait(timeout=10)


def test_listing_is_not_blocked(tmp_path, capsys):
    """④ 見るだけの操作は止めない。"""
    book = _prepared(tmp_path)
    child = _hold_in_child(ailine.RUN_LOCK_FILE)
    try:
        rc = ailine.cmd_undo(argparse.Namespace(book=str(book), list=True))
        assert rc == 0, capsys.readouterr().out
    finally:
        child.kill(); child.wait(timeout=10)


def test_run_and_undo_share_one_organ():
    """⑤ 書き写さない ── 同じ器官を通る。"""
    import inspect
    for fn in (ailine.cmd_run, ailine.cmd_undo):
        src = inspect.getsource(fn)
        assert "under_run_lock(" in src, f"{fn.__name__} が器官を通っていない"
        assert "acquire_run_lock()" not in src, f"{fn.__name__} が自前で取っている"
