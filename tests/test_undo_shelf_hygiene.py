# 復元の重大9（2026-08-24 の盲検）── 失敗した undo・no-op の undo でも棚に積む。
#
# ★ 実測: `restore_backup` は `make_backup(book, shelf=True)` を書き込みの**前**に
#   呼んでいた。読み取り専用で 3 回失敗させたら棚が 2→5 件に増えた。
#   さらに致命1 のループでは **棚 10 件すべてが同一内容の原本コピーで埋まり、
#   本物の run1/run2 の結果が押し出されて全滅**した。
#   ★ 棚は「undo をやり直す材料」── 何も起きなかった回に積むと、材料の方が消える。
#
# 契約:
#   ① 書き込みが失敗したら棚に積まない
#   ② 中身が変わらなかったら棚に積まない
#   ③ 本当に戻した時は、**戻す前の中身**が棚に残る（やり直せる）
#   ④ 棚の上限は従来どおり効く

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


@pytest.fixture(autouse=True)
def _isolated(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")


def _prepared(tmp_path):
    book = tmp_path / "b.xlsx"
    book.write_bytes(b"v0")
    ailine.make_backup(book)
    book.write_bytes(b"v1")
    return book


def test_shelf_keeps_what_we_replaced(tmp_path):
    """③ 戻す前の中身が残る ── やり直す材料。"""
    book = _prepared(tmp_path)
    ailine.restore_backup(book)
    assert book.read_bytes() == b"v0"
    shelved = [p.read_bytes() for p in ailine.list_undo_shelf(book)]
    assert b"v1" in shelved, f"戻す前の中身が棚に無い（やり直せない）: {shelved}"


def test_failed_restore_does_not_grow_the_shelf(tmp_path, monkeypatch):
    """① 失敗したら積まない。"""
    book = _prepared(tmp_path)
    before = len(ailine.list_undo_shelf(book))

    def boom(src, dst):
        raise PermissionError(13, "Permission denied")
    monkeypatch.setattr(ailine.shutil, "copy2", boom)
    with pytest.raises(PermissionError):
        ailine.restore_backup(book)
    assert len(ailine.list_undo_shelf(book)) == before, "失敗したのに棚が増えた"


def test_noop_restore_is_refused_before_it_can_pollute(tmp_path):
    """② no-op の undo は、そもそも起きない。

    ★ 検体の訂正: 初版は「no-op でも積まない」を測ろうとしたが、実際には
      **指し先の仕組み（致命1 の直し）が先に止める** ── 同じ中身の世代の上に
      立っていることが分かるので、`NoOlderBackupError` で拒否される。
      no-op が起きない以上、棚が汚れる道もここには無い。
      ★ 「起きないこと」を測るのだから、**拒否されること**を縛るのが正しい。
    """
    book = tmp_path / "b.xlsx"
    book.write_bytes(b"same")
    ailine.make_backup(book)          # 世代 = same（いまの中身と同じ）
    before = len(ailine.list_undo_shelf(book))
    with pytest.raises(ailine.NoOlderBackupError):
        ailine.restore_backup(book)
    assert len(ailine.list_undo_shelf(book)) == before, "拒否したのに棚が増えた"


def test_shelf_is_still_pruned(tmp_path):
    """④ 上限は効く。"""
    book = tmp_path / "b.xlsx"
    for i in range(15):
        book.write_bytes(f"v{i}".encode())
        ailine.make_backup(book)
    for _ in range(14):
        try:
            ailine.restore_backup(book)
        except ailine.NoOlderBackupError:
            break
    assert len(ailine.list_undo_shelf(book)) <= ailine.DEFAULT_KEEP_BACKUPS


def test_list_count_matches_what_you_can_actually_reach(tmp_path):
    """中12（2026-08-24 の盲検）── `undo --list` の世代数と到達段数の食い違い。

    ★ 実測では「3 世代」と表示され、原本も棚に実在するのに**到達できるのは 1 段**だった。
    ★ 測ったら、致命1（現在地を同一性で持つ）と「同じ中身の世代を積まない」の
      **副産物として既に解けていた**。推測で「直った」と言わず、回帰の番人として残す。
    """
    book = tmp_path / "b.xlsx"
    book.write_bytes(b"v0")
    ailine.make_backup(book); book.write_bytes(b"v1")
    ailine.make_backup(book); book.write_bytes(b"v2")
    ailine.restore_backup(book)                 # ここで「同じ中身」が生まれうる形にする
    ailine.make_backup(book); book.write_bytes(b"v3")

    listed = len(ailine.list_backups(book))
    said = ailine.undo_steps_left(book)
    reached = 0
    while True:
        try:
            ailine.restore_backup(book)
            reached += 1
        except ailine.NoOlderBackupError:
            break
    assert listed == reached, f"--list は {listed} 世代と言うのに {reached} 段しか行けない"
    assert said == reached, f"「あと {said} 回」と言うのに {reached} 回しか行けない"


def test_noop_undo_does_not_pollute_the_misclass_sensor(tmp_path, monkeypatch):
    """中14（2026-08-24 の盲検）── no-op の undo が misclass.jsonl を汚す。

    ★ 実測では、致命1 のループ 1 周ごとに「人が判断をひっくり返した容疑」が 1 件記録され、
      12 件すべてが同一 task の偽陽性だった ── 需要センサの信号が壊れる。
    ★ 測ったら、no-op の undo 自体が起きなくなったので既に消えていた。
      **推測で「直った」と言わず**、回帰の番人として残す。
    """
    import argparse
    monkeypatch.setattr(ailine, "MISCLASS_FILE", tmp_path / "m.jsonl")
    monkeypatch.setattr(ailine, "RUN_LOCK_FILE", tmp_path / "run.lock")
    book = tmp_path / "b.xlsx"
    book.write_bytes(b"same")
    ailine.make_backup(book)          # 世代 = same（いまの中身と同じ）
    for _ in range(3):
        ailine.cmd_undo(argparse.Namespace(book=str(book), list=False))
    n = (len(ailine.MISCLASS_FILE.read_text(encoding="utf-8").splitlines())
         if ailine.MISCLASS_FILE.exists() else 0)
    assert n == 0, f"何も起きていない undo が需要センサに {n} 件の偽陽性を残した"
