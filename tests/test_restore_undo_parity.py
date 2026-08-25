# 復元の中#13（2026-08-24 の盲検）── restore と undo の非対称。
#
# ★ 実測: 同じ `restore_backup` を呼ぶのに、restore 側だけが劣化版だった。
#   - フォルダに対して `restore` すると「× w10 のバックアップが無い」と**的外れな理由**
#     （undo は「フォルダに対する undo はありません」と正しく言う）
#   - Excel ロックの関所を通らない／例外を言葉にしない／残り回数を言わない
#   ★ undo 側で直したものが restore に届かない ── 片配線。
#
# 契約:
#   ① フォルダには同じ理由で断る
#   ② 書けない時は同じ関所で止まる
#   ③ 復元できた時は同じ情報（残り回数）を出す
#   ④ 分岐を持たず委譲する（2 つ書けば必ずずれる）

import argparse
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
    monkeypatch.setattr(ailine, "MISCLASS_FILE", tmp_path / "misclass.jsonl")


def _prepared(tmp_path):
    book = tmp_path / "売上.xlsx"
    book.write_bytes(b"v0")
    ailine.make_backup(book)
    book.write_bytes(b"v1")
    return book


def test_delegates_instead_of_duplicating():
    """④ 分岐を持たない ── ここが崩れると、また片方だけ直る日が来る。"""
    import inspect
    src = inspect.getsource(ailine.cmd_restore)
    assert "cmd_undo(a)" in src, "委譲していない（2 つ書いている）"
    assert "restore_backup(" not in src, "restore が自前で復元している"


def test_folder_is_refused_with_the_same_reason(tmp_path, capsys):
    """① 的外れな理由を言わない。"""
    d = tmp_path / "folder"
    d.mkdir()
    rc = ailine.cmd_restore(argparse.Namespace(book=str(d), list=False))
    out = capsys.readouterr().out
    assert rc == 1
    assert "フォルダ" in out, f"的外れな理由を言った: {out}"


def test_excel_lock_stops_restore_too(tmp_path, capsys):
    """② undo で入れた関所が restore にも効く。"""
    book = _prepared(tmp_path)
    (tmp_path / f"~${book.name}").write_bytes(b"lock")
    before = book.read_bytes()
    rc = ailine.cmd_restore(argparse.Namespace(book=str(book), list=False))
    assert rc == ailine.EXIT_WRITE_BLOCKED, capsys.readouterr().out
    assert book.read_bytes() == before


def test_restore_reports_remaining_steps(tmp_path, capsys):
    """③ 残り回数を言う（undo と同じ情報）。"""
    book = _prepared(tmp_path)
    rc = ailine.cmd_restore(argparse.Namespace(book=str(book), list=False))
    out = capsys.readouterr().out
    assert rc == 0, out
    assert "戻せます" in out, f"残り回数が出ていない: {out}"
