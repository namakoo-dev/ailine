# 復元の致命5・重大7（2026-08-24 の盲検）── undo に書き込みの関所が無かった。
#
# ★ 致命5: run は Excel ロックで止まる（exit 5）のに、undo は素通りして exit 0 だった。
#   「Excel で結果を見て、気に入らないから戻す」は undo の**最も自然な使い方**。
#   そこだけ関所が無い ── しかも Excel が開いたままディスクを書き換えると、
#   次の Ctrl+S でメモリ像に上書きされて undo が無かったことになる（検分者の疑い S2）。
# ★ 重大7: 読み取り専用の原本に undo → **生の traceback**。命綱がスタックトレースで死ぬ。
#
# 契約:
#   ① Excel のロックファイルが在れば、書く前に止める
#   ② 書けない時は必ず言葉にする（traceback を出さない）
#   ③ 止めたら原本は 1 バイトも変わらない

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


def test_undo_stops_at_the_excel_lock(tmp_path, capsys):
    """①③ run と同じ関所を通る。"""
    import argparse
    book = _prepared(tmp_path)
    (tmp_path / f"~${book.name}").write_bytes(b"lock")
    before = book.read_bytes()
    rc = ailine.cmd_undo(argparse.Namespace(book=str(book), list=False))
    out = capsys.readouterr().out
    assert rc == ailine.EXIT_WRITE_BLOCKED, f"undo が素通りした: exit={rc} / {out}"
    assert book.read_bytes() == before, "止めたのに原本が変わった"
    assert "ロックファイル" in out and "Excel で開いています" in out, out


def test_undo_without_lock_still_restores(tmp_path, capsys):
    """誤爆防止: ロックが無ければ従来どおり戻る。"""
    import argparse
    book = _prepared(tmp_path)
    rc = ailine.cmd_undo(argparse.Namespace(book=str(book), list=False))
    assert rc == 0, capsys.readouterr().out
    assert book.read_bytes() == b"v0"


def test_permission_error_becomes_words_not_a_traceback(tmp_path, monkeypatch, capsys):
    """② 命綱がスタックトレースで死なない。"""
    import argparse
    book = _prepared(tmp_path)

    def boom(_b):
        raise PermissionError(13, "Permission denied")
    monkeypatch.setattr(ailine, "restore_backup", boom)
    rc = ailine.cmd_undo(argparse.Namespace(book=str(book), list=False))
    out = capsys.readouterr().out
    assert rc == ailine.EXIT_WRITE_BLOCKED
    assert "書き込めませんでした" in out and "読み取り専用" in out, out
