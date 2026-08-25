# 復元の致命3・致命4（2026-08-24 の盲検）── 安全網が無いのに「戻せる」と言う。
#
# ★ 致命3: `--keep-backups 0` は「作った直後に全部消す」＝安全網ゼロで原本を書き換える。
#   `prune_backups` は keep < 0 だけを無制限扱いにし、0 では `backups[0:]`＝
#   **いま作ったバックアップも含めて全部**削除する。argparse に下限検証は無い。
#   それでいて「（もとに戻す: ailine undo）」と表示する ── 嘘。
#   ★ `--help` は「負数で無制限」と書くが、**0 を無制限と読む利用者は普通に居る**
#     （多くの CLI がそう）。その 1 文字で原本が戻らなくなる。
#
# ★ 致命4: 剪定が無言で、undo の停止メッセージが嘘をつく。
#   上限超過で古い世代を黙って捨てたあと「最も古い状態です」と言う。
#   実際は「**まだ残っている中で**一番古い」でしかなく、原本は既に消してある。
#
# 契約:
#   ① keep=0 は受け付けない（意味が「全部消す」なので、事故にしかならない）
#   ② 負数は従来どおり無制限
#   ③ 剪定で世代を捨てたら、捨てたと言う
#   ④ 「これ以上は戻せません」は、原本まで残っている時だけ「最も古い状態」と言える

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402


@pytest.fixture(autouse=True)
def _isolated(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")


def test_keep_zero_is_refused_by_the_parser():
    """① 0 は「全部消す」── 受け付けない。"""
    parser = ailine.build_parser()
    with pytest.raises(SystemExit):
        parser.parse_args(["run", "b.xlsx", "何かして", "--keep-backups", "0"])


def test_keep_zero_message_says_what_to_use_instead(capsys):
    parser = ailine.build_parser()
    with pytest.raises(SystemExit):
        parser.parse_args(["run", "b.xlsx", "何かして", "--keep-backups", "0"])
    err = capsys.readouterr().err
    assert "0" in err
    assert "無制限" in err or "-1" in err, f"代わりに何を使えばいいか言っていない: {err}"


def test_negative_still_means_unlimited(tmp_path):
    """② 誤爆防止: 負数は従来どおり無制限（削除しない）。"""
    book = tmp_path / "b.xlsx"
    for i in range(3):
        book.write_bytes(f"v{i}".encode())
        ailine.make_backup(book, keep=-1)
    assert len(ailine.list_backups(book)) == 3


def test_pruning_is_disclosed(tmp_path):
    """③ 黙って捨てない。"""
    # ★ 治具の訂正: 告げるのは**捨てる前**なので、剪定後に呼んでも何も返らない
    #   （初版はそこを取り違えていた）。実際に人が見るのは make_backup の出力なので、
    #   そちらを捕まえる。
    import io, contextlib
    book = tmp_path / "b.xlsx"
    said = []
    for i in range(4):
        book.write_bytes(f"v{i}".encode())
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            ailine.make_backup(book, keep=2)
        if buf.getvalue().strip():
            said.append(buf.getvalue().strip())
    assert said, "剪定したのに何も言わない"
    assert any("捨てました" in t for t in said), said


def test_oldest_message_is_honest_about_pruning(tmp_path):
    """④ 原本まで残っていない時に「最も古い状態です」と言わない。"""
    book = tmp_path / "b.xlsx"
    for i in range(4):
        book.write_bytes(f"v{i}".encode())
        ailine.make_backup(book, keep=2)     # 原本 v0 は剪定で消えている
    while True:
        try:
            ailine.restore_backup(book)
        except ailine.NoOlderBackupError as e:
            msg = str(e)
            break
    assert "最も古い状態です" not in msg, f"原本は消してあるのに『最も古い』と言った: {msg}"
    assert "剪定" in msg or "捨て" in msg or "残っている中で" in msg, msg
