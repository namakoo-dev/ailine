# 復元の盲検 2 回目・致命②③（2026-08-25）── 命綱の穴。どちらも「今朝入れた物の隣」。
#
# 致命②: 壊れたバックアップを無検査で原本に被せて「✓ 復元した」
#   `restore_backup` は copy2 するだけで、復元元が開ける xlsx かを見なかった。
#   ★ この検査は今朝、**反映側にだけ**入れた（_why_output_is_unusable）。片配線の 4 度目。
#
# 致命③: `--copy` の成果物が、次の原本反映 run に黙って上書き・削除される
#   ★ **今朝入れた関所そのものの穴**。「この道具が過去にそこへ書いたか」で判定するので、
#     利用者がその後どれだけ手を入れても素通りした。
#   ★ 判定には三項が要る（依頼/宣言/実体）── 実体の項＝いま在る物の指紋。
#
# 契約:
#   ① 開けないバックアップは原本に被せない（原本は 1 バイトも変わらない）
#   ② 黙って別の世代へずらさない（どれを使うかは人が決める）
#   ③ 健全なバックアップは従来どおり復元できる（誤爆しない）
#   ④ 人が手を入れた .out は消さない
#   ⑤ 手つかずの .out は従来どおり作り直す（誤爆しない）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402


def _xlsx(p: Path, value="ok"):
    wb = openpyxl.Workbook(); wb.active["A1"] = value; wb.save(p)
    return p


# --- ①②③ 壊れたバックアップ ---------------------------------------------------------

def test_broken_backup_is_never_laid_over_the_original(tmp_path, monkeypatch):
    backups = tmp_path / "backups"
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = _xlsx(tmp_path / "book.xlsx", "current")
    ns = backups / ailine._backup_namespace(book)
    ns.mkdir(parents=True)
    (ns / "book.20260101T000000Z.xlsx").write_bytes(b"this is not a zip at all")
    _xlsx(ns / "book.20250101T000000Z.xlsx", "older-and-fine")   # ② ずらし先の誘惑
    before = book.read_bytes()
    with pytest.raises(ailine.BrokenBackupError) as ei:
        ailine.restore_backup(book)
    assert book.read_bytes() == before, "壊れたものを原本に被せた"
    assert "20260101" in str(ei.value), f"どの世代が壊れているか名指ししていない: {ei.value}"
    # ② 黙って古い方へずらしていないこと（ずらしていたら原本が変わっている）
    assert openpyxl.load_workbook(book).active["A1"].value == "current"


def test_healthy_backup_still_restores(tmp_path, monkeypatch):
    """③ 誤爆しない。"""
    backups = tmp_path / "backups"
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = _xlsx(tmp_path / "book.xlsx", "current")
    ns = backups / ailine._backup_namespace(book)
    ns.mkdir(parents=True)
    _xlsx(ns / "book.20260101T000000Z.xlsx", "previous")
    ailine.restore_backup(book)
    assert openpyxl.load_workbook(book).active["A1"].value == "previous"


def test_unprobeable_format_is_not_called_broken(tmp_path, monkeypatch):
    """★ 調べられないことを「壊れている」と言わない（命綱を丸ごと塞がない）。"""
    backups = tmp_path / "backups"
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = tmp_path / "book.ods"
    book.write_bytes(b"current-ods")
    ns = backups / ailine._backup_namespace(book)
    ns.mkdir(parents=True)
    (ns / "book.20260101T000000Z.ods").write_bytes(b"older-ods")
    ailine.restore_backup(book)
    assert book.read_bytes() == b"older-ods"


# --- ④⑤ 人が育てた .out を消さない -----------------------------------------------------

def _history_with_out(tmp_path, monkeypatch, out: Path, sha):
    import json
    hist = tmp_path / "history.jsonl"
    entry = {"ts": "2026-08-25T00:00:00+00:00", "book": str(tmp_path / "book.xlsx"),
              "task": "t", "ok": True, "out": str(out)}
    if sha is not None:
        entry["out_sha"] = sha
    hist.write_text(json.dumps(entry, ensure_ascii=False) + "\n", encoding="utf-8")
    monkeypatch.setattr(ailine, "HISTORY_FILE", hist)


def test_edited_copy_output_is_refused_not_overwritten(tmp_path, monkeypatch, capsys):
    """④ ★ README が慎重な人に勧める --copy の成果物が、原本反映 1 回で消えていた。"""
    book = _xlsx(tmp_path / "book.xlsx")
    out = _xlsx(ailine.out_book_path(book), "ailine-made")
    _history_with_out(tmp_path, monkeypatch, out, ailine._file_digest(out))
    _xlsx(out, "人が続きを書き足した")            # 人が手を入れた
    rc = ailine.refuse_if_output_is_someone_elses(book)
    assert rc == 7, "人が育てた成果物を黙って上書きしようとした"
    printed = capsys.readouterr().out
    assert "変更されています" in printed, printed


def test_untouched_copy_output_is_still_rebuilt(tmp_path, monkeypatch):
    """⑤ 誤爆しない: 手つかずなら従来どおり作り直す。"""
    book = _xlsx(tmp_path / "book.xlsx")
    out = _xlsx(ailine.out_book_path(book), "ailine-made")
    _history_with_out(tmp_path, monkeypatch, out, ailine._file_digest(out))
    assert ailine.refuse_if_output_is_someone_elses(book) is None


def test_old_history_without_a_fingerprint_behaves_as_before(tmp_path, monkeypatch):
    """指紋を残す前の記録では判定材料が無い ── 従来どおり（新しい嘘を足さない）。"""
    book = _xlsx(tmp_path / "book.xlsx")
    out = _xlsx(ailine.out_book_path(book), "ailine-made")
    _history_with_out(tmp_path, monkeypatch, out, None)
    _xlsx(out, "人が書き足した")
    assert ailine.refuse_if_output_is_someone_elses(book) is None


# --- 重大5: 名前を変えたら「無い」ではなく「別の名前で在る」と言う ----------------------

def test_renaming_points_at_the_generations_instead_of_saying_none(tmp_path, monkeypatch):
    """★ 実測: mv したら `× ... のバックアップが無い` だけ。世代はそのまま在るのに。"""
    backups = tmp_path / "backups"
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = _xlsx(tmp_path / "見積 書.xlsx")
    ns = backups / ailine._backup_namespace(book)
    ns.mkdir(parents=True)
    _xlsx(ns / "見積 書.20260101T000000Z.xlsx", "v0")
    renamed = _xlsx(tmp_path / "見積 書_最終版.xlsx")
    with pytest.raises(FileNotFoundError) as ei:
        ailine.restore_backup(renamed)
    msg = str(ei.value)
    assert "別の名前の世代が 1 件" in msg, f"在り処を言わず行き止まりにした: {msg}"
    assert str(ns) in msg, msg


def test_no_note_when_the_shelf_is_genuinely_empty(tmp_path, monkeypatch):
    """誤爆防止: 本当に 1 件も無い時は余計なことを言わない。"""
    backups = tmp_path / "backups"
    monkeypatch.setattr(ailine, "BACKUP_DIR", backups)
    book = _xlsx(tmp_path / "b.xlsx")
    with pytest.raises(FileNotFoundError) as ei:
        ailine.restore_backup(book)
    assert "別の名前の世代" not in str(ei.value), str(ei.value)
