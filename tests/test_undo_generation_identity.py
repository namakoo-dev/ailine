# 復元の致命1（2026-08-24 の盲検・最重）── undo が世代の中で振動し、原本へ永久に到達できない。
#
# ★ 実測（検分者が end-to-end で再現）:
#   トリガは `run → undo → run` ── **README が勧めている使い方そのもの**。
#   undo 直後に run すると、make_backup が「いま復元したばかりの内容」をもう一度世代に積む。
#   `_undo_position` は現在地を**バイト一致の最初のヒット**で決めるので、同内容の世代が
#   2 つ並んだ瞬間に位置が確定できなくなり、歩みが壊れる:
#       undo → 原本   undo → v1（時間を**前に**進んだ）   undo → 原本   undo → v1 …（無限）
#   `undo --list` は「3 世代」と出し、原本もバックアップ棚に**実在する**のに到達できない。
#   端の検出（NoOlderBackupError）は i+1 >= len(backups) なので、i が 0 に張り付くと
#   **永遠に発火しない**。exit は毎回 0（＝嘘の成功）。
#
# ★ 根: 判定に要る三項のうち「いま自分はどの世代の上に立っているか（**同一性**）」を
#   持たず、「内容の等値」で代用している。**内容は世代の一意キーではない。**
#
# 契約:
#   ① 同じ内容の世代が 2 つ並んでも、現在地が一意に決まる
#   ② run → undo → run → undo… で、必ず 1 段ずつ古い方へ進む（振動しない）
#   ③ 端に着いたら止まる（no-op で「復元した」と言わない）
#   ④ 人が Excel で直接編集した後は「新しい編集の直後」扱いに戻る（指し先が実体と違う）
#   ⑤ 指し先を持たない古いバックアップ置き場でも壊れない（後方互換）

import shutil
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402


@pytest.fixture(autouse=True)
def _isolated_home(tmp_path, monkeypatch):
    monkeypatch.setattr(ailine, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(ailine, "HISTORY_FILE", tmp_path / "history.jsonl")


def _write(p: Path, text: str) -> None:
    p.write_bytes(text.encode("utf-8"))


def _edit(book: Path, text: str) -> None:
    """実編集を模す: バックアップを取ってから中身を替える（run と同じ順序）。"""
    ailine.make_backup(book)
    _write(book, text)


def test_duplicate_generations_do_not_break_the_walk(tmp_path):
    """① 同内容の世代が並んでも現在地が一意に決まる（これが事故の芯）。"""
    book = tmp_path / "b.xlsx"
    _write(book, "v0")
    _edit(book, "v1")
    _edit(book, "v2")
    ailine.restore_backup(book)                 # v2 → v1
    assert book.read_bytes() == b"v1"
    _edit(book, "v3")                           # ★ ここで「v1」が世代に二重に積まれる
    ailine.restore_backup(book)                 # v3 → v1
    assert book.read_bytes() == b"v1"
    ailine.restore_backup(book)                 # v1 → **v0 へ進む**（振動しない）
    assert book.read_bytes() == b"v0", "同内容の世代で歩みが壊れた（振動）"


def test_walk_reaches_the_original_and_then_stops(tmp_path):
    """②③ 1 段ずつ進み、端で止まる。"""
    book = tmp_path / "b.xlsx"
    _write(book, "v0")
    _edit(book, "v1")
    ailine.restore_backup(book)
    assert book.read_bytes() == b"v0"
    with pytest.raises(ailine.NoOlderBackupError):
        ailine.restore_backup(book)


def test_steps_left_is_honest_with_duplicates(tmp_path):
    """★ 実測: `undo --list` が 3 世代と言うのに到達できるのは 1 段だった。"""
    book = tmp_path / "b.xlsx"
    _write(book, "v0")
    _edit(book, "v1")
    _edit(book, "v2")
    ailine.restore_backup(book)
    _edit(book, "v3")
    left = ailine.undo_steps_left(book)
    seen = []
    for _ in range(left):
        ailine.restore_backup(book)
        seen.append(book.read_bytes())
    assert seen[-1] == b"v0", f"あと {left} 回と言ったのに原本へ着かない: {seen}"
    with pytest.raises(ailine.NoOlderBackupError):
        ailine.restore_backup(book)


def test_manual_edit_resets_to_fresh(tmp_path):
    """④ 人が Excel で直接触ったら、指し先は当てにならない ── 新しい編集の直後に戻す。"""
    book = tmp_path / "b.xlsx"
    _write(book, "v0")
    _edit(book, "v1")
    ailine.restore_backup(book)                 # 指し先 = v0 の世代
    _write(book, "人が直接編集")                # ★ ailine を通さない変更
    ailine.restore_backup(book)                 # 最新世代（v0 を含む列の先頭）へ戻るはず
    assert book.read_bytes() == b"v0"


def test_works_without_a_pointer_file(tmp_path):
    """⑤ 後方互換: 指し先の記録が無い既存のバックアップ置き場でも動く。"""
    book = tmp_path / "b.xlsx"
    _write(book, "v0")
    _edit(book, "v1")
    for p in (ailine.BACKUP_DIR).rglob("*.at"):
        p.unlink()
    ailine.restore_backup(book)
    assert book.read_bytes() == b"v0"


def test_identity_is_what_locates_us_not_content(tmp_path):
    """★ 変異試験で発覚した穴の検体。

    「同じ中身の世代を積まない」だけでも上の検体は全部通ってしまい、**同一性の判定が
    一度も試されていなかった**（指し先を無視する変異を入れても全部緑だった）。

    内容が本当に重複する形 ── 行き来して同じ値に戻る編集（v1 → v2 → v1）── を作る。
    ここでは世代列に v1 が 2 回現れるのが**正しい履歴**なので、重複を防ぐ側では解けない。
    内容の等値で現在地を決めていると、必ず古い方（または新しい方）に張り付いて歩みが壊れる。
    """
    book = tmp_path / "b.xlsx"
    _write(book, "A")
    _edit(book, "B")
    _edit(book, "A")        # ★ 内容が A に戻った ── 世代列は [B, A]、現在の中身も A
    _edit(book, "C")        # 世代列は [A, B, A]
    assert [p.read_bytes() for p in ailine.list_backups(book)] == [b"A", b"B", b"A"], \
        "前提: 内容の重複した世代列になること"
    ailine.restore_backup(book)
    assert book.read_bytes() == b"A"        # 新しい方の A
    ailine.restore_backup(book)
    assert book.read_bytes() == b"B", "内容の等値で位置を決めている（新しい A に張り付いた）"
    ailine.restore_backup(book)
    assert book.read_bytes() == b"A"        # 古い方の A（＝最初の中身）
    with pytest.raises(ailine.NoOlderBackupError):
        ailine.restore_backup(book)


def test_identical_content_is_not_recorded_as_a_new_generation(tmp_path):
    """★ 世代列は「変化の履歴」であって「実行の履歴」ではない。

    undo の直後に run すると、いま復元したばかりの内容がもう一度積まれ、undo が
    **同じ中身を 2 回通る**（歩みは壊れないが、使う側には「効いていない」ように見える）。
    """
    # ★ 治具の訂正: _edit は「バックアップを取ってから書き換える」ので、直後の最新世代は
    #   **書き換える前**の中身になる（現在の中身とは違う）。同じ中身になるのは undo の直後
    #   ── そこが実際の事故の形なので、そのまま作る。
    book = tmp_path / "b.xlsx"
    _write(book, "v0")
    _edit(book, "v1")
    ailine.restore_backup(book)                 # v1 → v0（最新世代の中身 = いまの中身）
    before = len(ailine.list_backups(book))
    ailine.make_backup(book)                    # ★ undo 直後の run が積もうとする世代
    assert len(ailine.list_backups(book)) == before, "同じ中身の世代を積んだ"


def test_the_undo_shelf_still_records_identical_content(tmp_path):
    """★ 誤爆防止: 退避棚（undo 自体を可逆にする記録）は同じ中身でも残す。
       ここまで畳むと「undo をやり直す」材料が消える。"""
    book = tmp_path / "b.xlsx"
    _write(book, "v0")
    ailine.make_backup(book, shelf=True)
    n1 = len(ailine.list_undo_shelf(book))
    ailine.make_backup(book, shelf=True)
    assert len(ailine.list_undo_shelf(book)) == n1 + 1, "退避棚まで畳んでしまった"
