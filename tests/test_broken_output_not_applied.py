# 復元の中10（2026-08-24 の盲検）── 壊れた成果物を検証せずに原本へ被せる。
#
# ★ 実測: 反映前の関門に「開ける xlsx か」の検査が無く、zip として読めない成果物を
#   そのまま原本へ被せてから
#       ⚠ c.xlsx に適用しましたが、読み戻して確認できませんでした（File is not a zip file）
#   と言っていた。★ 報告は正直だが、**確認は原本を潰した後**だった。順序が逆。
#   （この時は undo で復旧できたが、命綱に頼るべき場面ではない）
#
# 契約:
#   ① 開けない成果物は原本に被せない（原本は 1 バイトも変わらない）
#   ② 止めた理由を言い、成果物は残す（人が中を見られる）
#   ③ 中身の正しさは見ない ── ここが見るのは「開けるか」だけ（事後条件の仕事を奪わない）
#   ④ --copy は原本に触らないので対象外

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402
from test_golden_transcripts import _isolate, _run_main  # noqa: E402


def _book(tmp_path):
    p = tmp_path / "c.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["商品", "金額"]); ws.append(["a", 100]); ws.append(["b", 250])
    wb.save(p)
    return p


def _corrupting_apply(out_book, code, workdir, helper_files=(), timeout=None):
    Path(out_book).write_bytes(b"NOT A ZIP - truncated/corrupt result")
    return True, None, "ok"


def test_broken_output_never_touches_the_original(tmp_path, monkeypatch, capsys):
    """①② 原本は無傷・理由を言う・成果物は残す。"""
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path)
    before = book.read_bytes()
    monkeypatch.setattr(
        ailine, "translate_task",
        lambda model, task, book_meta, temperature=0.1:
        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "basrun_apply", _corrupting_apply)
    rc, out = _run_main(["run", str(book), "金額で降順に並べ替えて"], capsys)
    assert book.read_bytes() == before, "壊れた結果を原本に被せた"
    assert rc != 0, f"壊れているのに成功を名乗った: {out}"
    assert "壊れています" in out, f"止めた理由を言っていない: {out}"
    assert "変更していません" in out, f"原本が無事だと言っていない: {out}"
    assert "c.out.xlsx" in out, f"成果物の在り処を言っていない（人が中を見られない）: {out}"


def test_the_gate_only_checks_openability(tmp_path):
    """③ 中身の正しさは見ない（事後条件の仕事を奪わない）。"""
    p = _book(tmp_path)
    assert ailine._why_output_is_unusable(p) is None
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a zip")
    assert ailine._why_output_is_unusable(broken) is not None


def test_copy_path_is_out_of_scope(tmp_path, monkeypatch, capsys):
    """④ --copy は原本に触らないので、この関門は掛からない。"""
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path)
    before = book.read_bytes()
    monkeypatch.setattr(
        ailine, "translate_task",
        lambda model, task, book_meta, temperature=0.1:
        {"op": "SORT", "args": {"col": "金額", "order": "desc"}})
    monkeypatch.setattr(ailine, "basrun_apply", _corrupting_apply)
    rc, out = _run_main(["run", str(book), "金額で降順に並べ替えて", "--copy"], capsys)
    assert book.read_bytes() == before
