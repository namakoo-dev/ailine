# 複数ファイルの盲検・致命③⑥（2026-08-26）── 嘘の成功報告と、生の traceback。
#
# 致命③: 金額列が全部数式（キャッシュ値なし）の請求書に「金額が1以上の行を抜き出して」
#   → 『2 中 0 ファイルで計 0 行一致』＋『行の完全会計: 成立』＋ exit 0。
#   「金額 1 円以上の請求は 1 件も無い」という**嘘の成功報告**。
#   ★ stack 側には警告が在るのに、run <フォルダ> の経路だけ無かった（片配線）。
#
# 致命⑥: 壊れた .xlsx が 1 本混ざると verify が **生の traceback** で落ちる（exit 1）。
#   同じ verify.py の中でも _attribution_mismatch は try/except を持ち、
#   _expected_rows_for_source は持たない ── **同一ファイル内の非対称**。
#
# 契約:
#   ① 読めなかったセルを「条件に合わなかった」と言わない（件数を名指しで開示）
#   ② 判定は変えない ── 読めないものを「合う」とは言わない
#   ③ 壊れた冊で生の traceback を出さない
#   ④ 読めなかった冊は分母から黙って消さず、名指しして非零で止まる

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402
from test_golden_transcripts import _isolate, _run_main  # noqa: E402


def _folder_with_formula_amounts(tmp_path):
    d = tmp_path / "seikyu"
    d.mkdir()
    for i, n in enumerate(("01.xlsx", "02.xlsx"), start=1):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "明細"
        ws.append(["日付", "得意先", "数量", "単価", "金額"])
        ws.append([f"2026-01-0{i}", f"{i}社", 2, 500, "=C2*D2"])
        wb.save(d / n)
    return d


def _plain_folder(tmp_path):
    d = tmp_path / "src"
    d.mkdir()
    for i, n in enumerate(("01.xlsx", "02.xlsx"), start=1):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "明細"
        ws.append(["日付", "得意先", "金額"])
        ws.append([f"2026-01-0{i}", f"{i}社", i * 1000])
        wb.save(d / n)
    return d


# --- ①② 嘘の成功報告 -----------------------------------------------------------------

def _fix_extract(monkeypatch):
    """翻訳の揺れを外す（測りたいのは抽出後の報告であって、翻訳精度ではない）。"""
    monkeypatch.setattr(
        ailine, "translate_task",
        lambda model, task, book_meta, temperature=0.1:
        {"op": "EXTRACT", "args": {"col": "金額", "cmp": "gte", "value": 1}})


def test_folder_run_does_not_call_unreadable_cells_a_mismatch(tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    _fix_extract(monkeypatch)
    d = _folder_with_formula_amounts(tmp_path)
    rc, out = _run_main(["run", str(d), "金額が1以上の行を抜き出して"], capsys)
    assert "計 0 行一致" in out, f"前提: この検体で 0 件になること: {out}"
    assert "数式" in out and "確かめられませんでした" in out, \
        f"『1 件も無い』としか言っていない（嘘の成功報告）: {out}"
    assert "合わなかったのではありません" in out, out


def test_folder_run_is_silent_when_values_are_there(tmp_path, monkeypatch, capsys):
    """誤爆しない: 値が入っていれば 1 文字も増えない。"""
    _isolate(monkeypatch, tmp_path)
    _fix_extract(monkeypatch)
    d = _plain_folder(tmp_path)
    rc, out = _run_main(["run", str(d), "金額が1以上の行を抜き出して"], capsys)
    assert "確かめられませんでした" not in out, out


# --- ③④ 壊れた冊 ---------------------------------------------------------------------

def test_verify_names_a_broken_workbook_instead_of_crashing(tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    d = _plain_folder(tmp_path)
    out_book = tmp_path / "stacked.xlsx"
    rc, _ = _run_main(["stack", str(d), "--out", str(out_book)], capsys)
    assert rc == 0
    (d / "zz_broken.xlsx").write_bytes(b"garbage")
    rc2, out = _run_main(["verify", str(out_book), str(d)], capsys)
    assert rc2 != 0, f"読めない冊が在るのに合格にした: {out}"
    assert "Traceback" not in out, out
    assert "zz_broken.xlsx を読めませんでした" in out, \
        f"読めなかった冊を名指ししていない: {out}"
    assert "分母に入っていません" in out, out


def test_stack_names_a_broken_workbook_instead_of_crashing(tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    d = _plain_folder(tmp_path)
    (d / "zz_broken.xlsx").write_bytes(b"garbage")
    rc, out = _run_main(["stack", str(d), "--out", str(tmp_path / "o.xlsx")], capsys)
    assert "Traceback" not in out, out
    assert rc != 0, out
    assert "zz_broken.xlsx" in out, out
