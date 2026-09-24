# -*- coding: utf-8 -*-
"""壊れたブック（中身がテキストなのに拡張子だけ .xlsx）は、どの入口でも人の言葉で断る（2026-09-13）。

★★ 買い手役 3 体の初見で、事務職が**離脱を宣言した**所: `run` に渡すと英語のトレースバック 30 行
  （`zipfile.BadZipFile`）。1 つ前の `.xls` では日本語で完璧に案内していたのに。`forms` は
  「読み込み失敗: BadZipFile」と例外名を生で見せていた。「無い」の門（`input_path.require_*`）の
  隣に「**開けない**」の門（`input_path.explain_unreadable`）が無かった。
★ 分母は argparse から導く（`test_a_typo_in_the_path_is_refused_the_same_way` と同じ形）。
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from test_a_typo_in_the_path_is_refused_the_same_way import _path_routes   # noqa: E402

def _adopt_onto_a_real_book(broken, folder, work):
    """adopt は壊れた**下書き**を、開ける原本へ清書しようとする形で測る（原本が無いと 9 の別の断りになる）。"""
    import openpyxl
    book = work / "原本.xlsx"
    if not book.exists():
        openpyxl.Workbook().save(book)
    return ["adopt", str(broken), str(book)]


#: 壊れたブックを渡す形（★ `broken` は中身がテキストの .xlsx・`folder` はそれが 1 冊入ったフォルダ）。
CASES = {
    "run":        lambda broken, folder, work: ["run", str(broken), "合計を出して"],
    "export-csv": lambda broken, folder, work: ["export-csv", str(broken), "--sheet", "Sheet1"],
    "export-pdf": lambda broken, folder, work: ["export-pdf", str(broken)],
    "scan":       lambda broken, folder, work: ["scan", str(folder)],
    "stack":      lambda broken, folder, work: ["stack", str(folder), "--out", str(work / "o.xlsx")],
    "forms":      lambda broken, folder, work: ["forms", str(folder), "--out", str(work / "o.xlsx")],
    "split":      lambda broken, folder, work: ["split", str(broken), "--by", "担当",
                                                "--out", str(work / "配る")],
    "accounts":   lambda broken, folder, work: ["accounts", str(broken), "--past", str(work / "過去.xlsx"),
                                                "--out", str(work / "o.xlsx")],
    "verify":     lambda broken, folder, work: ["verify", str(broken), str(folder)],
    "accounts-apply": lambda broken, folder, work: ["accounts-apply", str(broken), str(broken),
                                                    "--out", str(work / "o.xlsx")],
    "adopt":      _adopt_onto_a_real_book,
}

#: 「壊れたブック」が当たらない入口（宣言・理由つき）。
NOT_A_BOOK_INPUT = {
    "csv": "受け取るのは CSV（テキスト）── 壊れた zip という事故が無い",
    "restore": "原本でなくバックアップの有無の話", "undo": "同上", "redo": "同上",
}

#: 断りに必ず入る語（次の一手 ── 例外名ではなく、人が動ける言葉）。
NEXT_STEP = "保存し直して"


def _run(argv, cwd):
    env = dict(os.environ)
    env["PYTHONPATH"] = str(REPO / "src") + os.pathsep + env.get("PYTHONPATH", "")
    return subprocess.run([sys.executable, "-m", "ailine", *argv], cwd=str(cwd), env=env,
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=300)


def _make(tmp_path):
    folder = tmp_path / "受領"
    folder.mkdir()
    broken = folder / "こわれた請求書.xlsx"
    broken.write_bytes("これはテキストです\n".encode("utf-8"))
    (tmp_path / "過去.xlsx").write_bytes(broken.read_bytes())   # accounts の --past 用（同じく壊れている）
    return broken, folder


def test_the_table_covers_every_route_that_takes_a_book():
    covered = set(CASES) | set(NOT_A_BOOK_INPUT)
    routes = _path_routes()
    assert routes - covered == set(), f"壊れたブックを測っていない入口: {sorted(routes - covered)}"
    assert covered - routes == set(), f"もう無い入口が表に残っている: {sorted(covered - routes)}"
    assert len(CASES) >= 9


def test_no_route_shows_a_traceback_or_a_raw_exception_name(tmp_path):
    broken, folder = _make(tmp_path)
    bad = []
    for name, build in sorted(CASES.items()):
        r = _run(build(broken, folder, tmp_path), tmp_path)
        both = (r.stdout or "") + (r.stderr or "")
        if "Traceback (most recent call last)" in both:
            bad.append((name, "トレースバック"))
        for raw in ("BadZipFile", "InvalidFileException", "zipfile"):
            if raw in both:
                bad.append((name, f"例外名 {raw}"))
    assert not bad, bad


def test_every_route_names_the_book_and_says_what_to_do(tmp_path):
    """★ 「壊れている」だけでは動けない ── 保存し直せ、まで言う。"""
    broken, folder = _make(tmp_path)
    bad = []
    for name, build in sorted(CASES.items()):
        r = _run(build(broken, folder, tmp_path), tmp_path)
        both = (r.stdout or "") + (r.stderr or "")
        if "こわれた請求書.xlsx" not in both or NEXT_STEP not in both:
            bad.append((name, r.returncode, both.strip().splitlines()[:3]))
    assert not bad, f"名指しか次の一手が無い: {bad}"
