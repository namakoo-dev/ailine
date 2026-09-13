# -*- coding: utf-8 -*-
"""空のフォルダ・サブフォルダしか無いフォルダは、成功に見せない（2026-09-13・買い手役 3 体のうち 2 体）。

★★ `0 ファイル中 0 冊を読みました`・exit 0・ファイルを作らず、作らなかったとも言わなかった。
  フォルダを選び間違えた人が気づけない（エクスプローラで探して無くて戸惑った、と 2 体）。
  「文書が無い」は 9（ENGINEERING.md の表）── 打ち間違いと同じ番号・同じ家系。
★ 分母は argparse から導く（位置引数が `folder` の入口）。断りの文は `multifile.nothing_to_read` 1 本。

同じ日に直した「引数の間違いだけ英語」（2 体）の番人もここに置く ── どちらも
「間違えた瞬間に人の言葉が出るか」の話。
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine                                                   # noqa: E402

MISSING_INPUT_EXIT = 9

CASES = {
    "scan":  lambda folder, work: ["scan", str(folder)],
    "stack": lambda folder, work: ["stack", str(folder), "--out", str(work / "o.xlsx")],
    "forms": lambda folder, work: ["forms", str(folder), "--out", str(work / "o.xlsx")],
}


def _folder_routes() -> set:
    sub = [ac for ac in ailine.build_parser()._actions
           if isinstance(ac, argparse._SubParsersAction)][0]
    return {name for name, sp in sub.choices.items()
            if any(not ac.option_strings and ac.dest == "folder" for ac in sp._actions)}


def _run(argv, cwd):
    env = dict(os.environ)
    env["PYTHONPATH"] = str(REPO / "src") + os.pathsep + env.get("PYTHONPATH", "")
    return subprocess.run([sys.executable, "-m", "ailine", *argv], cwd=str(cwd), env=env,
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=300)


def test_the_table_covers_every_folder_route():
    routes = _folder_routes()
    assert routes == set(CASES), (sorted(routes), sorted(CASES))
    assert len(CASES) >= 3


def test_an_empty_folder_is_refused_by_name_and_nothing_is_written(tmp_path):
    empty = tmp_path / "空っぽ"
    empty.mkdir()
    bad = []
    for name, build in sorted(CASES.items()):
        r = _run(build(empty, tmp_path), tmp_path)
        both = (r.stdout or "") + (r.stderr or "")
        if r.returncode != MISSING_INPUT_EXIT or "1 冊もありません" not in both \
                or "ファイルは作っていません" not in both or (tmp_path / "o.xlsx").exists():
            bad.append((name, r.returncode, both.strip().splitlines()[:2]))
    assert not bad, bad


def test_a_folder_with_only_subfolders_says_it_did_not_look_inside(tmp_path):
    """★ 一番多い選び間違い ── 親フォルダを指した。中は見ないと言い、中を指せと言う。"""
    nested = tmp_path / "受領"
    (nested / "2026-09").mkdir(parents=True)
    checked = []
    for name, build in sorted(CASES.items()):
        r = _run(build(nested, tmp_path), tmp_path)
        assert r.returncode == MISSING_INPUT_EXIT, (name, r.returncode, r.stdout)
        assert "サブフォルダが 1 件" in r.stdout and "中のフォルダを直接指定" in r.stdout, (name, r.stdout)
        checked.append(name)
    assert len(checked) == len(CASES)


def test_a_folder_with_books_is_not_refused(tmp_path):
    """★ 陰性対照 ── 1 冊でも在れば今までどおり（断りを広げすぎない）。"""
    import openpyxl
    folder = tmp_path / "受領"
    folder.mkdir()
    wb = openpyxl.Workbook()
    wb.active.append(["商品", "金額"])
    wb.active.append(["a", 100])
    wb.save(folder / "a.xlsx")
    wb.close()
    r = _run(["scan", str(folder)], tmp_path)
    assert r.returncode == 0 and "1 冊もありません" not in r.stdout, r.stdout


# --- 引数の間違いも人の言葉で（exit 2 は表の契約のまま）------------------------------------

def test_a_missing_argument_is_explained_in_japanese(tmp_path):
    r = _run(["forms"], tmp_path)
    assert r.returncode == 2, (r.returncode, r.stderr)
    assert "必要な指定が足りません" in r.stderr and "<フォルダ>" in r.stderr, r.stderr
    assert "the following arguments are required" not in r.stderr, r.stderr


def test_an_unknown_flag_and_a_missing_command_are_explained_in_japanese(tmp_path):
    r = _run([], tmp_path)
    assert r.returncode == 2 and "必要な指定が足りません: <入口>" in r.stderr, r.stderr
    r = _run(["forms", "x", "--out", "y", "--dry-run"], tmp_path)
    assert r.returncode == 2 and "知らない指定があります: --dry-run" in r.stderr, r.stderr
    assert "ailine ops" in r.stderr, r.stderr


# --- 検疫した CSV は入力（2026-09-13・2 回目の買い手役・事務職が離脱した所）------------------
#
# ★★ 「CSV が届く → `ailine csv` で xlsx に → `stack`」という毎月の入口で、検疫の出力が「自分の出力」
#   として全部入力から外れて 0 冊になり、断りが「`ailine csv` で xlsx にできます」と**循環**した。
#   結果（縦積み・一覧・分けた冊…）は二重計上を防ぐため外す。検疫は入力への変換なので外さない。


def _quarantined(tmp_path, name, rows):
    src = tmp_path / f"{name}.csv"
    src.write_text("\n".join(",".join(map(str, r)) for r in rows) + "\n", encoding="utf-8-sig")
    r = _run(["csv", str(src)], tmp_path)
    assert r.returncode == 0, r.stdout
    out = tmp_path / f"{name}.xlsx"
    assert out.exists(), r.stdout
    return out


def test_books_made_by_ailine_csv_are_inputs_not_outputs(tmp_path):
    folder = tmp_path / "受領"
    folder.mkdir()
    for name in ("4月", "5月"):
        book = _quarantined(tmp_path, name, [["商品", "売上"], ["a", 100], ["b", 200]])
        book.replace(folder / book.name)
    r = _run(["stack", str(folder), "--out", str(tmp_path / "縦積み.xlsx")], tmp_path)
    assert r.returncode == 0, r.stdout
    assert "2 ファイル中 2 積んだ" in r.stdout, r.stdout
    assert "1 冊もありません" not in r.stdout


def test_a_folder_of_ailine_results_names_them_instead_of_circular_advice(tmp_path):
    """★ 結果しか無いフォルダ ── 「CSV を xlsx に」でなく「この冊は結果なので数えない」と言う。"""
    folder = tmp_path / "受領"
    folder.mkdir()
    (folder / "4月.csv").write_text("商品,売上\na,100\n", encoding="utf-8-sig")
    r = _run(["csv", str(folder / "4月.csv")], tmp_path)
    assert r.returncode == 0, r.stdout
    r = _run(["stack", str(folder), "--out", str(tmp_path / "縦積み.xlsx")], tmp_path)
    assert r.returncode == 0, r.stdout
    stale = tmp_path / "受領2"
    stale.mkdir()
    (tmp_path / "縦積み.xlsx").replace(stale / "前回の縦積み.xlsx")
    r = _run(["stack", str(stale), "--out", str(tmp_path / "x.xlsx")], tmp_path)
    assert r.returncode == MISSING_INPUT_EXIT, (r.returncode, r.stdout)
    assert "ailine が作った冊 1 冊は入力に数えていません: 前回の縦積み.xlsx" in r.stdout, r.stdout
