# パスの打ち間違いは、どの入口でも同じ返事になる（2026-09-13・買い手の初回体験 B2）
#
# ★★ なぜこの番人が要るか（測ってから作った・分母つき）: 「受領フォルダ」を
#   「受領フォルダ_typo」と打っただけで、**同じ事故に 4 通りの返事**が出ていた（12 経路の実測）:
#
#     scan / stack / forms          Python の**トレースバック**（FileNotFoundError）・exit 1
#     csv / export-csv / export-pdf 「文書が無い」・exit 1
#     split / accounts / verify      名指しの断り・exit 4
#     run                            「文書が無い」・exit 9（`ENGINEERING.md` の表どおり）
#
#   非プログラマが最も高い確率でやる間違いで、トレースバックは「道具が壊れた」に読める。
# ★ 分母は argparse から導く（下の `_path_routes`）── 表を手で並べると、新しい入口が
#   黙って抜ける。ここに無い入口が増えたら、この試験が先に赤くなる。
# ★ 番号は 9 を**直に**書く ── `input_path.MISSING_INPUT_EXIT` を読むと、定数を変えた日に
#   試験も一緒に動いて恒真になる（契約は `docs/ENGINEERING.md` の表）。
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine                                                   # noqa: E402

MISSING_INPUT_EXIT = 9          # ★ ENGINEERING.md の表「9 実行の前提が無い／文書が無い」

#: パスを受け取る位置引数の名前（argparse の dest）。
_PATH_DESTS = {"book", "file", "folder", "out"}

#: 打ち間違いを渡す形（★ 必須の指定は「在っても無くても結果が変わらない」ダミーで埋める）。
CASES = {
    "run":        lambda miss, work: ["run", str(miss), "合計を出して"],
    "csv":        lambda miss, work: ["csv", str(miss / "x.csv")],
    "export-csv": lambda miss, work: ["export-csv", str(miss / "x.xlsx"), "--sheet", "Sheet1"],
    "export-pdf": lambda miss, work: ["export-pdf", str(miss / "x.xlsx")],
    "scan":       lambda miss, work: ["scan", str(miss)],
    "stack":      lambda miss, work: ["stack", str(miss), "--out", str(work / "o.xlsx")],
    "forms":      lambda miss, work: ["forms", str(miss), "--out", str(work / "o.xlsx")],
    "split":      lambda miss, work: ["split", str(miss / "x.xlsx"), "--by", "担当",
                                      "--out", str(work / "配る")],
    "accounts":   lambda miss, work: ["accounts", str(miss / "x.csv"), "--past", str(miss),
                                      "--out", str(work / "o.xlsx")],
    "verify":     lambda miss, work: ["verify", str(miss / "o.xlsx"), str(miss)],
}

#: 「無い」が**別の事故**の入口（宣言 ── 理由つき・番人はこの宣言も分母に数える）。
NOT_A_MISSING_INPUT = {
    "restore": "原本の有無でなく**バックアップ**の有無の話（『バックアップが無い』と言う）",
    "undo": "同上 ── 戻せる世代が無いことを言う入口",
    "redo": "直前の undo が在るかの話（原本が在っても『やり直せません』になる）",
}


def _path_routes() -> set:
    """位置引数にパスを取る入口の名前（★ argparse から導く ── 手で並べない）。"""
    sub = [ac for ac in ailine.build_parser()._actions
           if isinstance(ac, argparse._SubParsersAction)][0]
    return {name for name, sp in sub.choices.items()
            if any(not ac.option_strings and ac.dest in _PATH_DESTS for ac in sp._actions)}


def _run(argv, cwd) -> subprocess.CompletedProcess:
    env = dict(os.environ)
    env["PYTHONPATH"] = str(REPO / "src") + os.pathsep + env.get("PYTHONPATH", "")
    return subprocess.run([sys.executable, "-m", "ailine", *argv], cwd=str(cwd), env=env,
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=300)


def test_the_table_covers_every_route_that_takes_a_path():
    """★ 分母の番人 ── 新しい入口が増えたら、ここが先に赤くなる（黙って抜けない）。"""
    covered = set(CASES) | set(NOT_A_MISSING_INPUT)
    routes = _path_routes()
    assert routes - covered == set(), f"打ち間違いを測っていない入口がある: {sorted(routes - covered)}"
    assert covered - routes == set(), f"もう無い入口が表に残っている: {sorted(covered - routes)}"
    assert len(CASES) >= 10, f"分母が縮んでいる（{len(CASES)} 経路）"


def test_a_typo_in_the_path_never_shows_a_traceback(tmp_path):
    """★★ 買い手が見るもの ── トレースバックは「自分の打ち間違い」に読めない。"""
    miss = tmp_path / "受領フォルダ_typo"
    checked = []
    for name, build in sorted(CASES.items()):
        r = _run(build(miss, tmp_path), tmp_path)
        both = (r.stdout or "") + (r.stderr or "")
        assert "Traceback (most recent call last)" not in both, f"{name}:\n{both[-800:]}"
        checked.append(name)
    assert len(checked) == len(CASES), checked


def test_a_typo_in_the_path_always_exits_nine_and_names_the_path(tmp_path):
    """★ 番号は 1 つ（自動化が「打ち間違い」と「中身が決まらない（4）」を見分けられる）。
    ★ 名前も必ず出す ── どのパスが無いのか分からない断りは、直せない断り。"""
    miss = tmp_path / "受領フォルダ_typo"
    bad = []
    for name, build in sorted(CASES.items()):
        r = _run(build(miss, tmp_path), tmp_path)
        both = (r.stdout or "") + (r.stderr or "")
        if r.returncode != MISSING_INPUT_EXIT or "受領フォルダ_typo" not in both:
            bad.append((name, r.returncode, both.strip().splitlines()[:2]))
    assert not bad, f"打ち間違いの返事が揃っていない: {bad}"


def test_the_exit_code_is_the_one_the_table_documents():
    """★ 契約は文書側（`docs/ENGINEERING.md` の表）── 定数を読み合って恒真にしない。"""
    table = (REPO / "docs" / "ENGINEERING.md").read_text(encoding="utf-8")
    row = [ln for ln in table.splitlines() if ln.startswith(f"| {MISSING_INPUT_EXIT} |")]
    assert len(row) == 1, f"表に 9 の行が {len(row)} 本ある"
    assert "パス" in row[0] or "打ち間違い" in row[0], (
        "表の 9 に『パスの打ち間違い』が書かれていない ── 番号だけ揃えて文書に無いなら、"
        f"買い手も自動化も知りようがない: {row[0]}")
