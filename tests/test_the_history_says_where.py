# -*- coding: utf-8 -*-
"""実行履歴が「どの冊か」を言う（2026-09-14・買い手役 2 体が別の回に指した）。

★★ 「`history` が PC 単位（共有 PC だと他人の作業が混ざる）」「`モデル` 欄 None」。
  実測は報告より悪く、**別フォルダ・別ドライブの 3 冊が全部同じ `売上.xlsx`** に見えていた。
  `ailine undo` は冊のパスを引数に取る ── 履歴は「何をどの冊にしたか」を引く唯一の場所なのに、
  そこから**引数が作れなかった**。幅が溢れるより曖昧が悪い。

契約:
  - 『文書』の欄は場所まで出す（cwd の配下なら相対・外ならフルパス・記録が相対なら そのまま）
  - 1 行目に範囲と台帳の置き場所
  - `--folder` で絞れる。★ 絞った回は**隠した件数を必ず言う**（出ないことを信号にしない）
  - `--folder` の打ち間違いは「0 件」でなく exit 9 で名指しして断る
  - モデルを使っていない回は `None` でなく `—`
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core import cli_render   # noqa: E402

_ROWS = [
    {"ts": "2026-09-14T09:00:00", "ok": True, "attempts": 1, "model": "qwen2.5-coder:7b",
     "book": r"C:\仕事\9月請求\売上.xlsx", "task": "B列を太字にして"},
    {"ts": "2026-09-14T09:10:00", "ok": True, "attempts": 1, "model": None,
     "book": r"C:\仕事\10月請求\売上.xlsx", "task": "合計を足して"},
]


def _home(tmp_path):
    home = tmp_path / "home"
    home.mkdir()
    (home / "history.jsonl").write_text(
        "".join(json.dumps(r, ensure_ascii=False) + "\n" for r in _ROWS), encoding="utf-8")
    return home


def _run(home, *extra, cwd=None):
    import os
    env = dict(os.environ, AILINE_HOME=str(home), PYTHONPATH=str(REPO / "src"))
    return subprocess.run([sys.executable, "-m", "ailine", "history", *extra],
                          capture_output=True, text=True, encoding="utf-8", errors="replace",
                          env=env, cwd=str(cwd or REPO), timeout=120)


def test_two_books_with_the_same_name_are_told_apart(tmp_path):
    """★★ これが本体 ── 同名の 2 冊が別の行として読める（名前だけなら区別できない）。"""
    r = _run(_home(tmp_path))
    assert r.returncode == 0, r.stderr
    assert r.stdout.count("売上.xlsx") == 2, r.stdout
    assert "9月請求" in r.stdout and "10月請求" in r.stdout, r.stdout


def test_the_scope_and_the_ledger_are_named(tmp_path):
    home = _home(tmp_path)
    r = _run(home)
    assert "この PC のこのアカウント" in r.stdout, r.stdout
    assert str(home / "history.jsonl") in r.stdout, r.stdout


def test_a_run_without_a_model_is_not_shown_as_none(tmp_path):
    r = _run(_home(tmp_path))
    assert "None" not in r.stdout, r.stdout
    assert cli_render.NO_MODEL in r.stdout, r.stdout


def test_filtering_says_how_many_it_hid(tmp_path):
    """★ 出ないことを信号にしない ── 絞った回は隠した件数を言う。"""
    folder = tmp_path / "9月請求"
    folder.mkdir()
    rows = [dict(_ROWS[0], book=str(folder / "売上.xlsx")), _ROWS[1]]
    home = tmp_path / "h2"
    home.mkdir()
    (home / "history.jsonl").write_text(
        "".join(json.dumps(r, ensure_ascii=False) + "\n" for r in rows), encoding="utf-8")
    r = _run(home, "--folder", str(folder))
    assert r.returncode == 0, r.stderr
    assert "1 件" in r.stdout and "ほかの場所の記録 1 件は出していません" in r.stdout, r.stdout
    assert "10月請求" not in r.stdout, r.stdout


def test_an_empty_folder_does_not_claim_there_is_no_history(tmp_path):
    """★ 絞って 0 件の回に「履歴はまだ無い」と言わない（記録は在る ── 嘘になる）。"""
    empty = tmp_path / "空"
    empty.mkdir()
    r = _run(_home(tmp_path), "--folder", str(empty))
    assert r.returncode == 0, r.stderr
    assert "履歴はまだ無い" not in r.stdout, r.stdout
    assert "ほかの場所の記録 2 件は出していません" in r.stdout, r.stdout


def test_a_typo_in_the_folder_is_refused_not_counted_as_zero(tmp_path):
    """★★ 打ち間違いが「ここには記録がありません」に化けるのが一番悪い ── exit 9 で断る。"""
    r = _run(_home(tmp_path), "--folder", str(tmp_path / "ない場所"))
    assert r.returncode == 9, f"exit={r.returncode} / {r.stdout}{r.stderr}"
    assert "フォルダが見つかりません" in (r.stdout + r.stderr)


def test_a_relative_record_is_left_as_recorded():
    """★ 古い記録（相対パスのまま）は作り直さない ── 嘘の場所を書くより、記録のまま出す。"""
    assert cli_render.history_place("b.xlsx") == "b.xlsx"
    assert cli_render.history_place("") == "(パス不明)"


def test_a_book_under_the_current_folder_is_shown_relative(tmp_path):
    """★ cwd の配下は相対（そのまま `ailine undo <文書>` に貼れる）。"""
    got = cli_render.history_place(str(tmp_path / "出力" / "a.xlsx"), base=tmp_path)
    assert got == str(Path("出力") / "a.xlsx"), got
