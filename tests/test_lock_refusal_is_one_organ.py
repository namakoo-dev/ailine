# 初回体験の盲検 3 回目・CONFUSING 3（2026-08-26）── 同じ判断が 4 箇所に書き写され、
# 3 通りに散らばっていた。
#
#   run（単一ブック）: 断定しない・心当たり 2 行（08-24 に直した正しい形）
#   undo:             断定しない・心当たり 1 行（最後の 1 行が抜けている）
#   run（2 冊照合）:   「Excel で開かれています」と**断定**・しかも
#                      `{lock_a}` で **タプルをそのまま印字**していた
#
# ★ 直しは「3 つとも直す」ではなく **1 つに畳んで呼び出し側に持たせない**。
#   番人も 1 本で全経路を縛る（変異試験で同時に赤くなることを確認済み）。
#
# 契約:
#   ① check_excel_lock を呼ぶのは refuse_if_locked ただ 1 箇所
#   ② 書き込めないだけの時に「Excel で開かれています」と原因を断定しない
#   ③ タプルの生表示をしない
#   ④ ロックファイルが在る時は、残骸の可能性と在り処を言う（重大7 の開示）
#   ⑤ ロックが無ければ何も言わず None（誤爆しない）

import inspect
import re
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from _product_source import product_text  # noqa: E402 ── ★ 番人は製品コード全体を読む

SRC = product_text()   # ★ 本体決め打ちだと、実装が ailine_core へ移った日に空振りする


def test_only_one_place_calls_the_detector():
    """① 書き写しを構造で禁じる。"""
    callers = [ln for ln in SRC.splitlines()
                if "check_excel_lock(" in ln and not ln.lstrip().startswith("def ")]
    assert len(callers) == 1, (
        "検出器の呼び出しが 1 箇所でない（書き写しが復活した）: " + " / ".join(callers))
    assert "reason = check_excel_lock(book)" in callers[0]


def test_unwritable_file_is_not_blamed_on_excel(tmp_path, capsys, monkeypatch):
    """②③ 書き込めないだけの時に原因を断定しない・タプルを見せない。"""
    book = tmp_path / "b.xlsx"
    book.write_bytes(b"x")
    monkeypatch.setattr(ailine, "check_excel_lock",
                         lambda p: ("unwritable", "このファイルに書き込めません"))
    rc = ailine.refuse_if_locked(book)
    out = capsys.readouterr().out
    assert rc == ailine.EXIT_WRITE_BLOCKED
    assert "Excel で開かれています" not in out, f"原因を断定した: {out}"
    assert "心当たり" in out, out
    assert not re.search(r"\('[a-z]+',", out), f"タプルを生で印字した: {out}"


def test_lock_file_discloses_that_it_may_be_a_leftover(tmp_path, capsys, monkeypatch):
    """④ 重大7 の開示: Excel を開いていないのに出る人に、次の手を渡す。"""
    book = tmp_path / "見積 書.xlsx"
    book.write_bytes(b"x")
    monkeypatch.setattr(ailine, "check_excel_lock",
                         lambda p: ("excel", "Excel のロックファイル ~$見積 書.xlsx が在ります"))
    ailine.refuse_if_locked(book)
    out = capsys.readouterr().out
    assert "残骸" in out, f"残骸の可能性に触れていない（恒久的な行き止まり）: {out}"
    assert "~$見積 書.xlsx" in out, f"在り処を出していない: {out}"


def test_silent_when_nothing_is_locked(tmp_path, capsys, monkeypatch):
    """⑤ 誤爆しない。"""
    book = tmp_path / "b.xlsx"
    book.write_bytes(b"x")
    monkeypatch.setattr(ailine, "check_excel_lock", lambda p: None)
    assert ailine.refuse_if_locked(book) is None
    assert capsys.readouterr().out == ""


@pytest.mark.parametrize("fn", ["_cmd_undo_body", "_cmd_run_body", "cmd_run_match"])
def test_every_write_path_goes_through_the_organ(fn):
    """★ 4 経路が同じ器官を通る（run 単一・undo・2 冊照合）。

    どれか 1 つが自前で判定に戻ったら赤くなる ── 片配線の再発を構造で止める。
    """
    if not hasattr(ailine, fn):
        pytest.skip(f"{fn} が無い（経路の名前が変わった ── 契約を読み直すこと）")
    src = inspect.getsource(getattr(ailine, fn))
    assert "refuse_if_locked(" in src, f"{fn} が共通の器官を通っていない"
    assert "check_excel_lock(" not in src, f"{fn} が検出器を直に呼んでいる（書き写し）"
