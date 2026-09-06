# -*- coding: utf-8 -*-
"""黙って死んだ回を、空のエラーとして返さない ── その番人（2026-09-06）。

★★ 出所（実測）: 実機テストが大量に落ちた回の画面が、これ 1 行だけだった ──

        AssertionError: basrun_apply が失敗した:

  `basrun` が **stdout も stderr も空**のまま非ゼロで終了し、呼び出し側が
  `raw.strip()[-800:]`＝空文字をエラーとして返していた。★ 空文字は
  「エラーが無い」ではなく「**黙って死んだ**」── まったく別の事実なのに、
  画面では区別がつかなかった。結果、**原因を今も特定できていない**
  （当時の推測「居残りプロセスのせい」は、後日の再現実験で否定された）。

★ ここで守るのは**正しさ**ではなく**次に読めること**。終了コードを添え、
  環境側を疑う導線まで出す。★ 原因が分からない事故に対して打てる手は、
  「次に起きた時に分かるようにする」だけのことがある。
"""
from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

import ailine


class _DeadProc:
    """何も出力せずに非ゼロで終わったプロセス（★ 実測で起きた形）。"""

    def __init__(self, code: int = 1):
        self.returncode = code
        self.pid = 4242

    def communicate(self, timeout=None):
        return "", ""


@pytest.fixture
def silent_death(monkeypatch):
    monkeypatch.setattr(subprocess, "Popen", lambda *a, **k: _DeadProc())
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("basrun.py"))


def _apply(tmp_path):
    code = "Sub Run(oDoc)" + chr(10) + "End Sub" + chr(10)
    return ailine.basrun_apply(tmp_path / "b.xlsx", code, tmp_path / "w")


def test_a_silent_death_is_named_not_left_blank(tmp_path, silent_death):
    ok, err, _raw = _apply(tmp_path)
    assert ok is False
    assert err and err.strip(), "空のエラーを返した（黙って死んだのか分からない）"
    assert "何も出力せず" in err, err
    assert "終了コード" in err, err


def test_the_exit_code_is_carried_so_the_next_person_can_read_it(tmp_path, monkeypatch):
    """★ 終了コードを**そのまま**運ぶこと（丸めると次回の手掛かりが消える）。"""
    monkeypatch.setattr(subprocess, "Popen", lambda *a, **k: _DeadProc(code=9))
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("basrun.py"))
    _ok, err, _raw = _apply(tmp_path)
    assert "9" in err, err


def test_a_real_error_is_still_passed_through(tmp_path, monkeypatch):
    """★ 対で縛る ── 中身が在る回は、今までどおりそれを返すこと。"""
    class _Noisy(_DeadProc):
        def communicate(self, timeout=None):
            return "", "Basic runtime error: Variable not defined"
    monkeypatch.setattr(subprocess, "Popen", lambda *a, **k: _Noisy())
    monkeypatch.setattr(ailine, "basrun_path", lambda: Path("basrun.py"))
    _ok, err, _raw = _apply(tmp_path)
    assert "Variable not defined" in err, err
    assert "何も出力せず" not in err, err
