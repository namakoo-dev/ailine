# -*- coding: utf-8 -*-
"""買い手に渡す前の環境確認が、**赤くなれる**こと（2026-09-20）。

★★ 出所（4 体目の汚染 1・2 ── どちらもこちらの落ち度）:
  ・入っていた `ailine` が古く、買い手は「README に在るコマンドが無い」で **30 分**溶かした
  ・`AILINE_TRACE` が買い手のシェルに継承されておらず「環境を整えてある」が嘘になった
  ★ どちらも「設定したつもり」で止まっていた。

★★ この試験が守るのは **検査が緑を出せること**ではなく、**赤を出せること**。
  この repo が何度も踏んだ「在っても鳴らない」── 全部緑の検査は、守っているのか
  見ていないのか区別がつかない。だから壊れた環境を注入して、1 つずつ赤を確かめる。

★ 併せて、**測定が対象を汚さない**ことも縛る（試し打ちの 1 行が corpus に混ざると、
  冊の無い依頼として `freeze` に拾われる ── 実測で気づいた）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

import blind_session_core as core  # noqa: E402


@pytest.fixture
def session(tmp_path, monkeypatch):
    """使い捨ての 1 体（本物の corpus を触らない）。"""
    monkeypatch.setattr(core, "CORPUS", tmp_path / "blind")
    monkeypatch.setattr(core, "_version_line", lambda: "✓ ailine 版 (0.2.5)")
    (tmp_path / "blind" / "使い捨て" / "home").mkdir(parents=True)
    return "使い捨て"


def _ok_probe(home: Path, trace: Path) -> tuple:
    """効いている環境を真似る ── この回の home と trace の両方に 1 行足す。"""
    for p in (home / "history.jsonl", trace):
        p.parent.mkdir(parents=True, exist_ok=True)
        with p.open("a", encoding="utf-8") as f:
            f.write('{"ok": true}\n')
    return True, ""


def _rows(name, probe) -> dict:
    """★ 鍵は**検査が持つ** ── 文面から切り出さない（赤い回は文面が変わる）。"""
    return {key: ok for key, ok, _text in core.check(name, run_fn=probe)}


def test_a_working_environment_is_green(session):
    """★ 陰性対照 ── 効いている環境では全部緑（ここが赤いと下の試験が読めない）。"""
    rows = core.check(session, run_fn=_ok_probe)
    assert all(ok for _k, ok, _t in rows), [t for _k, ok, t in rows if not ok]
    assert len(rows) >= 6, f"検査の項目が減っている: {len(rows)}"


def test_an_old_installed_version_is_red(session, monkeypatch):
    """★★ 4 体目の汚染 1 ── 入っている版が古い回を赤にすること。"""
    monkeypatch.setattr(core, "_version_line",
                        lambda: "× ailine 版 — 0.2.5 が走っていますが、記録は 0.1.1 です")
    assert _rows(session, _ok_probe)["版"] is False


def test_a_command_the_readme_promises_but_the_product_lacks_is_red(session, monkeypatch):
    """★★ 4 体目の汚染 1 の**本体** ── 買い手が最初に触るのは README だけ。

    ★ 「README に在って製品に無い」を赤にする。逆（製品に在って README に無い）は
      赤にしない ── それは不足であって**嘘**ではない。
    """
    monkeypatch.setattr(core, "readme_commands",
                        lambda: core.parser_commands() | {"そんなコマンドは無い"})
    assert _rows(session, _ok_probe)["README"] is False
    monkeypatch.setattr(core, "readme_commands", lambda: {"run"})
    assert _rows(session, _ok_probe)["README"] is True, (
        "★ 製品に在って README に無い分まで赤にしている（不足と嘘を混ぜている）")


def test_a_leak_into_the_main_history_is_red(session, monkeypatch):
    """★★ 隔離が効いていない回を赤にすること（本体の履歴に漏れた）。

    ★ 「この回に増えた」だけでは隔離の証拠にならない ── **本体が増えていない**ことまで
      見て、はじめて効いていると言える。そこを外す変異がここで赤くなる。
    """
    counts = {"n": 0}
    real = core._count_lines

    def leaky(p: Path) -> int:
        # ★ 本体の履歴が試し打ちで 1 行増えた体にする
        if p == Path.home() / ".ailine" / "history.jsonl":
            counts["n"] += 1
            return counts["n"]
        return real(p)

    monkeypatch.setattr(core, "_count_lines", leaky)
    assert _rows(session, _ok_probe)["AILINE_HOME"] is False


def test_a_trace_that_records_nothing_is_red(session):
    """★★ 4 体目の汚染 2 ── 控えが 1 行も増えない回を赤にすること。"""
    def no_trace(home: Path, trace: Path) -> tuple:
        p = home / "history.jsonl"
        with p.open("a", encoding="utf-8") as f:
            f.write('{"ok": true}\n')
        return True, ""
    assert _rows(session, no_trace)["AILINE_TRACE"] is False


def test_a_probe_that_fails_is_red(session):
    """★ 試し打ちそのものが落ちた回を赤にすること（前提が無い環境）。"""
    assert _rows(session, lambda h, t: (False, "× ollama 不通"))["試し打ち"] is False


def test_the_check_leaves_no_trace_of_itself(session):
    """★★ 測定が対象を汚さないこと ── 試し打ちの跡を消して返す。

    ★ 残すと、冊の無い依頼として `freeze` に拾われ、corpus が自分の足跡で濁る。
    ★ 元から在った行は**消さない**（消すと買い手の記録を壊す）── バイトごと戻す。
    """
    d = core.CORPUS / session
    hist = d / "home" / "history.jsonl"
    hist.write_bytes(b'{"before": 1}\n')
    trace = d / "argv.jsonl"
    trace.write_bytes(b'{"before": 1}\n')
    rows = core.check(session, run_fn=_ok_probe)
    assert all(ok for _k, ok, _t in rows), [t for _k, ok, t in rows if not ok]
    assert hist.read_bytes() == b'{"before": 1}\n', "★ 元の履歴を壊している"
    assert trace.read_bytes() == b'{"before": 1}\n', "★ 元の控えを壊している"


def test_a_missing_session_says_so_instead_of_pretending(session):
    """★ prepare を走らせていない回は、緑を出さずにそう言うこと。"""
    import shutil
    shutil.rmtree(core.CORPUS / session / "home")
    rows = core.check(session, run_fn=_ok_probe)
    assert not all(ok for _k, ok, _t in rows), rows
    assert any("prepare" in t for _k, _ok, t in rows), rows
