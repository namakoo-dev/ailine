# -*- coding: utf-8 -*-
"""前提は**道具が自分で満たす** ── 式のままの列を人に開かせない（2026-09-20）。

★★ 出所（盲検 3 体目・製造業の購買）: 『金額』が式のままで計算結果を持たない冊に対して、
  道具はこう言っていた ──「Excel か LibreOffice で一度開いて保存すると値が入ります」。
  ★ ところが **ailine は LibreOffice を持っている**。`normalize_book`（コピーを開いて
    保存する）は単一ブックの run が毎回通っている器官で、照合だけが人に頼んでいた。
    **器官は在るが配線が無い** ── この repo が何度も踏んだ形。

★★ 直したあとの約束:
  ① 邪魔をした時だけ払う（列が決まる回に LibreOffice 往復は起きない）
  ② **原本には触らない**（コピーを開く）
  ③ **誰が計算した値かを言う**（LibreOffice が計算した、と画面に出す）
  ④ 入れられなければ**落ちずに断る**（LibreOffice が無い環境がある）
  ⑤ ★★ 検算は**使ったデータ**を読む ── ここを外して実測で破れた
    （『A側合計の独立再集計(ナット) 元 0 / 出力 80』）。原本を読み直すと、こちらが
    使ったものと違うものを検算することになり、必ず破れる。
"""
from __future__ import annotations

import argparse
import contextlib
import io
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402


def _book(p: Path, headers: list, rows: list) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(p)
    wb.close()
    return p


def _formula_pair(d: Path) -> tuple:
    """A は金額が**式のまま**（計算結果なし）、B は値つき。"""
    a = _book(d / "A.xlsx", ["品名", "金額"], [["ボルト", "=100*2"], ["ナット", "=40*2"]])
    b = _book(d / "B.xlsx", ["品名", "金額"], [["ボルト", 150], ["ワッシャー", 300]])
    return a, b


def _run(a: Path, b: Path, task: str, home: Path):
    return subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(a), str(b), task],
        cwd=str(REPO), capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=900,
        env={**os.environ, "PYTHONPATH": str(REPO / "src"), "AILINE_HOME": str(home)})


@pytest.mark.local
def test_the_tool_fills_in_the_values_itself(tmp_path):
    """★★ 本体 ── 人に頼まず、道具が自分で値を入れて**到達する**こと。

    ★ 旧版はここで exit 3 で断り、「開いて保存してください」と言っていた。
    """
    a, b = _formula_pair(tmp_path)
    before = a.read_bytes()
    r = _run(a, b, "金額で突き合わせて", tmp_path / "home")
    assert r.returncode == 0, "★ まだ断っている:\n" + r.stdout[-700:]
    assert "LibreOffice で開いて計算させました" in r.stdout, r.stdout[-500:]
    # ★ ③ 誰が計算した値かを言う（観測していないことを主張しない）
    assert "LibreOffice が計算したもの" in r.stdout, r.stdout[-500:]
    # ★ ② 原本には触らない（1 バイトも）
    assert a.read_bytes() == before, "★ 原本を書き換えている"


@pytest.mark.local
def test_the_independent_check_reads_what_was_actually_used(tmp_path):
    """★★ ⑤ 検算が**使ったデータ**を読むこと ── 実測で破れた所。

    ★ 独立検算は原本を `xml_readback`（openpyxl とは別実装）で読み直す。埋めた回に
      原本を読むと、値の入っていない側と突き合わせることになり **必ず破れる**
      （実測: 『A側合計の独立再集計(ナット) 元 0 / 出力 80』）。
    ★ 独立性は「読み手の実装が別」であることで担保される ── 読む**冊**は
      実際に使ったものでなければ、検算が別の問いに答えてしまう。
    ★ だから数まで見る: 式 `=100*2` `=40*2` の結果が出力に出ていること。
    """
    a, b = _formula_pair(tmp_path)
    r = _run(a, b, "金額で突き合わせて", tmp_path / "home")
    assert r.returncode == 0, r.stdout[-700:]
    assert "事後条件が破れた" not in r.stdout, "★ 検算が別のデータを読んでいる:\n" + r.stdout[-700:]
    # ★ =100*2 → 200（B は 150 なので +50）、=40*2 → 80
    assert "+50" in r.stdout and "200" in r.stdout, r.stdout[-700:]
    assert "80" in r.stdout, r.stdout[-700:]


def test_nothing_is_paid_when_the_columns_already_resolve(tmp_path, monkeypatch):
    """★ ① 邪魔をしていない回に LibreOffice 往復を払わないこと。

    ★ 毎回払うと、照合が数秒遅くなる（列が決まる回は圧倒的に多い）。
    ★ 器官を呼んだら赤くする ── 「呼んでいない」を機械で確かめる。
    """
    called = []
    monkeypatch.setattr(ailine, "read_with_values_filled_in",
                        lambda *a, **k: called.append(1) or (None, "呼ぶな"))
    # ★★ 検体は「式のままの列が**在る**のに、邪魔していない」形にする（2026-09-20）。
    #   式の列が 1 つも無い冊だと、どちらの分岐でも往復は起きず、**変異が素通りする**
    #   （実際に変異試験が指した ── 検体が弱かった）。
    #   ここでは 備考 が式のままだが、依頼文がキーを名指ししているので**最初から決まる**。
    #   ★ 依頼文で名指ししないと、式のままの 備考 は**非数値なのでキー候補に入り**、
    #     最初の解決が失敗する（埋めると数値になって候補から外れる ── 実測で気づいた）。
    a = _book(tmp_path / "A.xlsx", ["品名", "金額", "備考"],
              [["ボルト", 120, "=1+1"], ["ナット", 80, "=2+2"]])
    b = _book(tmp_path / "B.xlsx", ["品名", "金額"], [["ボルト", 150], ["ワッシャー", 300]])
    monkeypatch.setenv("AILINE_HOME", str(tmp_path / "home"))
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = ailine.cmd_run_match(argparse.Namespace(json=False), a, b,
                                   "品名をキーに金額で突き合わせて")
    assert rc == 0, buf.getvalue()[-500:]
    assert not called, "★ 列が決まっているのに LibreOffice 往復を払っている"


def test_it_refuses_instead_of_crashing_when_libreoffice_cannot_open(tmp_path, monkeypatch):
    """★★ ④ 入れられない環境で**落ちない**こと（断りへ落ちる）。

    ★ `normalize_book` は失敗すると `SystemExit(9)` を投げる。そのまま抜けると、
      照合が前提の話で落ちる ── 買い手には「壊れた」としか見えない。
    ★★ そして**既にこちらが試した回**は、同じことを人に頼まない ──
      「開いて保存してください」と言えば、それは**通らない道を示す**形になる
      （導通の盤で一番重い失敗）。
    """
    def boom(book, workdir, timeout=None):
        raise SystemExit(9)

    monkeypatch.setattr(ailine, "normalize_book", boom)
    monkeypatch.setenv("AILINE_HOME", str(tmp_path / "home"))
    a, b = _formula_pair(tmp_path)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = ailine.cmd_run_match(argparse.Namespace(json=False), a, b, "金額で突き合わせて")
    out = buf.getvalue()
    assert rc == 3, f"★ 前提の話で落ちている（exit {rc}）:\n{out[-500:]}"
    assert "入れられませんでした" in out, out[-500:]
    assert "開いて保存すると値が入ります" not in out, (
        "★ こちらが既に試したのに、同じことを人に頼んでいる（通らない道）:\n" + out[-500:])


def test_the_organ_never_lets_an_exit_escape(tmp_path, monkeypatch):
    """★ 器官そのものの契約 ── どんな失敗でも `(None, 理由)` を返すこと。

    ★ ここが漏れると、呼び出し側がどれだけ丁寧でも前提の話で落ちる。
    """
    for boom in (SystemExit(9), RuntimeError("LibreOffice が居ない")):
        def raiser(book, workdir, timeout=None, _e=boom):
            raise _e
        monkeypatch.setattr(ailine, "normalize_book", raiser)
        got, why = ailine.read_with_values_filled_in(tmp_path / "x.xlsx", tmp_path / "w")
        assert got is None and why, (boom, got, why)
