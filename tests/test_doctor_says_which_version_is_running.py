# -*- coding: utf-8 -*-
"""`ailine doctor` が、いま走っている版を言う（2026-09-19・盲検 4 体目 ⑥）。

★★ 出所: 買い手に渡した `ailine` が **v0.1.0（古い版）**で、README に在るコマンドが
  無く 30 分溶けた。こちらは doctor を見せて「環境は整っている」と言ったが、
  **doctor は版を一言も言わなかった**ので、買い手も俺たちも気づけなかった。
  ★ 4 体目の汚染 3 件のうち、**そこから出た唯一の本物**がこれ。

★★ 版の正は**実体の在り処で決める**（2026-09-19 に実測）:

      作業木から実行中   importlib.metadata は **0.1.1**（インストール側の古い記録）
      走っているコード    作業木 ＝ pyproject.toml の 0.2.5

  ★ metadata だけを出すと **doctor が嘘をつく**。作業木では pyproject.toml が正。
  ★ 配られた版（wheel）に pyproject.toml は同梱されない ── そこでは metadata が正。
  ★★ 両方在って食い違う回は、**食い違いそのものを言う** ── それが事故の形だった。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def test_doctor_names_the_version_at_the_top():
    """★★ 事故そのもの: doctor の出力に版が在り、しかも**一番上**に在ること。

    ★ 買い手が README と突き合わせる最初の 1 行で、渡した側が汚染に気づく手掛かり。
    """
    names = [n for n, _ok, _d in ailine.doctor_checks()]
    assert names and names[0] == "ailine 版", names
    text, _all_ok = ailine.format_doctor_report(ailine.doctor_checks())
    first = text.splitlines()[0]
    assert "ailine 版" in first, first
    assert re.search(r"\d+\.\d+\.\d+", first), f"版の数字が出ていない: {first}"


def test_the_version_matches_pyproject(monkeypatch):
    """★ 版を 2 か所に書かない ── 作業木では pyproject.toml が正。

    ★ ここが無いと、画面の版だけが手書きで残って静かにずれる。
    """
    want = re.search(r'(?m)^version\s*=\s*"([^"]+)"',
                     (REPO / "pyproject.toml").read_bytes().decode("utf-8"))
    assert want, "pyproject.toml に version が無い"
    assert ailine._pyproject_version() == want.group(1)
    # ★ インストール記録と揃っている時は、その版がそのまま出る
    monkeypatch.setattr(ailine, "_installed_version", lambda: want.group(1))
    ok, detail = ailine.version_report()
    assert ok and detail.startswith(want.group(1)), (ok, detail)


def test_a_stale_install_is_named(monkeypatch):
    """★★ 4 体目の汚染そのもの: 走っている版とインストール記録が食い違えば、そう言う。

    ★ これが「器官は在るが鳴らない」にならないよう、**偽装して鳴ることを確かめる**。
    """
    monkeypatch.setattr(ailine, "_pyproject_version", lambda: "0.2.5")
    monkeypatch.setattr(ailine, "_installed_version", lambda: "0.1.0")
    ok, detail = ailine.version_report()
    assert ok is False, "★ 食い違っているのに ✓ を出している"
    assert "0.2.5" in detail and "0.1.0" in detail, detail
    assert "pip install" in detail, "★ 直し方を言っていない（断りの 5 条）"


def test_a_shipped_build_reports_the_installed_version(monkeypatch):
    """★ 配られた版（pyproject.toml が無い）では、インストール記録が正。"""
    monkeypatch.setattr(ailine, "_pyproject_version", lambda: None)
    monkeypatch.setattr(ailine, "_installed_version", lambda: "0.2.5")
    ok, detail = ailine.version_report()
    assert ok and detail.startswith("0.2.5"), (ok, detail)
    assert "インストール済み" in detail, detail


def test_an_unknown_version_is_not_a_check_mark(monkeypatch):
    """★ 版が分からない回に ✓ を出さない（分からないことを分かると言わない）。"""
    monkeypatch.setattr(ailine, "_pyproject_version", lambda: None)
    monkeypatch.setattr(ailine, "_installed_version", lambda: None)
    ok, detail = ailine.version_report()
    assert ok is False and "分かりません" in detail, (ok, detail)


def test_the_version_is_read_not_hardcoded():
    """★★ 恒真殺し: 版の数字を製品のコードに書き写していないこと。

    ★ 書き写すと、pyproject を上げた日に画面だけ古いままになる（この repo が
      何度も踏んだ「数は手書きしない」）。
    """
    from _product_source import window_around
    body = window_around("def _pyproject_version(", after=900)
    assert "pyproject.toml" in body, "★ 版の出所が pyproject でない"
    assert not re.search(r'"\d+\.\d+\.\d+"', body), "★ 版の数字がコードに書いてある"
