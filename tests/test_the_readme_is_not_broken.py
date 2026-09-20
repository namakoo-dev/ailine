# -*- coding: utf-8 -*-
"""README が**壊れていないこと**を機械が見る（2026-09-21）。

★★ 出所（自分の事故）: 2026-09-20 に `--column` の説明を足した時、2 つ同時に壊した ──

    ① 表の**途中**に文章を挿し込み、表が 2 つに割れた
       （`| ailine verify …` 以下 3 行が宙に浮いて、ただの行として表示されていた）
    ② コード例に `\\n` が**文字として**残った（ヒアドキュメントがバックスラッシュを食った）

  そのまま push まで通った。**番人が 1 つも鳴らなかった。**
  ★ README は買い手が**最初に読む唯一のもの**（盲検でそう決めている）。ここが壊れていると、
    道具がどれだけ正しくても伝わらない ── 5 体目は「README に書いてない機能が
    いちばん役に立った」と言っている（別件だが、同じ「README が仕事をしていない」形）。

★ 見るのは**形**だけ（文面の good/bad は人が決める）:
    ・表の行が、表の外に落ちていないか
    ・コード例に**エスケープの残骸**が混じっていないか
    ・案内しているコマンドが**実在する**か（環境確認と同じ線）
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

README = (REPO / "README.md").read_bytes().decode("utf-8")
LINES = README.replace("\r\n", "\n").split("\n")


def _fenced() -> set:
    """コードブロックの中の行番号（0 起点）── そこは表でも散文でもない。"""
    inside, out, fence = False, set(), 0
    for i, ln in enumerate(LINES):
        if ln.startswith("```"):
            inside = not inside
            fence = i
            continue
        if inside:
            out.add(i)
    assert not inside, f"★ 閉じていないコードブロックがある（{fence + 1} 行目付近）"
    return out


def test_no_table_row_is_orphaned():
    """★★ 表の行が**表の外**に落ちていないこと（2026-09-20 に自分でやった）。

    ★ 表の行は `|` で始まる。その直前も `|` で始まるか、見出し区切り（`|---|`）の
      並びに属していなければ、それは**割れた表**。
    ★ 段落を表の途中に挿すと、そこから下は Markdown では表として描かれない ──
      買い手の画面では**ただの行**が並ぶ。
    """
    fenced = _fenced()
    orphans = []
    for i, ln in enumerate(LINES):
        if i in fenced or not ln.startswith("|"):
            continue
        prev = LINES[i - 1] if i else ""
        if prev.startswith("|"):
            continue
        # ★ 表の先頭行（見出し）は直後が区切り行であることで見分ける
        nxt = LINES[i + 1] if i + 1 < len(LINES) else ""
        if re.match(r"^\|[\s:|-]+\|$", nxt.strip()):
            continue
        orphans.append(f"{i + 1} 行目: {ln[:48]}")
    assert not orphans, (
        "★ 表の外に落ちている表の行がある（表を段落で割った）:\n  " + "\n  ".join(orphans))


def test_no_escape_leftovers_in_code_examples():
    """★★ コード例に**エスケープの残骸**が混じっていないこと。

    ★ 2026-09-20 に `\\n` が文字として残ったまま push した。買い手がそのまま打つと
      `\\n` 付きのコマンドを打つことになる ── **通らない道を示す**形。
    ★ ヒアドキュメントはバックスラッシュを食う（この repo が 1 日に何度も踏む）。
    """
    bad = []
    for i in sorted(_fenced()):
        for mark in ("\\n", "\\t", "\\\\"):
            if mark in LINES[i]:
                bad.append(f"{i + 1} 行目: {LINES[i][:60]}")
                break
    assert not bad, "★ コード例にエスケープの残骸がある:\n  " + "\n  ".join(bad)


def test_every_command_the_readme_shows_exists():
    """★ README が見せるサブコマンドが**実在する**こと。

    ★ 4 体目の買い手は「README に在るコマンドが無い」で 30 分溶かした。
      `blind_session check` が渡す前に同じことを測るが、**こちらは常に走る**。
    """
    import argparse

    import ailine

    parser = ailine.build_parser()
    subs = [a for a in parser._actions if isinstance(a, argparse._SubParsersAction)][0]
    have = set(subs.choices)
    shown = set(re.findall(r"ailine ([a-z][a-z-]+)", README))
    missing = sorted(shown - have)
    assert not missing, f"★ README に在って製品に無いコマンド: {missing}"


def test_the_two_book_form_is_documented():
    """★★ 2 冊の突き合わせが README に在ること（盲検 5 体目）。

    > README に書いてない機能が、いちばん役に立った。… 20 分は無駄にしました。

    ★ 買い手が月 2,000〜3,000 円を出すと言った唯一の機能が、載っていなかった。
    ★ ここは**在ること**だけを縛る（文面は人が決める）。
    """
    assert re.search(r"ailine run \S+\.xlsx \S+\.xlsx", README), (
        "★ 2 冊を並べる形が README に無い（買い手がいちばん欲しがった機能）")
