"""文書に書かれた数字を、**1 箇所で全文書ぶん**集める。

★★ 2026-08-31（提出用の文書を足した時に気づいた）:
  それまで数字の番人は `README.md` だけを名指しで読んでいた。
  評価者に見せる文書がもう 1 本増えた瞬間、**同じ数字が 2 箇所に在って
  片方だけ古くなる**形になる ── この repo で何度も踏んだ「片配線」そのもの。

★ 処方は「両方読む」ではなく、**呼び出し側に文書名を持たせない**こと。
  ここが repo 内の `.md` を全部走査して印を集めるので、
  文書を足しても番人の側は 1 行も変わらない（＝足した文書が黙って素通りしない）。

★ 分母は入力側から取る ── 「印が在る文書」を数えるのではなく、
  **全 .md を見てから**印を抜く。
"""

from __future__ import annotations

import re
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

# ★ 走査から外す所（生成物・依存・履歴）。ここを増やす時は理由を書くこと。
_SKIP = (".git", "__pycache__", "node_modules", ".venv", "dist", ".pytest_cache")


def all_markdown() -> list[Path]:
    """repo 内の .md を全部（除外先を除く）。"""
    out = []
    for p in REPO.rglob("*.md"):
        if any(part in _SKIP for part in p.parts):
            continue
        out.append(p)
    return sorted(out)


def marked(name: str) -> list[tuple[Path, str]]:
    """`<!-- NAME -->…<!-- /NAME -->` の中身を、見つかった文書ぶん全部返す。"""
    pat = re.compile(rf"<!--\s*{name}\s*-->(.*?)<!--\s*/{name}\s*-->", re.S)
    hits: list[tuple[Path, str]] = []
    for p in all_markdown():
        for m in pat.findall(p.read_text(encoding="utf-8")):
            hits.append((p, m.strip()))
    return hits


def assert_all_agree(name: str, want: str, *, at_least: int = 1) -> None:
    """印のある全文書が `want` と一致すること（1 つも無ければそれも赤）。

    ★ `at_least` は「印が消えた」を捕まえるため ── 印ごと消せば
      不一致は起きないので、**在ることも確かめる**（出ないことは信号ではない）。
    """
    hits = marked(name)
    assert len(hits) >= at_least, (
        f"{name} の印が {len(hits)} 箇所しか無い（{at_least} 箇所以上のはず）── "
        "印ごと消すと、番人は黙って通す")
    bad = [(str(p.relative_to(REPO)), got) for p, got in hits if got != want]
    assert not bad, f"{name} が古い文書がある: {bad}（正: {want!r}）"
