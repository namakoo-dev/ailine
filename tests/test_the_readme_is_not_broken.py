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
    ★★ 2026-09-21 に締め直した ── **初版は緑のまま買い手が見つけられなかった**。
      初版は「README のどこかに在るか」を見ていた。実際、例は載っていた（下の節）。
      買い手が見たのは**コマンド表**で、そこに 2 冊の行が無かったから「載っていない」と
      判断した。★ **在るかどうかでなく、探す人が見る場所に在るか**が問われていた
      （「在っても鳴らない」の人間側）。
    ★ 文面は人が決める ── 縛るのは**表の中に在ること**だけ。
    """
    lines = README.splitlines()
    rows = [i for i, ln in enumerate(lines) if ln.startswith("|") and "ailine scan" in ln]
    assert rows, "コマンド表が見つからない（`ailine scan` の行が目印）"
    i = rows[0]
    lo = hi = i
    while lo > 0 and lines[lo - 1].startswith("|"):
        lo -= 1
    while hi + 1 < len(lines) and lines[hi + 1].startswith("|"):
        hi += 1
    table = chr(10).join(lines[lo:hi + 1])
    assert re.search(r"ailine run <ブック> <ブック>", table), (
        "★ 2 冊を並べる形が**コマンド表**に無い ── 節に書いてあっても、"
        "表しか読まない人には無いのと同じ（買い手は 20 分を失った）")


# --- 先頭の道案内（2026-09-21・盲検 3 体目 ⑩ / 4 体目 ⑧）------------------------------

def _anchor(text: str) -> str:
    """GitHub が見出しから作る錨と同じ規則（小文字化・記号を落とす・空白を -）。"""
    import unicodedata
    t = text.strip().lower()
    keep = [c for c in t if c.isalnum() or c in " -_"
            or (ord(c) > 127 and not unicodedata.category(c).startswith("P"))]
    return "".join(keep).replace(" ", "-")


def test_every_internal_link_points_at_a_real_heading():
    """★★ README の中のリンクが**実在する見出し**を指すこと。

    ★ 「導線が嘘なら、導線が無いより悪い」── この repo が
      `tests/test_examples_actually_work.py` で出した結論を、文書の側にも置く。
    ★ 見出しを直した日に、リンクだけ古いまま残るのが一番静かな壊れ方。
    """
    heads = {_anchor(m.group(1)) for m in re.finditer(r"^#{1,6} (.+)$", README, re.M)}
    broken = [m.group(1) for m in re.finditer(r"\]\(#([^)]+)\)", README)
              if m.group(1) not in heads]
    assert not broken, f"飛び先の無いリンク: {broken}"


def test_the_router_is_at_the_top_and_names_the_monthly_runbook():
    """★★ 買い手が**最初に**行き先を見つけられること。

    ★★ 出所（盲検・2 人が独立に）:
      3 体目「README が約 700 行で『**月末に何を打てばいいか**』が無い」
      4 体目「README が 700 行で、買い手が**読む所を見つけられない**」
      ★ 月末の紙（docs/月末の締めのやり方.md・A4 1 枚）は**元から在った** ──
        リンクが表のセルの末尾と 600 行目の奥にしか無かった。
        在るかどうかではなく、**探す人の目の位置に在るか**の問題だった。
    ★ ここが縛るのは「先頭に在ること」と「月末の紙を名指ししていること」だけ
      （文面は人が決める）。
    """
    lines = README.splitlines()
    where = [i for i, ln in enumerate(lines) if ln.startswith("## ") and "どこを読む" in ln]
    assert where, "先頭の道案内が消えている"
    assert where[0] < 40, f"道案内が {where[0]} 行目にある ── 買い手は上から読む"
    router = chr(10).join(lines[where[0]:where[0] + 20])
    assert "月末の締めのやり方" in router, "月末に何を打つかの紙を案内していない"
    assert "セットアップ" in router, "動かし方への行き先が無い"
