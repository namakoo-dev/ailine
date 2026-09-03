# 改行の**台帳**（2026-09-03）── 1 つのファイルの中で改行が混ざっていないこと。
#
# ★★ なぜ在るか: この日、改行を **2 回**壊した。
#   ① `src/ailine/__init__.py`（CRLF）を一括編集して LF に潰した。原因は書く側でなく
#      **読む側** ── `Path.read_text()` はユニバーサル改行モードなので、読んだ時点で
#      CRLF が LF に潰れている。`write_bytes` で書いても手遅れだった。
#      気づいたのは `git diff --stat` が 35,339 行の変更を出したから（★ 人の目視）。
#   ② 記憶ファイルに LF の断片を追記して混在を作った（脳側・repo 外）。
#
# ★ そして数えたら、**index に混在が 15 件**あった ── 今日作ったものではなく、
#   前から入っていた。tests/*.py の 7 本は過去に LF で追記した事故の痕跡で、
#   CRLF に統一した。★ **後から見つけるのは一苦労**（Namakoo 2026-09-03）だから、
#   見つけた時点で機械に持たせる。
#
# ★ index の側だけを見る理由: 作業ツリーの改行は checkout の設定（core.autocrlf）で
#   変わる。GitHub の Windows runner は autocrlf=true なので、作業ツリーを見る契約は
#   CI でだけ赤くなる（この repo は 2026-08-21 に golden で同じ罠を踏み、
#   .gitattributes の `tests/golden/** -text` で塞いだ）。**index は環境に依らない。**
#
# 契約:
#   ① 1 ファイルの中で改行が混ざっていないこと（既知の例外だけ許す）
#   ② 例外は理由つきで台帳に在ること
#   ③ もう混ざっていないファイルが台帳に残っていたら赤（古い不安を配らない）

import subprocess
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent

# --- 混ざっていてよいもの（★ 理由を書く。無い行は赤になる）-----------------------------
ALLOWED_MIXED = {
    "bench/realworld/logs/": "実行結果の記録。本文は走らせた道具の出力そのままで、"
                              "最後の 1 行だけ記録スクリプトが `EXIT:n` を LF で書き足す。"
                              "★ 混在はこの構造そのものなので直さない（直すと記録が"
                              "『そのまま』でなくなる）。",
    "bench/realworld/logs_fix/": "同じ記録の、直したあとに走らせ直したぶん。本文は道具の出力そのままで、最後の 1 行だけ `EXIT:n` が LF で付く。",
}


def _index_eol():
    """(改行の種別, パス) の一覧を git の index から取る。"""
    r = subprocess.run(["git", "ls-files", "--eol"], cwd=REPO,
                        capture_output=True, text=True)
    if r.returncode != 0:
        pytest.skip("git が使えない環境（sdist からの実行など）")
    out = []
    for line in r.stdout.splitlines():
        parts = line.split("\t")
        if len(parts) != 2:
            continue
        kind = parts[0].split()[0]          # 例: "i/mixed"
        out.append((kind, parts[1].strip()))
    return out


def _mixed():
    return sorted(p for kind, p in _index_eol() if kind == "i/mixed")


def _is_allowed(path: str) -> bool:
    return any(path.startswith(prefix) for prefix in ALLOWED_MIXED)


def test_no_file_mixes_line_endings():
    """① 1 ファイルの中で改行が混ざらないこと。

    ★ 直し方: そのファイルの多数派に揃える。読む時は必ず `read_bytes()` を使い、
      `read_text()` は使わない（ユニバーサル改行モードが入口で潰す）。
    """
    bad = [p for p in _mixed() if not _is_allowed(p)]
    assert not bad, (
        f"改行が混ざっているファイル: {bad} ── "
        "多数派に揃えるか、混ざってよい理由を ALLOWED_MIXED に書くこと")


def test_the_allowed_list_is_not_stale():
    """③ もう混ざっていない例外が残っていたら赤（古い不安を配らない）。"""
    mixed = _mixed()
    unused = [prefix for prefix in ALLOWED_MIXED
              if not any(p.startswith(prefix) for p in mixed)]
    assert not unused, f"もう混在していないのに台帳に残っている: {unused}"


def test_every_exception_states_a_reason():
    """② 例外を黙って増やせない ── 理由が文になっていること。"""
    for prefix, why in ALLOWED_MIXED.items():
        assert len(why) >= 20 and "。" in why, f"ALLOWED_MIXED[{prefix}] の理由が薄い"


def test_the_guard_can_see_something():
    """★ 陽性対照 ── index を読めていること（空集合どうしの比較を通さない）。

    ★ この repo は同じ日に「読めていることを先に確かめずに空集合を比べる」形を
      2 度踏んでいる。分母が取れているかを先に見る。
    """
    rows = _index_eol()
    assert len(rows) >= 300, f"index から {len(rows)} 件しか読めていない"
    kinds = {k for k, _ in rows}
    assert {"i/lf", "i/crlf"} <= kinds, f"改行の種別が取れていない: {kinds}"
