# 制御文字の混入を止める番人（2026-08-24）。
#
# ★ なぜ在るか: 盲検の査定者が「導入で詰む」として名指しした README の設置コマンドが
#   `setx BASRUN "%CD%\basrun\basrun.py"` ではなく `%CD%^Hasrun^Hasrun.py` になっていた ──
#   `\b` が**バックスペース(0x08)として焼き込まれていた**。目で見ても分からない
#   （端末では前の文字が消えて見えるだけ）。src/ailine/__init__.py のコメントにも 1 箇所あった。
#   ★ 記憶の「異文字混入と CRLF 巻き込み」の再来。1 箇所直しても、書き方が同じなら再発する。
#   だから**探し方を機械にする**（明示コードポイント範囲で全ファイルを走査）。
#
# 対象外にする文字: 改行(0x0a)・復帰(0x0d)・タブ(0x09)。それ以外の C0 制御文字は
# 文書にもコードにも現れる理由が無い。

import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _targets():
    for pattern in ("*.md", "src/**/*.py", "src/**/*.bas", "tests/**/*.py", "bench/**/*.py"):
        for f in REPO.glob(pattern):
            if ".git" in f.parts or "__pycache__" in f.parts:
                continue
            yield f


def test_no_stray_control_characters():
    found = []
    for f in _targets():
        try:
            text = f.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        for m in CONTROL.finditer(text):
            line = text[:m.start()].count("\n") + 1
            found.append(f"{f.relative_to(REPO)}:{line} に {hex(ord(m.group(0)))}")
    assert not found, (
        "制御文字が混入している（ヒアドキュメントの \b \t \a 等が展開された痕）: "
        + " / ".join(found))


def test_the_guardian_can_actually_see_one(tmp_path):
    """★ 変異試験: 番人が本当に見えるかを確かめる（恒真の番人を置かない）。"""
    p = tmp_path / "x.md"
    p.write_text("path C:\x08asrun", encoding="utf-8")
    assert CONTROL.search(p.read_text(encoding="utf-8")), "番人の目が節穴"


# --- 改行コードの巻き込みを止める番人（2026-08-24）------------------------------------
#
# ★ なぜ在るか: 制御文字を 4 文字直しただけの commit で `git diff --stat` が
#   **1286 行変更**になっていた ── Windows の `write_text` が LF を CRLF に変えていた。
#   中身の変更が改行の海に沈むと、人がレビューできず、ゴールデンも壊れる。
#   ★ 記憶の「異文字混入と CRLF 巻き込み」の処方（diff --stat で規模照合）が効いた事例。
#
# ★ 縛り方: 「repo 全体を LF に統一」ではない（既に CRLF で入っている 33 ファイルを
#   一斉に書き換えると、それ自体が巨大な差分になる）。縛るのは**変えたこと**だけ ──
#   git に LF で入っているファイルが、作業ツリーで CRLF になっていたら赤。
#   きれいな checkout では自明に緑・手元の事故だけを捕まえる。


def test_line_endings_are_not_rewritten():
    import subprocess
    changed = subprocess.run(["git", "diff", "--name-only"],
                              cwd=REPO, capture_output=True, text=True).stdout.split()
    offenders = []
    for name in changed:
        f = REPO / name
        if not f.exists():
            continue
        head = subprocess.run(["git", "show", f"HEAD:{name}"],
                               cwd=REPO, capture_output=True).stdout
        if b"\r\n" not in head and b"\r\n" in f.read_bytes():
            offenders.append(name)
    assert not offenders, (
        "git には LF で入っているのに作業ツリーが CRLF になっている"
        "（Windows の write_text が改行を書き換えた痕・中身の差分が沈む）: "
        + " / ".join(offenders))
