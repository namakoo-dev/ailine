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
    # ★ 2026-08-27: docs/ を足した。README を docs/ENGINEERING.md へ移した瞬間、
    #   この番人の走査から外れていた ── 「在るのに、その事故の形では鳴らない」。
    for pattern in ("*.md", "docs/**/*.md", "gui/*.html", "src/**/*.py",
                     "src/**/*.bas", "tests/**/*.py", "bench/**/*.py"):
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


CRLF_BYTES = (chr(13) + chr(10)).encode("ascii")   # ★ ヒアドキュメントで消える escape を避ける


def test_line_endings_are_not_rewritten():
    """★ 2026-08-24（2 度目）: 初版は **LF→CRLF の一方向**しか見ておらず、
       同じ日に自分で **CRLF→LF** を踏んで 176 行の変更が **11,334 行**に膨れた。
       番人に変異試験が足りていなかった ── 両方向を見る。
       縛るのは「変えたこと」だけ（repo 全体の統一はしない）。"""
    import subprocess
    # ★ 2026-08-27: 段階済みの変更も見る（git mv はその場で段階されるため、
    #   --name-only だけだと改名してから改行を潰した形がすり抜ける）。
    changed = subprocess.run(["git", "diff", "--name-only"],
                              cwd=REPO, capture_output=True, text=True).stdout.split()
    changed += subprocess.run(["git", "diff", "--cached", "--name-only"],
                               cwd=REPO, capture_output=True, text=True).stdout.split()
    offenders = []
    for name in sorted(set(changed)):
        f = REPO / name
        if not f.exists():
            continue
        # ★ 改名した直後は HEAD にその名前が無い ── 空を 0 件と数えると
        #   「改行が変わった」と誤報する。段階（index）を下がりの基準にする。
        head = None
        for ref in ("HEAD:" + name, ":" + name):
            r = subprocess.run(["git", "show", ref], cwd=REPO, capture_output=True)
            if r.returncode == 0:
                head = r.stdout
                break
        if head is None:
            continue          # git がまだ知らない新規ファイル（比べる先が無い）
        head_crlf = head.count(CRLF_BYTES)
        now_crlf = f.read_bytes().count(CRLF_BYTES)
        if (head_crlf == 0) != (now_crlf == 0):
            offenders.append(f"{name}(HEAD {head_crlf} / 現 {now_crlf})")
    assert not offenders, (
        "改行コードが git の中身と食い違っている（write_text が書き換えた痕・"
        "中身の差分が改行の海に沈む）: " + " / ".join(offenders))


def test_the_eol_guardian_sees_both_directions():
    """★ 変異試験（両向き）: LF→CRLF と CRLF→LF の**どちらも**検出できること。"""
    lf = ("a" + chr(10) + "b" + chr(10)).encode()
    crlf = ("a" + chr(13) + chr(10) + "b" + chr(13) + chr(10)).encode()

    def differs(head: bytes, now: bytes) -> bool:
        return (head.count(CRLF_BYTES) == 0) != (now.count(CRLF_BYTES) == 0)

    assert differs(lf, crlf), "LF→CRLF が見えない"
    assert differs(crlf, lf), "CRLF→LF が見えない（初版が見逃した向き）"
    assert not differs(lf, lf) and not differs(crlf, crlf), "同じなのに鳴る（誤爆）"
