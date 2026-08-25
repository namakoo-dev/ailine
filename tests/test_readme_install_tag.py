# 初回体験の盲検 2 回目・所見1（2026-08-25）── README が案内する版と、README の中身がずれる。
#
# ★ 実測: 検分者が README どおり `uv tool install ...@v0.1.0` で入れ、README の
#   「まずこれを打つ」に従って `ailine demo` を叩いたら:
#       ailine: error: argument cmd: invalid choice: 'demo'
#   **最初の 1 コマンドで詰まった。** README は新しい版の内容を書きながら、
#   README 自身が案内する入れ方は古い版を固定していた。
#   ★ しかも `ailine doctor` は「✓ demo/（サンプルがあります）」と言い続ける ──
#     「あるはずのものに手が届かない」という一番わかりにくい壊れ方。
#
# ★ 根: 「タグを指定して固定する」という**正しい設計**が、版を上げるたびに README を
#   直す義務を生むのに、それを守る仕組みが無かった。人の記憶に頼っていた。
#
# 契約:
#   ① README が案内するタグは、この repo の**最新タグ**と一致する
#   ② README が「まずこれを打つ」と書いたコマンドは、そのタグに実在する

import re
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
README = REPO / "README.md"


def _tags():
    r = subprocess.run(["git", "tag", "--sort=-version:refname"],
                        cwd=str(REPO), capture_output=True, text=True)
    return [t for t in (r.stdout or "").splitlines() if t.strip()]


def test_readme_points_at_the_newest_tag():
    """① 版を上げたら README も上がる ── 人の記憶でなく機械で縛る。"""
    tags = _tags()
    if not tags:
        pytest.skip("タグがまだ無い（測れない回は skip と書く）")
    newest = tags[0]
    text = README.read_text(encoding="utf-8")
    pinned = sorted(set(re.findall(r"ailine@(v[0-9][0-9.]*)", text)))
    assert pinned, "README に入れ方のタグ指定が無い"
    assert pinned == [newest], (
        f"README が案内するタグ {pinned} が最新タグ {newest} と違う ── "
        "この案内で入れた人は、README と違う版を使うことになる")


def test_the_first_command_exists_in_that_tag():
    """② 「まずこれを打つ」が、その版に実在すること。

    ★ 実測した壊れ方そのもの: README は demo を勧め、案内するタグには demo が無かった。
    """
    tags = _tags()
    if not tags:
        pytest.skip("タグがまだ無い")
    newest = tags[0]
    r = subprocess.run(["git", "show", f"{newest}:src/ailine/__init__.py"],
                        cwd=str(REPO), capture_output=True, text=True,
                        encoding="utf-8", errors="replace")
    if r.returncode != 0:
        pytest.skip(f"{newest} の本体が読めない（レイアウトが違う版）")
    tagged_src = r.stdout
    text = README.read_text(encoding="utf-8")
    # ★ 治具の訂正: 「まずこれを打つ」の直後は**コードブロック**で、バッククォート
    #   囲みではない（初版はそれを探して skip していた ── skip は「守っている」ではない）。
    lines = text.split(chr(10))
    i = next((i for i, l in enumerate(lines) if "まずこれを打つ" in l), None)
    assert i is not None, "README に「まずこれを打つ」の節が無い"
    cmd = None
    for l in lines[i:i + 8]:
        m2 = re.match(r"\s*ailine\s+([\w-]+)", l)
        if m2:
            cmd = m2.group(1)
            break
    assert cmd, f"「まずこれを打つ」の直後にコマンドが見つからない: {lines[i:i + 8]}"
    # ★ 番人が恒真だった（実測）: 素朴に `"demo"` を探すと `HERE / "demo"`（フォルダ名）に
    #   当たり、サブコマンドが無いのに「在る」と判定していた。
    #   ★ **サブコマンドの登録**で見る ── 探すものを、意味の在る形で書く。
    registered = re.findall(r'sub\.add_parser\("([a-z][a-z-]*)"', tagged_src)
    assert cmd in registered, (
        f"README が最初に打てと言う `ailine {cmd}` が、案内するタグ {newest} に無い"
        f"（{newest} のサブコマンド: {' '.join(sorted(set(registered)))}）")
