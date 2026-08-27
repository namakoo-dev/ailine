# 文書から張ったローカルへのリンクが、実在する先を指しているかの番人（2026-08-27）。
#
# ★ なぜ在るか: README を docs/ENGINEERING.md へ割ったこの日、文書が 2 つになった。
#   2 つになった瞬間から、片方を動かすともう片方のリンクが静かに死ぬ。
#   ★ 「消えたものは差分に出ない」── リンク切れは、消したファイルの側の diff にしか
#     現れない。**張った側から**確かめる必要がある。
#
# 対象: repo 内の .md が張る相対リンク（http/https/mailto/アンカーだけのものは対象外）。
import re
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
LINK = re.compile(r"\[[^\]]*\]\(([^)\s]+)\)")
SKIP_PREFIX = ("http://", "https://", "mailto:", "#")


def _docs():
    for pattern in ("*.md", "docs/**/*.md"):
        for f in REPO.glob(pattern):
            if ".git" in f.parts:
                continue
            yield f


def test_relative_links_point_at_something_that_exists():
    broken = []
    for f in _docs():
        text = f.read_text(encoding="utf-8")
        for target in LINK.findall(text):
            if target.startswith(SKIP_PREFIX):
                continue
            path = (f.parent / target.split("#", 1)[0]).resolve()
            if not path.exists():
                broken.append(f"{f.relative_to(REPO).as_posix()} → {target}")
    assert not broken, (
        "文書のリンクが実在しない先を指している（文書を動かした時に静かに死ぬ形）: "
        + " / ".join(broken))


def test_the_two_documents_point_at_each_other():
    """★ 割った 2 つは、両方向に繋がっていること ── 片道だと、片方が孤児になる。"""
    readme = (REPO / "README.md").read_text(encoding="utf-8")
    eng = (REPO / "docs" / "ENGINEERING.md").read_text(encoding="utf-8")
    assert "docs/ENGINEERING.md" in readme, "README から技術詳細への道が無い"
    assert "../README.md" in eng, "docs/ENGINEERING.md から README への戻り道が無い"
