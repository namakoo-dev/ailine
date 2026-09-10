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
#   ② README が「まずこれを打つ」と書いたコマンドは、そのタグに**実在する**
#      ── または、その場で版の要求を**開示している**（断れない時は開示する）。
#      実在するようになったら但し書きの残骸を同じ試験が赤にする（両方向に噛む）。

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


def _version_in_pyproject():
    """pyproject の version を "vX.Y.Z" で返す（★ 版はここが正）。"""
    try:
        text = (REPO / "pyproject.toml").read_text(encoding="utf-8")
    except OSError:
        return None
    m = re.search(r'^version\s*=\s*"([0-9][0-9.]*)"', text, re.M)
    return "v" + m.group(1) if m else None


def _readme_is_current(pinned, newest, next_ver):
    """★ 判定はこの 1 本だけ。試験も本体もここを通る（分岐を書き写さない）。"""
    if pinned == [newest]:
        return True
    return next_ver is not None and pinned == [next_ver]


def test_readme_points_at_the_newest_tag():
    """① 版を上げたら README も上がる ── 人の記憶でなく機械で縛る。

    ★★ 2026-09-11: この番人は**リリースのたびに必ず一度赤くなる**構造だった。

        先に README を上げる → そのタグはまだ存在しない          → 赤
        後で README を上げる → タグが先に出て README が古くなる  → 赤

      どちらの順序でも通れない。実際 2 日続けて両方の踏み方をした
      （9-10 は後者・9-11 は前者）。★「手順が無い」のではなく、
      **番人が要求する順序が存在しなかった。**

    ★ 抜け道は**最小**にする ── 「これから出す版（pyproject の version）と一致」
      している時だけ通す。tagpr のリリースブランチからは誰も install しないので、
      元の実害（README どおりに入れたら最初の 1 コマンドで詰まる）は起きない。
    ★ 緩めが本物の事故を素通りさせないことは、下の負の被覆の試験が縛る。
    """
    tags = _tags()
    if not tags:
        pytest.skip("タグがまだ無い（測れない回は skip と書く）")
    newest = tags[0]
    text = README.read_text(encoding="utf-8")
    pinned = sorted(set(re.findall(r"ailine@(v[0-9][0-9.]*)", text)))
    assert pinned, "README に入れ方のタグ指定が無い"
    next_ver = _version_in_pyproject()
    assert _readme_is_current(pinned, newest, next_ver), (
        f"README が案内するタグ {pinned} が、最新タグ {newest} とも "
        f"これから出す版 {next_ver} とも違う ── "
        "この案内で入れた人は、README と違う版を使うことになる")


def test_the_escape_hatch_does_not_let_a_stale_readme_through():
    """★ 負の被覆 ── 緩めた穴が、本物の事故を通さないこと。

    実害の形は「README が**古い**版を指す」（初回体験の盲検で @v0.1.0 で入れて
    `ailine demo` が無く詰まった、あの事故）。抜け道は「これから出す版」だけを許す。
    ★ 本体と同じ _readme_is_current を通す（判定を書き写すと恒真になる）。
    """
    newest, next_ver = "v0.2.4", "v0.2.5"
    assert _readme_is_current([newest], newest, next_ver), "最新タグ → 通るはず"
    assert _readme_is_current([next_ver], newest, next_ver), "これから出す版 → 通るはず"
    assert not _readme_is_current(["v0.1.0"], newest, next_ver), \
        "★ 古い版を指す README が通った ── 抜け道が広すぎる"
    assert not _readme_is_current(["v99.0.0"], newest, next_ver), \
        "★ 存在しない未来の版が通った"
    assert not _readme_is_current([newest, next_ver], newest, next_ver), \
        "★ 2 つ書いてあるのに通った（案内は 1 つのはず）"
    assert not _readme_is_current([next_ver], newest, None), \
        "★ pyproject が読めない時に通った（抜け道は版が確定している時だけ）"


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
    # ★ 断れない時は開示する: 版を上げるのは出荷の判断で、README だけでは直せない。
    #   だから契約は「実在する」**または**「その場で版の要求を開示している」の二択にする。
    #   ── ただし**両方向に噛む**: 実在するようになったら、古い但し書きが残っていることを
    #   同じ試験が赤にする（開示は消し忘れると嘘になる）。
    nearby = chr(10).join(lines[i:i + 14])
    discloses = ("v0.1.2" in nearby or "以降" in nearby) and cmd in nearby
    if cmd in registered:
        assert not discloses, (
            f"`ailine {cmd}` は案内するタグ {newest} に実在するのに、README に版の但し書きが"
            "残っている（開示は消し忘れると嘘になる）")
    else:
        assert discloses, (
            f"README が最初に打てと言う `ailine {cmd}` が、案内するタグ {newest} に無く、"
            f"その場に版の但し書きも無い（{newest} のサブコマンド: "
            f"{' '.join(sorted(set(registered)))}）")
