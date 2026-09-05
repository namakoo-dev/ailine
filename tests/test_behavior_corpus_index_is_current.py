"""挙動コーパスの索引が、実体とずれていないこと（2026-09-05）。

★★ 出所（盲検の査定・所見⑧）: `docs/behavior-corpus/nodes/` に 19 本あるのに、
  索引は 18 本しか載せていなかった。欠けていたのは
  `arithmetic-identity-check.md` ── **HEAD 直下の commit で追加されたばかりのもの**。

★ いちばん痛いのは、同じ `MEMORY.md` が既にこう書いていたことだ:

    ★ 件数は書かない: 旧文は「上記 9 ノード」と書いていたが、索引が伸びても追随せず
      実数と食い違っていた（実測: 索引 16 件に対し 9 のまま）

  **教訓は書かれ、件数は消され、索引そのものは手書きのまま残り、最後の commit で再発した。**
  査定者の言葉:「正しい教訓を言語化する能力は高いが、自分の運用に落とし切る所で毎回
  一段足りない」── ここはその指摘に対する処置そのもの。

★ 「件数を書かない」は正しかったが、足りなかった。**索引そのものを機械が突き合わせる。**
"""
import re
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
CORPUS = REPO / "docs" / "behavior-corpus"
INDEX = CORPUS / "MEMORY.md"
NODES = CORPUS / "nodes"

LINK_RE = re.compile(r"nodes/([a-z0-9-]+\.md)")


def _listed() -> set:
    return set(LINK_RE.findall(INDEX.read_text(encoding="utf-8")))


def _present() -> set:
    return {p.name for p in NODES.glob("*.md")}


def test_the_corpus_is_not_empty():
    """★ 下限 ── 検出が壊れて 0 件になれば、以下は恒真になる。"""
    assert _present(), "ノードが 1 本も無い（検出が壊れている疑い）"
    assert _listed(), "索引が 1 本も指していない（同上）"


def test_every_node_is_in_the_index():
    """★ 本命: 書いたのに索引に出ない（＝誰にも読まれない）を防ぐ。"""
    missing = sorted(_present() - _listed())
    assert not missing, (
        f"索引に載っていないノード: {missing} ── {INDEX.relative_to(REPO)} に 1 行足すこと")


def test_the_index_does_not_point_at_ghosts():
    """★ 逆向き: 消したノードを索引が指し続けない（リンク切れ）。"""
    ghosts = sorted(_listed() - _present())
    assert not ghosts, f"実体の無いノードを索引が指している: {ghosts}"


def test_the_index_still_declares_no_count():
    """★ 先に学んだ教訓が消えていないこと（件数を書けば必ず腐る）。"""
    text = INDEX.read_text(encoding="utf-8")
    assert "件数は書かない" in text, "「件数は書かない」の戒めが索引から消えている"
    # ★ 鉤括弧の中は**引用**（戒めの本文が「上記 9 ノード」を悪い例として引いている）。
    #   ここを数えると、番人が自分の教訓文で鳴く ── 実際に一度鳴いた。
    body = re.sub("「[^」]*」", "", LINK_RE.sub("", text))
    stale = re.findall(r"(?:上記|全部で|計)\s*\d+\s*ノード", body)
    assert not stale, f"索引に件数が書かれている（伸びたら腐る）: {stale}"
