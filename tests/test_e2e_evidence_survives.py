# 実測の証跡が消えていないこと（2026-09-04）。
#
# ★★ なぜ在るか: 外部の査定が `e2e_work/` を「どこからも参照されていない残骸（grep 0 件）」
#   と読んだ。★ **grep が届いていなかった** ── 実際は製品コード・テスト・文書の 5 箇所が
#   名指しで参照している。消していたら「ここで実測した」という主張の裏付けが消えていた。
#
# ★ 教訓は 2 つ:
#   ① **外の目も測定器を持っていて、それも疑う対象**（この repo が自分に課しているのと同じ）
#   ② ★ 参照が人の目にしか見えない状態だったのが悪い。**機械が守れば、次は誰も消さない**
#
# 契約:
#   ① コード・テスト・文書が名指しした証跡ファイルが、実在すること
#   ② 入力ブックとログが対で残っていること（片方だけでは再現できない）
#   ③ ★ 説明（e2e_work/README.md）が在ること ── 説明の無い証跡は残骸に見える

import re
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
E2E = REPO / "e2e_work"

# 名指しで参照している側（★ ここが増えたら足す。減ったらこの試験も見直す）
CITERS = [
    "src/ailine/__init__.py",
    "tests/test_ailine.py",
    "bench/realworld/BASELINE.md",
    "docs/behavior-corpus/nodes/verification-scope-honesty.md",
]
_REF = re.compile(r"e2e_work/([A-Za-z0-9_/]+\.(?:txt|xlsx|log))")


def _cited() -> set:
    out = set()
    for rel in CITERS:
        p = REPO / rel
        if not p.exists():
            continue
        out |= set(_REF.findall(p.read_bytes().decode("utf-8", errors="replace")))
    return out


def test_every_cited_evidence_file_exists():
    """① 名指しされた証跡が実在すること。

    ★ ここが赤くなったら、消したのは「残骸」ではなく**主張の裏付け**。
      コード側のコメントやテストが「実測した」と言っている根拠が消えている。
    """
    cited = _cited()
    assert cited, "参照が 1 件も取れていない（★ 検出が壊れている疑い ── 下限）"
    missing = sorted(n for n in cited if not (E2E / n).exists())
    assert not missing, (
        f"名指しされた実測の証跡が消えている: {missing} ── "
        f"参照元: {CITERS}。消す前に、参照している側の主張をどうするか決めること")


def test_inputs_and_logs_stay_paired():
    """② ログが在る所には、入力ブックも残っていること。

    ★ 片方だけ残すと再現できない ── ログだけでは何を入れたか分からず、
      ブックだけでは何が起きたか分からない。

    ★★ この試験は**落ちない試験だった**（2026-09-05・盲検の査定が指摘）:
      `logs` と `books` を作って**一度も比べていなかった**。ブックを全部消しても緑。
      docstring が「対で残っていること」と宣言した検査が、body に存在しなかった。

    ★ 直すときに分かったこと: **名前で 1 対 1 に対応させる規則は、このデータに無い**
      （ログ `e2e1_log.txt` に対して入力は `sample_e2e.xlsx`）。
      `logs == books` と書けば「守っているように見えて常に赤」になるだけで、
      それは検査ではなく飾りだ。★ 実際に真で、かつ**消したら赤くなる**規則を選ぶ:
      **ログを置いたディレクトリには、必ずブックが同居していること**。
      （逆向きは成り立たない ── `.ailine_*` は道具の作業場で、証跡ではない）
    """
    logs = sorted(E2E.rglob("*_log.txt"))
    books = sorted(E2E.rglob("*.xlsx"))
    assert logs, "ログが 1 つも無い（★ 検出が壊れている疑い）"
    assert books, "入力ブックが 1 つも無い（★ 同上）"
    orphaned = sorted(str(d.relative_to(E2E)) for d in {p.parent for p in logs}
                      if not any(d.glob("*.xlsx")))
    assert not orphaned, (
        f"ログだけが残ってブックが消えたディレクトリ: {orphaned} ── "
        "何を入れたら その出力になったのかが再現できない")


def test_the_evidence_is_explained():
    """③ ★ 説明が在ること ── 説明の無い証跡は、外の目には残骸に見える。

    ★ 実際そう読まれた（2026-09-04 の査定）。中身が正しくても、
      **なぜそこに在るかが書かれていなければ、片付いていないと判断される。**
    """
    doc = E2E / "README.md"
    assert doc.exists(), "e2e_work/README.md が無い（説明の無い証跡は残骸に見える）"
    t = doc.read_bytes().decode("utf-8")
    assert "消さない" in t and "参照" in t, "何のために在るかが書かれていない"
    for rel in CITERS:
        assert Path(rel).name in t or rel in t, f"参照元 {rel} が説明に載っていない"
