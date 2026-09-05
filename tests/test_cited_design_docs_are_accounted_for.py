"""コードが名指しした設計文書は、**在る**か「無い」と書いてあるかのどちらかであること。

★★ 出所（2026-09-05・盲検の査定・所見②）: コードとテストのコメントが、判断の根拠として
  9 本の設計文書を **58 箇所**で名指ししていた。**1 本も repo に無い。**
  「これは非公開です」と断る一文も、どこにも無かった。

  査定者の言葉:「この repo の最大の武器は『判断の跡が追える』こと。追おうとした
  評価者が 55 回行き止まりに当たる。**武器が一番効くべき瞬間に空振りする。**」

★ 構造の穴（この repo の系譜そのもの）: `tests/test_doc_links_resolve.py` は
  **Markdown のリンク切れ**を機械で守っていた。同じ契約が**コードのコメント側には
  配線されていなかった** ── 片配線。ここがその欠けていた側を塞ぐ。

★ 処置は「文書を出す」ではない（取引先名と他社の帳票を含んでいたので公開しなかった・
  もう残っていない）。**出せないと書く**ことと、**次に同じことが起きたら赤くする**こと。
"""
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

#: 名指しの形（日付つきの設計・レビュー・封）
CITE_RE = re.compile(r"\b((?:DESIGN|REVIEW|SEALED)-\d{8}-[A-Za-z0-9-]+\.md)\b")

#: この表を持つ文書（無い文書を「無い」と認めている場所）
LEDGER = REPO / "docs" / "ENGINEERING.md"

#: 走査する場所（生成物と .git は見ない）
SCAN_DIRS = ("src", "tests", "scripts", "gui", "docs")
SKIP_PARTS = {"__pycache__", ".git", ".pytest_cache", "node_modules"}


def _cited() -> dict:
    """{文書名: [引用元の相対パス, ...]}"""
    found = {}
    for d in SCAN_DIRS:
        for p in (REPO / d).rglob("*"):
            if p.is_dir() or set(p.parts) & SKIP_PARTS:
                continue
            if p.suffix not in (".py", ".md", ".txt", ".yml", ".yaml"):
                continue
            if p == LEDGER:
                continue          # ★ 名簿そのものは引用元に数えない
            try:
                text = p.read_text(encoding="utf-8")
            except (UnicodeDecodeError, OSError):
                continue
            for name in set(CITE_RE.findall(text)):
                found.setdefault(name, []).append(str(p.relative_to(REPO)))
    return found


def test_the_scan_finds_something():
    """★ 下限 ── 検出が壊れて 0 件になったら、以下は全部恒真になる。"""
    assert _cited(), "設計文書の引用が 1 件も取れていない（検出が壊れている疑い）"


def test_every_cited_document_exists_or_is_declared_missing():
    """★ 本命: 名指しした文書は、repo に在るか、名簿に「無い」と書いてあること。"""
    ledger = LEDGER.read_text(encoding="utf-8")
    unaccounted = {}
    for name, citers in sorted(_cited().items()):
        if (REPO / "docs" / name).exists() or list(REPO.rglob(name)):
            continue                       # 在る
        if name in ledger:
            continue                       # 「無い」と書いてある
        unaccounted[name] = citers[:3]
    assert not unaccounted, (
        f"repo に無く、名簿にも載っていない設計文書を名指ししている: {unaccounted}"
        f" ── 読む側は行き止まりに当たる。{LEDGER.name} の表に足すか、文書を置くこと")


def test_the_ledger_does_not_list_documents_nobody_cites():
    """★ 逆向き ── 誰も引用していない名前が名簿に残り続けないこと（腐る側）。"""
    ledger = LEDGER.read_text(encoding="utf-8")
    listed = set(CITE_RE.findall(ledger))
    cited = set(_cited())
    stale = sorted(n for n in listed - cited if not list(REPO.rglob(n)))
    assert not stale, (
        f"名簿に在るが、もうどこからも引用されていない: {stale} ── 名簿から外すこと")
