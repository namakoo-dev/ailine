# 製品コードを**分割の場所に依らず**読むための共通の芯（2026-09-03）。
#
# ★★ なぜ在るか: この repo の番人には「実装のソースを読んで契約を確かめる」型が
#   いくつもある（例:「この文言は 1 箇所でしか組み立てていないこと」）。
#   契約そのものは正しいのに、**読む場所が `src/ailine/__init__.py` 決め打ち**だった。
#   2026-09-03 に事後条件 45 関数を ailine_core/postconditions/ へ移したところ、
#   **7 件の番人が同時に空振りした** ── 実装は 2 箇所に分かれたのに、番人は 1 箇所しか
#   見ていない。★ この repo が「片配線」と呼んできた形を、**番人自身がやっていた**。
#
# ★ 処方は系譜どおり ── 両方直すのではなく **1 関数に畳んで呼び出し側に持たせない**。
#   ここを通れば、次にどこへ分割しても番人は空振りしない。
#
# ★ 使い分け:
#   count_in_product(needle)      … 製品コード全体での出現回数（「1 箇所だけ」の検査）
#   window_around(anchor, lines)  … anchor を含む**そのファイル**の窓（前後を読む検査）
#                                   ★ 窓はファイルをまたがない ── またぐと、隣の
#                                     ファイルの文字列を「近くにある」と誤読する

from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src"


def product_files() -> list:
    """製品コードの .py を全部（本体 + ailine_core の全モジュール）。"""
    files = [SRC / "ailine" / "__init__.py"]
    files += sorted(p for p in (SRC / "ailine_core").rglob("*.py"))
    return files


def count_in_product(needle: str) -> int:
    """製品コード全体での出現回数。★ 「1 箇所でしか書いていない」の検査に使う。"""
    return sum(p.read_bytes().decode("utf-8").count(needle) for p in product_files())


def files_containing(needle: str) -> list:
    """needle を含むファイル（どこに移ったかを見せる。赤くなった時に読みやすい）。"""
    return [p for p in product_files() if needle in p.read_bytes().decode("utf-8")]


def window_around(anchor: str, after: int = 4000, before: int = 0) -> str:
    """anchor を含むファイルから、その前後を切り出す（★ ファイルをまたがない）。

    anchor がどのファイルにも無ければ AssertionError。2 つ以上のファイルに
    あれば、それ自体が「1 箇所のはず」の違反なので AssertionError にする。
    """
    hits = files_containing(anchor)
    assert hits, f"anchor がどの製品ファイルにも無い: {anchor!r}"
    assert len(hits) == 1, (
        f"anchor が {len(hits)} 個のファイルに在る（窓が切れない）: "
        f"{[p.name for p in hits]} ── {anchor!r}")
    text = hits[0].read_bytes().decode("utf-8")
    i = text.index(anchor)
    return text[max(0, i - before): i + after]
