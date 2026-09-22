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
    """製品コードの .py を全部（本体 ＋ ailine_core ＋ ★ GUI）。

    ★ GUI を含める理由（2026-09-04）: 初版は src/ だけを見ていたため、
      `gui/server.py` にしか無い文字列を数えると **0 件**になり、
      「1 箇所だけ」の契約が「どこにも無い」と誤判定された。
      ★ 画面も製品の一部 ── 出荷されるコードは src/ だけではない。
    """
    files = [SRC / "ailine" / "__init__.py"]
    files += sorted(p for p in (SRC / "ailine_core").rglob("*.py"))
    files += sorted((REPO / "gui").glob("*.py"))
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

def branch_source(test_snippet: str) -> str:
    """`if <test_snippet ...>:` の**枝の中身**を丸ごと返す（AST で切る）。

    ★★ なぜ在るか（2026-09-20）: 番人が `window_around(anchor, after=1800)` のように
      **バイト距離**で窓を切っていると、その枝にコメントを 14 行足しただけで
      契約が窓から押し出されて赤くなる。守っている不変は 1 文字も変わっていない。
      ★ 今週 4 件目の「番人を字面（や距離）で書いた」事故 ── 関数名・印の字面・
        `-m local` の字面に続いて、今度は**窓の広さ**だった。
    ★ 枝は構文の単位なので、中身が増えても縮んでも同じものを指し続ける。

    ★ 見つからない／2 つ以上あるときは AssertionError（空回りさせない）。
    """
    import ast
    hits, exact = [], []
    for path in product_files():
        src = path.read_bytes().decode("utf-8")
        try:
            tree = ast.parse(src)
        except SyntaxError:          # ★ 製品でない断片は読み飛ばす（HTML 混じり等）
            continue
        for node in ast.walk(tree):
            if not isinstance(node, ast.If):
                continue
            cond = (ast.get_source_segment(src, node.test) or "").strip()
            if test_snippet not in cond:
                continue
            body = chr(10).join(
                ast.get_source_segment(src, st) or "" for st in node.body)
            hits.append(body)
            if cond == test_snippet.strip():
                exact.append(body)
    # ★ 完全一致が在ればそれを採る ── `forced_op` のような短い条件は、それを含む
    #   別の条件（`forced_op not in OP_SCHEMA` 等）にも当たってしまう。
    hits = exact or hits
    assert hits, f"その条件の枝が製品に無い: {test_snippet!r}"
    assert len(hits) == 1, f"枝が {len(hits)} 箇所ある（1 つに絞れない）: {test_snippet!r}"
    return hits[0]


def product_text() -> str:
    """製品コード全体を 1 つの文字列として返す（★ 「どこかに在るか」を見る用）。

    ★ 使い分け: `X in product_text()` は「製品のどこかに在る」を見る。
      **位置は見ない** ── ファイルが連結されているので、`.index()` で切った窓は
      ファイル境界をまたぐ。窓が要るときは `window_around()` を使うこと。
    """
    return chr(10).join(p.read_bytes().decode("utf-8") for p in product_files())


def code_only_text() -> str:
    """製品コードから**コメントと docstring を落とした**本文（2026-09-20）。

    ★★ なぜ要るか: 「この字面は 1 箇所にしか無い」という契約を数える番人が、
      **自分たちの説明文**に当たって誤判定する。今日だけで 2 回踏んだ
      （「soffice を直に呼ぶと…」というコメント／「宛先と同じ名前と判定された」という注釈）。
    ★ 3 度目に同じ物を書くところだったので、ここに畳む（呼び出し側に持たせない）。
    ★ 文字列リテラルは残す ── 契約はたいてい**リテラル**についてのものだから。
    """
    import ast
    out = []
    for path in product_files():
        src = path.read_bytes().decode("utf-8")
        try:
            tree = ast.parse(src)
        except SyntaxError:          # ★ Python でない断片は数に入れない
            continue
        for node in ast.walk(tree):
            if not isinstance(node, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef,
                                     ast.ClassDef)):
                continue
            body = getattr(node, "body", [])
            if (body and isinstance(body[0], ast.Expr)
                    and isinstance(body[0].value, ast.Constant)
                    and isinstance(body[0].value.value, str)):
                node.body = body[1:] or [ast.Pass()]
        out.append(ast.unparse(ast.fix_missing_locations(tree)))
    return chr(10).join(out)


def count_in_code(needle: str) -> int:
    """★ コードの中だけでの出現回数（コメント・docstring は数えない）。"""
    return code_only_text().count(needle)


def product_strings() -> list:
    """製品コードの**画面に出うる文字列**を (ファイル, 行, 中身) で全部。

    ★ docstring は除く（概念を説明する文は画面に出ない）。コメントは ast が元から見ない。
    ★ なぜ在るか（2026-09-22）: 「この語が画面に出ていないこと」を確かめる番人が
      **字面の名簿**で書かれていた ── 言い換えた日に、名簿の語が消えて**空振りで緑**になる
      （恒真の罠）。★ 分母は宣言から引き、**在り処で**縛る:
      「宣言した 1 ファイルの外に、その語を字面で持つ文字列が在ってはいけない」。
    """
    import ast
    out = []
    for f in product_files():
        tree = ast.parse(f.read_text(encoding="utf-8"))
        docs = set()
        for node in ast.walk(tree):
            if isinstance(node, (ast.Module, ast.FunctionDef,
                                 ast.AsyncFunctionDef, ast.ClassDef)):
                body = getattr(node, "body", [])
                if (body and isinstance(body[0], ast.Expr)
                        and isinstance(body[0].value, ast.Constant)
                        and isinstance(body[0].value.value, str)):
                    docs.add(id(body[0].value))
        for node in ast.walk(tree):
            if (isinstance(node, ast.Constant) and isinstance(node.value, str)
                    and id(node) not in docs):
                out.append((f, node.lineno, node.value))
    return out
