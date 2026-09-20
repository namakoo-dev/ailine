# -*- coding: utf-8 -*-
"""画面の「実物の見え方」が、**LibreOffice が描いたものをそのまま映す**こと（2026-09-20）。

★★ 出所（Namakoo 実測）:「今は ailine の確認用に LO を別で開いてるけど、そこがかなり
  煩わしい。往復せずに使いたい。ailine のウィンドウだけで LO 使えたら便利かなと思って」。
  確かめるたびにアプリを切り替えていた ── 見たいのは**実際にどう見えるか**
  （列幅・結合・罫線・グラフ）で、値の一覧では足りない。

★★ 採った形（A: 見るだけ）: LO → PDF → **ブラウザがそのまま描く**。
  ★ 新しい依存を 1 つも足さない（GUI の縛り④）── 画像化の道具も使わない。
  ★ 描くのは製品の `ailine export-pdf`（殻は本体を叩く・縛り③）。ここで soffice を
    直に呼ぶと、探し方と関所が 2 つ目の実装になる。
  ★ Collabora を埋める案（B・編集もできる）は Docker と常駐サーバを買い手に要求するので
    ローカル配布では採らなかった（クラウド版は別ラインで保留）。

★ この試験が守る契約:
  ① 画面は PDF を**本体経由で**取る（soffice を画面が直に呼ばない）
  ② 描けなかったら**理由を出す**（空を見せない ── 出ないことは信号でない）
  ③ 人のフォルダに描いたものを置かない（soffice に人のフォルダを触らせない）
  ④ 見え方の判断は 1 箇所（呼び出し側に書き写さない＝片配線を作らない）
"""
from __future__ import annotations

import re
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

GUI = REPO / "gui"
SERVER = (GUI / "server.py").read_bytes().decode("utf-8")
HTML = (GUI / "index.html").read_bytes().decode("utf-8")


def _code_only(py: str) -> str:
    """**コードだけ**を返す（行コメントも docstring も落とす）。

    ★★ 2026-09-20: 初版は行コメントしか落としておらず、番人が**自分の説明文**
      （「ここで soffice を直に呼ぶと…」）に当たって赤くなった。この repo は同じことを
      2 度している（JS 側でも踏んだ）── 契約はコードについてのものなので、見る対象を正す。
      ★ 緩めてはいない: 文字列リテラルは残る（`"export-pdf"` は見えたままでなければ困る）。
    """
    import ast
    tree = ast.parse(py)
    for node in ast.walk(tree):
        if not isinstance(node, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef,
                                 ast.ClassDef)):
            continue
        body = getattr(node, "body", [])
        if (body and isinstance(body[0], ast.Expr)
                and isinstance(body[0].value, ast.Constant)
                and isinstance(body[0].value.value, str)):
            node.body = body[1:] or [ast.Pass()]
    return ast.unparse(ast.fix_missing_locations(tree))


def _func_code(src: str, name: str) -> str:
    """その関数の**コードだけ**を返す（docstring と行コメントを落とす）。

    ★★ 文字数で窓を切らない ── 中身が増減した日に、番人が見る場所が黙ってずれる
      （この repo が今週 6 回踏んだ形。窓の広さ・関数名・印の字面・`-m local` の字面）。
    ★ 元のソースから切り出す（`ast.unparse` は引用符を書き換えるので、
      `"--out"` のような**字面の契約**が見えなくなる）。
    """
    import ast
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        if node.name != name:
            continue
        seg = ast.get_source_segment(src, node) or ""
        body = node.body
        if (body and isinstance(body[0], ast.Expr)
                and isinstance(body[0].value, ast.Constant)
                and isinstance(body[0].value.value, str)):
            doc = ast.get_source_segment(src, body[0]) or ""
            seg = seg.replace(doc, "", 1)
        return chr(10).join(ln.split("#")[0] for ln in seg.split(chr(10)))
    raise AssertionError(f"関数が無い: {name}")


def _server_module():
    """`gui/server.py` を**モジュールとして**読む（サーバは立てない）。

    ★ 殻は `if __name__ == "__main__"` でしか走らないので、import しても何も起きない。
    ★ 走らせて確かめるために要る ── ソースの字面を読むだけの番人は、変異を素通りさせた。
    """
    import importlib.util
    spec = importlib.util.spec_from_file_location("_ailine_gui_server", GUI / "server.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _script() -> str:
    m = re.search(r"<script>(.*?)</script>", HTML, re.S)
    assert m, "script が無い"
    return chr(10).join(ln.split("//")[0] for ln in m.group(1).split(chr(10)))


def test_the_picture_comes_from_the_product_not_from_soffice_directly():
    """★★ ① 殻は本体を叩く ── 画面が soffice を直に呼ばないこと。

    ★ 直に呼ぶと「LibreOffice の探し方」と「人のフォルダを触らせない関所」が
      2 つ目の実装になる（この repo が何度も潰してきた形）。
    """
    code = _code_only(SERVER)
    assert "export-pdf" in code, "★ 製品の export-pdf を使っていない"
    for forbidden in ("soffice", "--convert-to", "office_dir"):
        assert forbidden not in code, f"★ 画面が LibreOffice を直に呼んでいる: {forbidden}"


def test_a_failure_to_draw_is_said_out_loud():
    """★★ ② 描けなかったら理由を出す（空を見せない）。

    ★ LibreOffice が無い環境はある。そこで白い枠だけ出すと「表が空だ」と読まれる ──
      この repo の古い線「出ないことは信号でない」そのもの。
    """
    # ★★ 2026-09-20: 初版は `return None, ` の**回数**を数えていたので、
    #   理由を空文字にする変異（`return None, ""`）を緑で通した ── 字面を見ていた。
    #   ★ 走らせて確かめる: 描けなかった回は**空でない理由**が返ること。
    srv = _server_module()
    with tempfile.TemporaryDirectory() as td:
        book = Path(td) / "x.xlsx"
        book.write_bytes(b"not a real xlsx")
        real = srv._ailine
        srv._ailine = lambda args, answer=None: (9, "× LibreOffice を開けません", None)
        try:
            got, why = srv._render_pdf(book, None)
        finally:
            srv._ailine = real
    assert got is None, "★ 描けていないのに絵を返している"
    assert why and why.strip(), "★ 描けなかった理由が空（画面は白い枠を出すしかない）"

    js = _script()
    j = js.index("function showReal")
    pane = js[j:j + 1400]
    assert "error" in pane, "★ 画面が理由を出していない"
    assert "application/pdf" in pane, "★ PDF かどうかを見ていない（失敗を絵として出す）"


def test_nothing_is_drawn_into_the_human_folder():
    """★★ ③ 描いたものは人のフォルダに置かない。

    ★ 2026-08-26 の実測: soffice に人のフォルダを渡すと、同名の PDF が**予告なく消えた**
      （顧客へ送った確定版が exit 0 のまま消滅）。だから出力先は一時フォルダに固定する。
    """
    # ★★ 2026-09-20: 初版は `_RENDER_DIR` が**関数の中に在るか**だけを見ていたので、
    #   出力先を `path.with_suffix(".pdf")` に変える変異を緑で通した ── また字面だった。
    #   ★ 走らせて確かめる: 出来た絵が**人のフォルダに無い**こと。
    srv = _server_module()
    with tempfile.TemporaryDirectory() as td:
        book = Path(td) / "確認.xlsx"
        book.write_bytes(b"dummy")
        seen = {}

        def fake(args, answer=None):
            # ★ 本体の代わり ── `--out` で言われた所に絵を置いたことにする
            out = Path(args[args.index("--out") + 1])
            seen["out"] = out
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_bytes(b"%PDF-1.4 fake")
            return 0, "", None

        real = srv._ailine
        srv._ailine = fake
        try:
            got, why = srv._render_pdf(book, None)
        finally:
            srv._ailine = real
        assert why is None and got is not None, (why, got)
        assert Path(td) not in got.parents, f"★ 人のフォルダに描いている: {got}"
        assert Path(tempfile.gettempdir()) in got.parents, f"★ 一時フォルダの外: {got}"
        assert not list(Path(td).glob("*.pdf")), "★ 人のフォルダに PDF が残った"


def test_the_same_book_is_not_drawn_twice():
    """★ 同じ冊を何度も描かない ── ただし**中身が変わったら描き直す**こと。

    ★ 鍵に更新時刻が入っていないと、実行して変わった後も**古い絵**を見せ続ける
      （画面が嘘をつく形）。速さのために正しさを落とさない。
    """
    body = _func_code(SERVER, "_render_pdf")
    assert "st_mtime_ns" in body, "★ 鍵に更新時刻が入っていない（古い絵が居座る）"
    assert "if out.exists():" in body, "★ 溜めた絵を使い回していない"


def test_the_view_decision_lives_in_one_place():
    """★★ ④ 見え方の判断は 1 箇所 ── 呼び出し側に書き写さない。

    ★ 枠を描く所は 5 箇所ある。そこに「実物なら…」を書き写すと、片方だけ直る
      （この repo が何度も踏んだ「片配線」）。だから 1 関数に畳んで全部そこを通す。
    """
    js = _script()
    assert js.count('$("#viewmode").value === "real"') == 1, (
        "★ 見え方の判断が 2 箇所以上ある（1 関数に畳むこと）")
    assert "function drawPane" in js, "★ 畳んだ関数が無い"
    # ★ 2 つの枠は畳んだ関数を通ること（直接 drawTable を呼ぶと実物に切り替わらない）
    for pane in ('$("#before")', '$("#after")'):
        assert f"drawPane({pane}" in js, f"★ {pane} が畳んだ関数を通っていない"
    assert f'drawTable($("#before")' not in js, "★ 枠を直接描いている箇所が残っている"
    assert f'drawTable($("#after")' not in js, "★ 枠を直接描いている箇所が残っている"


def test_the_toggle_is_offered_and_defaults_to_the_cheap_one():
    """★ 切り替えが在り、既定は**速い方**（実物は描くのに数秒かかる）。"""
    assert 'id="viewmode"' in HTML, "★ 切り替えが無い"
    i = HTML.index('id="viewmode"')
    bar = HTML[i:i + 400]
    assert bar.index('value="grid"') < bar.index('value="real"'), (
        "★ 既定が実物になっている（毎回 LibreOffice を待たせる）")
    assert "LibreOffice" in bar, "★ 何が描いているのかを画面で言っていない"


def test_switching_back_redraws_instead_of_leaving_the_old_picture():
    """★ 簡易へ戻した時に、古い絵が居座らないこと。"""
    js = _script()
    i = js.index('$("#viewmode").onchange')
    block = js[i:i + 400]
    assert "redrawBasis()" in block, "★ 簡易へ戻しても表を描き直していない"
