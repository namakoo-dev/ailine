# GUI の番人（2026-08-26）── 画面は**判定を作らない**。
#
# ★ なぜ機械で縛るか: 「GUI は薄い殻にする」は口約束では守れない。
#   この repo は 2026-08 の盲検で「検算の分母が、疑うべき対象と同じ所から作られる」形の
#   欠陥を 4 回踏んだ。画面が advisories を数えたり postcondition から印を導いたりすれば、
#   それは 2 つ目の実装で、同じ欠陥をこちらで新造することになる。
#
# 契約:
#   ① 画面が読む判定は `verdict` ただ 1 つ（postcondition/advisories から導かない）
#   ② 判定の語彙は ailine 本体が出しうる値と一致する（画面が勝手な語を持たない）
#   ③ 殻は**この repo の本体**を叩く（site-packages の古い版を叩かない・実測で踏んだ）
#   ④ 新しい依存を足さない（標準ライブラリだけ）
#   ⑤ 画面を止めるモーダル（alert/confirm/prompt）を使わない

import ast
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
GUI = REPO / "gui"
HTML = (GUI / "index.html").read_text(encoding="utf-8")
SERVER = (GUI / "server.py").read_text(encoding="utf-8")
sys.path.insert(0, str(REPO / "src"))


def _script(html: str, code_only: bool = False) -> str:
    m = re.search(r"<script>(.*?)</script>", html, re.S)
    assert m, "script が無い"
    js = m.group(1)
    if not code_only:
        return js
    # ★ 番人が**自分の説明文**に引っかかった（コメントで `advisories` や `alert(` に触れて
    #   いる）。契約はコードについてのものなので、見る対象を正す ── 緩めてはいない。
    #   ★ 行コメントだけを落とす（この画面にブロックコメントは無い ── 下で機械が確かめる）。
    assert "/*" not in js, "ブロックコメントが増えた（この番人の前提が崩れる）"
    return chr(10).join(re.sub(r"//.*$", "", ln) for ln in js.split(chr(10)))


# --- ① 判定を導かない -----------------------------------------------------------------

def test_the_page_never_derives_a_verdict():
    js = _script(HTML, code_only=True)
    for forbidden in ("postcondition", "advisories", "claims"):
        assert forbidden not in js, (
            f"画面が `{forbidden}` を読んでいる ── 判定を自分で導きかけている。"
            "映してよいのは ailine が返した verdict だけ")
    assert "j.verdict" in js, "verdict を読んでいない"


def test_the_server_does_not_touch_the_verdict():
    tree = ast.parse(SERVER)
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            assert node.value not in ("verified", "warned", "unverified", "unobservable"), (
                "サーバが判定の語を組み立てている（本体の値をそのまま運ぶだけにする）")


# --- ② 語彙が本体と一致する -----------------------------------------------------------

def test_the_page_knows_exactly_the_verdicts_the_tool_can_emit():
    """★ 恒真殺し: 本体が出す語を**本体側から**取り、画面の表と突き合わせる。

    本体が新しい判定を足したのに画面が知らなければ、無印で表示されてしまう。
    """
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    i = src.index('result["verdict"] = (')
    block = src[i:i + 400]
    emitted = set(re.findall(r'"(verified|warned|unverified|unobservable)"', block))
    assert emitted == {"verified", "warned", "unverified", "unobservable"}, emitted
    emitted.add("not_applied")     # 適用まで行かなかった run（_finish_run の既定）
    js = _script(HTML)
    known = set(re.findall(r"^\s{2}(\w+):\s*\{mark:", js, re.M))
    assert emitted <= known, f"画面が知らない判定がある: {sorted(emitted - known)}"


# --- ③ この repo の本体を叩く ---------------------------------------------------------

def test_the_shell_runs_this_repo_not_the_installed_copy():
    """★ 実測（2026-08-26）: これが無いと子プロセスは site-packages の**古い版**を
       import する。盲検 2 回目で検分者が古いタグを測ったのと同じ形の事故。"""
    assert 'env["PYTHONPATH"]' in SERVER, "本体の在り処を子プロセスに渡していない"
    assert 'REPO / "src"' in SERVER


# --- ④⑤ 依存とモーダル ----------------------------------------------------------------

def test_the_shell_adds_no_dependency():
    tree = ast.parse(SERVER)
    mods = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            mods.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            mods.add(node.module.split(".")[0])
    third_party = mods - set(sys.stdlib_module_names)
    assert not third_party, f"標準ライブラリ以外を入れた: {sorted(third_party)}"


def test_no_blocking_modal():
    """★ 実測: alert() がレンダラを止め、画面が固まった（デモ中なら致命的）。"""
    js = _script(HTML, code_only=True)
    for bad in ("alert(", "confirm(", "prompt("):
        assert bad not in js, f"画面を止めるモーダルを使っている: {bad}"
