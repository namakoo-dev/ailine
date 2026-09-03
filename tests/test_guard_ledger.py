# 番人の**台帳**（2026-09-03）── 番人を数える番人。
#
# ★★ 発端（Namakoo）:「番人はそれぞれ独立していると思うが、管理しきれているか？
#   番人にも抜けがあるのが怖い。**静かな壊れ方の要因は番人が通してしまうことにより起こる**」
#   ── その日のうちに実例が出た。事後条件 45 関数を ailine_core へ移したら、
#   **7 件の番人が同時に空振りした**（`src/ailine/__init__.py` 決め打ちで読んでいた）。
#   ★ 実装は 1 つも壊れていない。契約もそのまま成り立っていた。**視野だけが古かった。**
#
# ★★ 「集める」のではなく「数える」で管理する（2026-09-03 に測って決めた）:
#   番人 126 個のうち **94% が守る対象の機能テストと同居**している。それは正しい配置で、
#   寄せると「対象を直す人の目に番人が入らなくなる」── 更新し忘れが起きる。
#   足りないのは置き場所ではなく ① 共通の芯（tests/_product_source.py）と ② この台帳。
#
# ★★ 静かに壊れる 3 つの形（この台帳が数えるもの）:
#   ① **場所の決め打ち** … `(REPO/"src"/"ailine"/"__init__.py").read_text()`
#      分割で実装が動いた瞬間に空振りする。★ 名前で引く `inspect.getsource(ailine.X)` なら
#      移動に追随し、公開面の凍結が「その名前が在ること」を守る（二重防御）
#   ② **quiet な assert の単独** … `not in` は探す場所が空なら必ず通る。
#      必ず「在ること」の assert（loud）と対で置く
#   ③ **回らないループ** … `for m in re.finditer(...): assert ...` は 0 件マッチで
#      1 回も回らずに通る。★ ①②より見つけにくい（assert 自体が実行されない）
#
# ★ 在庫は**ゼロから始めない**。いまの数（決め打ち 49 箇所 / 30 ファイル）を上限として
#   置き、減らしていく。★ ゼロを待つと置けないまま増え続ける ── この repo が
#   「番人が無い所は触らない」を守るために、まず数を持つ。

import ast
import re
from pathlib import Path

TESTS = Path(__file__).resolve().parent

# --- ① 場所の決め打ち（2026-09-03 の初回計測）------------------------------------------
HARDCODED_READS_AT_FIRST_COUNT = 0    # ★ 2026-09-04 に在庫ゼロを達成。1 箇所でも増えたら赤
HARDCODED_FILES_AT_FIRST_COUNT = 0    # ★ 同上

# ★ 「本体を場所で決め打ちして読む」形。★ クォート文字を一切見ない ──
#   クォートの種類に依存すると壊れる（ast.unparse はシングルクォートで出す）。
_HARD = re.compile(r"REPO[^)]{0,80}__init__[^)]{0,20}\)\.read_text")

# --- ② quiet を単独で持つ番人（★ 理由つきで許すものだけ）--------------------------------
QUIET_WITHOUT_LOUD = {
    "test_vanishing_shapes.py::test_detection_does_not_ask_openpyxl":
        "src.index('def vanishing_shapes(') が先にあるので、関数が消えれば ValueError で"
        "落ちる（実質 loud が対になっている）。★ この形は許す。",
}

# --- ③ 回らないループ（★ 最も見つけにくい）---------------------------------------------
LOOP_ONLY_ASSERTS = {}


def _code_text(path: Path) -> str:
    """コメントと docstring を落とした**コードだけ**の文字列。

    ★ なぜ要るか: 初版は生のテキストを正規表現で見ていたので、**この台帳自身の
      説明文**が検出に引っかかった（2026-08-31 に「番人が自分の説明文に引っかかる」
      形を一度踏んでいる ── 同じ轍）。番人は**コードだけ**を見る。
    """
    tree = ast.parse(path.read_bytes().decode("utf-8"))
    for node in ast.walk(tree):
        body = getattr(node, "body", None)
        if isinstance(body, list) and body and isinstance(body[0], ast.Expr)            and isinstance(body[0].value, ast.Constant)            and isinstance(body[0].value.value, str):
            body.pop(0)
    return ast.unparse(tree)


def _guard_files():
    """★ 台帳自身は数えない（自分を数えると必ず 1 件出て、意味が消える）。"""
    return [p for p in sorted(TESTS.glob("test_*.py")) if p.name != Path(__file__).name]


def _test_funcs(path: Path):
    tree = ast.parse(path.read_bytes().decode("utf-8"))
    return [n for n in tree.body
            if isinstance(n, ast.FunctionDef) and n.name.startswith("test_")]


def _reads_source(src: str) -> bool:
    return bool(re.search(r"read_text\(encoding|getsource\(|count_in_product|window_around",
                           src))


def _hard_reads():
    out = {}
    for p in _guard_files():
        n = len(_HARD.findall(_code_text(p)))
        if n:
            out[p.name] = n
    return out


def test_hardcoded_reads_do_not_grow():
    """① 場所の決め打ちが増えないこと（★ 減らすのは歓迎）。

    ★ 直し方: `from _product_source import count_in_product, window_around` を使うか、
      `inspect.getsource(ailine.関数名)` で**名前から引く**。
    """
    hard = _hard_reads()
    total, files = sum(hard.values()), len(hard)
    assert total <= HARDCODED_READS_AT_FIRST_COUNT, (
        f"本体を場所で決め打ちする番人が増えた（{total} 箇所 / 初回計測 "
        f"{HARDCODED_READS_AT_FIRST_COUNT}）── 分割で実装が動くと空振りする。"
        f"_product_source の芯か inspect.getsource を使うこと")
    assert files <= HARDCODED_FILES_AT_FIRST_COUNT, (
        f"決め打ちのファイルが増えた（{files} 本 / 初回計測 "
        f"{HARDCODED_FILES_AT_FIRST_COUNT}）")


def test_quiet_assertions_are_paired_with_loud_ones():
    """② 「無いこと」の assert を単独で置かない。

    ★ `not in` は**探す場所が空なら必ず通る**。だから「在ること」の assert と対にする。
      対になっていれば、視野が古くなった時に loud の側が鳴る。
    """
    lone = []
    for p in _guard_files():
        for fn in _test_funcs(p):
            s = ast.unparse(fn)
            if not _reads_source(s):
                continue
            asserts = [ast.unparse(x.test) for x in ast.walk(fn) if isinstance(x, ast.Assert)]
            quiet = [a for a in asserts if re.search(r"\bnot in\b", a)]
            # ★ loud = 「探す場所が空なら赤くなる」形。2026-09-03 に判定を広げた ──
            #   初版は `in` / `== N` しか見ておらず、`assert hits`（分母の確認）を
            #   loud と認めなかった。**分母を確かめる assert こそ最も loud**。
            loud = [a for a in asserts
                    if re.search(r"(?<!not )\bin\b|==\s*\d|>=\s*[1-9]", a)
                    or re.fullmatch(r"[\w.]+|len\([^)]*\)|list\([^)]*\)", a.strip())]
            if quiet and not loud:
                key = f"{p.name}::{fn.name}"
                if key not in QUIET_WITHOUT_LOUD:
                    lone.append(key)
    assert not lone, (
        f"「無いこと」の assert が単独で置かれている: {lone} ── "
        "探す場所が空でも通る。『在ること』の assert と対にするか、"
        "なぜ単独でよいかを QUIET_WITHOUT_LOUD に書くこと")


def test_the_ledger_does_not_keep_stale_entries():
    """★ 直したのに台帳に残っていたら赤くする（古い不安を配らない）。"""
    now = set()
    for p in _guard_files():
        for fn in _test_funcs(p):
            s = ast.unparse(fn)
            if not _reads_source(s):
                continue
            asserts = [ast.unparse(x.test) for x in ast.walk(fn) if isinstance(x, ast.Assert)]
            if any(re.search(r"\bnot in\b", a) for a in asserts):
                now.add(f"{p.name}::{fn.name}")
    stale = sorted(set(QUIET_WITHOUT_LOUD) - now)
    assert not stale, f"もう quiet を持たない番人が台帳に残っている: {stale}"


def test_every_exemption_states_a_reason():
    """★ 免除は増やせるが、黙っては増やせない。"""
    for table, label in ((QUIET_WITHOUT_LOUD, "QUIET_WITHOUT_LOUD"),
                          (LOOP_ONLY_ASSERTS, "LOOP_ONLY_ASSERTS")):
        for key, why in table.items():
            assert len(why) >= 30 and "。" in why, f"{label}[{key}] の理由が薄い"


def test_the_core_is_actually_used_somewhere():
    """★ 陽性対照 ── 芯（_product_source）が実際に使われていること。

    ★ 芯を作っただけで誰も通っていなければ、この台帳は「直す道がある」と嘘をつく。
    """
    users = [p.name for p in _guard_files() if "_product_source" in _code_text(p)]
    assert len(users) >= 6, f"芯を使っている番人が {len(users)} 本しかない: {users}"


def test_the_backlog_is_visible():
    """★ 在庫がゼロであること ── そして「検出が壊れて 0」と区別すること。

    ★★ 経緯（2026-09-03〜04）: 49 箇所 / 30 ファイルから始めた。ゼロを待つと
      台帳を置けないまま増え続けるので、**その日の実測を上限にして置いた**。
      翌日 47 箇所を芯へ載せ替え、残り 2 箇所も手で畳んで **ゼロに到達**した。

    ★ ゼロになると、初版の下限（`total >= 1` ── 0 は検出が壊れている疑い）が
      逆に赤くなる。**在庫ゼロと検出の故障を、数だけでは区別できない**ので、
      ここでは「芯が実際に使われていること」を陽性対照にする ──
      36 本以上の番人が芯を通っていれば、検出は生きている。
    """
    total = sum(_hard_reads().values())
    assert total == 0, (
        f"本体を場所で決め打ちする番人が {total} 箇所ある ── "
        "_product_source の芯か inspect.getsource を使うこと")
    users = [p.name for p in _guard_files()
             if "_product_source" in _code_text(p)]
    assert len(users) >= 30, (
        f"芯を使う番人が {len(users)} 本しかない ── 検出が壊れて 0 に見えている疑い")

def test_no_guard_asserts_only_inside_a_loop():
    """③ 「回らないループ」を数える。

    ★ `for m in re.finditer(...): assert ...` は **0 件マッチなら 1 回も回らずに通る**。
      quiet な assert より見つけにくい ── assert 自体が実行されないので、
      「探した結果 無かった」と「そもそも探せていない」の区別がつかない。
    ★ 直し方: マッチを list にして `assert hits` で**分母を先に確かめる**。
    """
    bad = []
    for p in _guard_files():
        for fn in _test_funcs(p):
            if not _reads_source(ast.unparse(fn)):
                continue
            loops = [n for n in ast.walk(fn) if isinstance(n, (ast.For, ast.AsyncFor))]
            # ★ リテラルのタプル/リストを回るループは除外する（2026-09-03）──
            #   `for code in ('1','3','4')` は**空になりえない**ので必ず回る。
            #   危ないのは「表を回る」形で、表が空になった日に黙る。
            loops = [lp for lp in loops
                     if not isinstance(lp.iter, (ast.Tuple, ast.List, ast.Set))]
            if not loops:
                continue
            in_loop = {id(x) for lp in loops for x in ast.walk(lp)
                       if isinstance(x, ast.Assert)}
            all_as = [x for x in ast.walk(fn) if isinstance(x, ast.Assert)]
            if all_as and all(id(x) in in_loop for x in all_as):
                key = f"{p.name}::{fn.name}"
                if key not in LOOP_ONLY_ASSERTS:
                    bad.append(key)
    assert not bad, (
        f"assert がループの中にしか無い番人: {bad} ── "
        "0 件マッチなら 1 回も回らずに通る。`assert hits` で分母を先に確かめること")
