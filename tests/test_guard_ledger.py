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
#      ★★ 2026-09-23: 初版の検出は**正規表現 1 つの形**しか見ておらず「0 件」と言っていたが、
#        AST で数え直すと **86 か所・53 ファイル**に在った（`Path(ailine.__file__).read_text()`・
#        `SRC / "ailine" / "__init__.py"`・`inspect.getsource(ailine)` ＝モジュール丸ごと）。
#        ★ 番人が「無い」と言ったのは、**無かった**のでなく**見えていなかった**。AST に替えた。
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

# ★ 「本体を場所で決め打ちして読む」形（2026-09-23 に正規表現から AST へ替えた）:
#   file_attr  … `ailine.__file__` から本体を読む（`.parents[..]` / `.parent.parent` で
#                **repo の根を探すだけ**の式は除く ── 読まないので無害）
#   path_lit   … `"__init__.py"` を含む短い文字列（`SRC / "ailine" / "__init__.py"` 等）
#   whole_mod  … `inspect.getsource(ailine)` ── **モジュール丸ごと＝本体 1 冊だけ**
#   main_file  … 芯の `MAIN_FILE`（本体の場所）を使う所 ── 使ってよいのは下の FILE_ITSELF だけ
_ROOT_ONLY = re.compile(r"\.parents\[|\.parent\.parent")
_READS = re.compile(r"read_text|read_bytes|open\(|\.py\b|ailine_core|getsource")

#: ★ 本体という**ファイルそのもの**を測る番人・道具（契約を読むのではない）。
#:   ここに載るものだけが場所を読んでよい。鍵は「repo からの相対パス::関数名」。
FILE_ITSELF = {
    "tests/test_line_budget.py::<module>":
        "本体というファイルそのものの行数を記録と突き合わせる番人。契約でなくファイルを"
        "測るのが目的なので、場所は芯の MAIN_FILE から引く。",
    "tests/test_readme_install_tag.py::test_the_first_command_exists_in_that_tag":
        "古いタグの時点の本体を git show で読む。その版の配置を名指しするので、"
        "今の場所（芯）に追随させてはいけない。",
    "scripts/deps_graph.py::modules":
        "モジュール地図を作る道具。本体というファイルを『ailine』モジュールとして図に"
        "載せるのが目的で、契約を読むのではない（場所は芯の MAIN_FILE から引く）。",
    "scripts/refresh_records.py::<module>":
        "本体の行数の記録（ailine_py_line_budget.txt）を書き直す道具。"
        "ファイルそのものの行数を測るのが目的。",
    "tests/split_progress_core.py::<module>":
        "分割の分母を出す測定器（つめ車と scripts/split_progress.py が共有）。本体という"
        "ファイルそのものの純ロジックを数えるのが目的で、場所は芯の MAIN_FILE から引く。",
}

#: ★ ailine_core を**再帰しない glob** で並べてよい所（理由つき）。
#:   `glob("*.py")` は ailine_core/postconditions/ のような下位パッケージを見ない ──
#:   2026-09-23 に 6 か所で見つかった（事後条件 5 ファイルが番人の視野の外だった）。
SHALLOW_CORE_GLOB = {
}

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


def _docstring_ids(tree) -> set:
    ids = set()
    for node in [tree] + [n for n in ast.walk(tree)
                          if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef))]:
        body = getattr(node, "body", None)
        if (body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant)
                and isinstance(body[0].value.value, str)):
            ids.add(id(body[0].value))
    return ids


def _parents_and_owner(tree):
    parents = {}
    for n in ast.walk(tree):
        for c in ast.iter_child_nodes(n):
            parents[c] = n

    def owner(n):
        while n in parents:
            n = parents[n]
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
                return n.name
        return "<module>"
    return parents, owner


def location_reads(source: str) -> list:
    """source の中で本体を**場所で**読む箇所を [(関数名, 形, 行)] で返す（AST で見る）。"""
    tree = ast.parse(source)
    parents, owner = _parents_and_owner(tree)
    docs = _docstring_ids(tree)

    def outer(n):                       # 式の一番外まで上がる（文に着いたら止まる）
        while isinstance(parents.get(n), ast.expr):
            n = parents[n]
        return ast.unparse(n)

    hits = []
    for n in ast.walk(tree):
        kind = None
        if (isinstance(n, ast.Attribute) and n.attr == "__file__"
                and isinstance(n.value, ast.Name) and n.value.id == "ailine"):
            whole = outer(n)
            if _ROOT_ONLY.search(whole) and not _READS.search(whole):
                continue                # repo の根を探すだけ（本体を読まない）
            kind = "file_attr"
        elif (isinstance(n, ast.Constant) and isinstance(n.value, str) and id(n) not in docs
              and "__init__.py" in n.value and len(n.value) < 60):
            kind = "path_lit"
        elif (isinstance(n, ast.Call) and ast.unparse(n.func).endswith("getsource")
              and n.args and isinstance(n.args[0], ast.Name) and n.args[0].id == "ailine"):
            kind = "whole_mod"
        elif ((isinstance(n, ast.Name) and n.id == "MAIN_FILE")
              or (isinstance(n, ast.Attribute) and n.attr == "MAIN_FILE")):
            kind = "main_file"
        if kind:
            hits.append((owner(n), kind, n.lineno))
    return hits


def shallow_core_globs(source: str) -> list:
    """ailine_core を**再帰しない glob** で並べている所を [(関数名, 行)] で返す。

    ★ 受け手が式の中で ailine_core を名指ししている形（`(SRC / "ailine_core").glob(...)`）と、
      ailine_core を指す名前に代入してから glob する形（`CORE_DIR.glob(...)`）の両方を拾う。
    ★ `rglob` と `glob("**/...")` は再帰するので拾わない。
    """
    tree = ast.parse(source)
    _parents, owner = _parents_and_owner(tree)
    core_names = {t.id for n in ast.walk(tree) if isinstance(n, ast.Assign)
                  and "ailine_core" in ast.unparse(n.value)
                  for t in n.targets if isinstance(t, ast.Name)}
    hits = []
    for n in ast.walk(tree):
        if not (isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                and n.func.attr == "glob" and n.args
                and isinstance(n.args[0], ast.Constant) and isinstance(n.args[0].value, str)
                and not n.args[0].value.startswith("**")):
            continue
        recv = n.func.value
        if ("ailine_core" in ast.unparse(recv)
                or (isinstance(recv, ast.Name) and recv.id in core_names)):
            hits.append((owner(n), n.lineno))
    return hits


def _scanned_files():
    """tests/ と scripts/ の .py（★ 芯と台帳自身は除く ── 芯は場所を書く唯一の所、
    台帳は検出器に食わせる悪い書き方を文字列で持つ）。"""
    repo = TESTS.parent
    out = []
    for p in sorted(list(TESTS.glob("*.py")) + list((repo / "scripts").glob("*.py"))):
        if p.name in ("_product_source.py", Path(__file__).name):
            continue
        out.append((p.relative_to(repo).as_posix(), p))
    return out


def _all_location_reads() -> dict:
    """{"相対パス::関数名": 件数} ── 免除（FILE_ITSELF）も含めた全部。"""
    out = {}
    for rel, p in _scanned_files():
        for fn, _kind, _ln in location_reads(p.read_bytes().decode("utf-8")):
            key = f"{rel}::{fn}"
            out[key] = out.get(key, 0) + 1
    return out


def _hard_reads():
    """場所で読む箇所のうち、FILE_ITSELF で許していないもの。"""
    return {k: n for k, n in _all_location_reads().items() if k not in FILE_ITSELF}


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
                          (LOOP_ONLY_ASSERTS, "LOOP_ONLY_ASSERTS"),
                          (FILE_ITSELF, "FILE_ITSELF"),
                          (SHALLOW_CORE_GLOB, "SHALLOW_CORE_GLOB")):
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


def test_the_file_itself_exemptions_are_still_needed():
    """★ 免除表（FILE_ITSELF）に、もう検出されない項目が残っていたら赤くする。

    ★ 古い免除は「ここは場所を読んでよい」という許可だけが残った状態 ──
      次に同じ関数へ契約の読みを足しても、誰にも鳴らない。
    """
    now = set(_all_location_reads())
    stale = sorted(set(FILE_ITSELF) - now)
    assert not stale, f"もう場所を読んでいないのに FILE_ITSELF に残っている: {stale}"


def test_the_location_detector_catches_the_known_shapes():
    """★ 陽性対照 ── 既知の悪い書き方を検出器に食わせて、全部拾うこと。

    ★★ 初版の正規表現は 1 つの形しか見ておらず、86 か所を「0 件」と言っていた。
      在庫ゼロと検出の故障は数だけでは区別できない ── 鳴るべき物で鳴ることを先に確かめる。
    """
    bad = {
        "file_attr": "import ailine\nfrom pathlib import Path\n"
                     "def f():\n    return Path(ailine.__file__).read_text(encoding='utf-8')\n",
        "path_lit": "def f():\n"
                    "    return (REPO / 'src' / 'ailine' / '__init__.py').read_text()\n",
        "whole_mod": "import inspect\nimport ailine\nSRC = inspect.getsource(ailine)\n",
        "main_file": "from _product_source import MAIN_FILE\n"
                     "def f():\n    return MAIN_FILE.read_text()\n",
    }
    for kind, src in bad.items():
        got = [k for _fn, k, _ln in location_reads(src)]
        assert got == [kind], f"{kind} の書き方を拾えていない: {got}"
    # ★ 陰性対照: 根を探すだけの式・名前で引く getsource は拾わない
    good = ("import ailine\nfrom pathlib import Path\n"
            "REPO = Path(ailine.__file__).resolve().parents[2]\n"
            "UP = Path(ailine.__file__).resolve().parent.parent\n"
            "import inspect\nS = inspect.getsource(ailine.codegen_dsl)\n")
    assert location_reads(good) == [], location_reads(good)


def test_no_shallow_glob_over_ailine_core():
    """★ tests/ と scripts/ で ailine_core を**再帰しない glob** で並べていないこと。

    ★★ 2026-09-23: `(SRC / "ailine_core").glob("*.py")` の形が 6 か所に在り、
      ailine_core/postconditions/ の 5 ファイルを**見ていなかった**（移植可能性の番人は
      postconditions が ailine を import しても鳴らなかった）。
    ★ 直し方: `_product_source` の `product_files()` / `src_files()` を使うか、`rglob`。
    """
    found = {}
    for rel, p in _scanned_files():
        for fn, ln in shallow_core_globs(p.read_bytes().decode("utf-8")):
            found.setdefault(f"{rel}::{fn}", []).append(ln)
    bad = {k: v for k, v in found.items() if k not in SHALLOW_CORE_GLOB}
    assert not bad, (
        f"ailine_core を再帰しない glob で並べている: {bad} ── 下位パッケージ"
        "（postconditions/ 等）が視野から落ちる。product_files() か rglob を使うこと")
    stale = sorted(set(SHALLOW_CORE_GLOB) - set(found))
    assert not stale, f"もう浅い glob を使っていないのに SHALLOW_CORE_GLOB に残っている: {stale}"


def test_the_shallow_glob_detector_catches_the_known_shapes():
    """★ 陽性対照 ── 浅い glob の 2 つの形を拾い、再帰する形は拾わないこと。"""
    inline = "def f():\n    return sorted((SRC / 'ailine_core').glob('*.py'))\n"
    named = ("CORE_DIR = REPO / 'src' / 'ailine_core'\n"
             "def g():\n    return [p for p in CORE_DIR.glob('*.py')]\n")
    assert [fn for fn, _ln in shallow_core_globs(inline)] == ["f"], shallow_core_globs(inline)
    assert [fn for fn, _ln in shallow_core_globs(named)] == ["g"], shallow_core_globs(named)
    recursive = ("CORE_DIR = REPO / 'src' / 'ailine_core'\n"
                 "A = sorted(CORE_DIR.rglob('*.py'))\n"
                 "B = sorted(CORE_DIR.glob('**/*.py'))\n"
                 "C = sorted((REPO / 'tests').glob('*.py'))\n")
    assert shallow_core_globs(recursive) == [], shallow_core_globs(recursive)
