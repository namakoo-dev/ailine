# 同名で定義された関数の**台帳**（2026-09-03）── 写し取られた実装を、増える前に見つける。
#
# ★★ 発端: 分割の前準備で数えたら、**同じ数行が最大 7 箇所に写されていた**
#   （`_is_number` 7 / `fmt_num` 3 / `_column_index` 2）。しかも docstring が互いを
#   指していた ──「total_row._is_number と同じ線」「stack.fmt_num と同じ線」。
#   ★ **重複を自覚したまま、注記で済ませていた。** 系譜「二重化した経路は片配線が
#   既定で起きる」の、いちばん静かな側の実例。
#
# ★ 処方は系譜どおり ── 両方直すのではなく **1 関数に畳んで呼び出し側に持たせない**
#   （→ `ailine_core/primitives.py`）。この台帳はその状態を凍結する番人。
#
# ★★ 「同名」だけでは判定にならない（2026-09-03 に 3 度読み違えた）:
#   ① クラスのメソッドを混ぜると `ok` のような偽の重複が出る → **トップレベルだけ**見る
#   ② AST が同じでも、**参照する定数の値が違えば挙動が違う** → 値まで比べる
#   ③ 実装が違って見えても、**別名を 1 つ挟んでいるだけ**のことがある
#      （`stack._is_blank` は `total_row._is_blank_cell` への 1 行委譲だった）
#
# 契約:
#   ① 同一実装の重複が**ゼロ**であること（畳んだ状態を守る）
#   ② 同名で違う実装のものは、**理由つきで**この台帳に在ること
#   ③ 台帳に無い同名が現れたら赤（黙って増やせない）

import ast
import hashlib
import importlib
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

# --- 同名だが「違っていてよい」もの（★ 理由を書く。無い行は赤になる）------------------
KNOWN_DIFFERENT = {
    "_values_agree": "csv_quarantine は型込み等値で TOLERANCE 不使用（種別を跨いだ一致を"
                      "認めない）。verify は帰属検算で許容誤差つき。★ 同じ名前で違う判断"
                      "なので、畳むと片方の規則がもう片方に漏れる。",
    "suggest_ops": "本体は照合プールを組んで ailine_core/suggest.py の純ロジックへ委譲する"
                    "薄いラッパ。core 側は pool を引数で受ける。★ 重複ではなく層の委譲。",
    "_row_has_any_value": "extract_multi は total_row._is_blank_cell を、stack は自分の "
                           "_is_blank（= total_row._is_blank_cell への 1 行委譲）を呼ぶ。"
                           "★ 実装は実質同じ。別名を 1 つ挟んでいるので機械には相違に見える。",
    "describe": "inspection は所見 1 件を 1 行に、row_identity は崩れた等式を 1 行にする。"
                 "★ 汎用名がぶつかっただけで、完全に別物。",
}

# --- ★ 畳んだもの（ここが再び増えたら退行）--------------------------------------------
FOLDED_INTO_PRIMITIVES = {"is_number", "fmt_num", "column_index"}

# --- ★★ 意図的な写し（畳んではいけない）------------------------------------------------
#
# ★ 2026-09-03 の学び: **重複が悪とは限らない。契約のための重複がある。**
#   `_is_number` を 7 箇所から畳んだとき、sum_identity だけは畳んではいけなかった ──
#   このモジュールは「標準ライブラリだけで閉じる」契約（言語非依存）を持っており、
#   ailine_core を import した瞬間に
#   test_module_is_portable_and_needs_no_spreadsheet_library が赤くなる。
#   ★ 畳もうとして契約を破り、**その番人が捕まえた**（人ではなく機械が止めた）。
#
# ★ 写しを許すなら、**写し同士が一致していることを機械が守る**必要がある ──
#   片方だけ直るのが、この repo の言う片配線そのものだから。
INTENTIONAL_COPIES = {
    ("ailine_core.sum_identity", "_is_number", "ailine_core.primitives", "is_number"):
        "sum_identity は『標準ライブラリだけで閉じる』契約（言語非依存・他言語へ移植できる"
        "形を保つ）。ailine_core を import できないので、同じ実装をあえて写している。",
}


def _toplevel_defs():
    """(関数名) → [(場所, 実装ハッシュ, 参照している自国のトップレベル名)]"""
    files = [("本体", REPO / "src/ailine/__init__.py", "ailine")]
    files += [(p.name, p, f"ailine_core.{p.stem}")
              for p in sorted((REPO / "src/ailine_core").glob("*.py"))
              if p.stem != "__init__"]
    out = {}
    for label, path, mod in files:
        text = path.read_bytes().decode("utf-8")
        tree = ast.parse(text)
        local = {n.name for n in tree.body if isinstance(n, (ast.FunctionDef, ast.ClassDef))}
        local |= {t.id for n in tree.body if isinstance(n, ast.Assign)
                  for t in n.targets if isinstance(t, ast.Name)}
        for n in tree.body:                       # ★ トップレベルだけ（メソッドは別物）
            if not isinstance(n, ast.FunctionDef):
                continue
            body = [x for x in n.body
                    if not (isinstance(x, ast.Expr) and isinstance(x.value, ast.Constant))]
            code = ast.unparse(ast.Module(body=body, type_ignores=[]))
            free = sorted({x.id for x in ast.walk(n) if isinstance(x, ast.Name)} & local)
            h = hashlib.sha1((ast.unparse(n.args) + code).encode()).hexdigest()[:8]
            out.setdefault(n.name, []).append((label, mod, h, free))
    return out


def _same_constants(entries) -> bool:
    """AST が同じ 2 つが、参照する定数の**値**まで一致しているか。

    ★ これが無いと「見た目が同じで挙動が違う」を同一と誤判定する。
    """
    names = sorted({f for _l, _m, _h, fr in entries for f in fr})
    for name in names:
        vals = set()
        for _l, mod, _h, fr in entries:
            if name not in fr:
                continue
            try:
                v = getattr(importlib.import_module(mod), name)
            except Exception:
                return False
            vals.add(f"<callable>" if callable(v) else repr(v))
        if len(vals) > 1:
            return False
    return True


def test_no_identical_implementation_is_duplicated():
    """① 同じ実装が 2 箇所以上に在ってはいけない（畳んだ状態を守る）。"""
    dup = []
    for name, entries in _toplevel_defs().items():
        if len(entries) < 2:
            continue
        if len({h for _l, _m, h, _f in entries}) > 1:
            continue                                   # 実装が違う → ② の担当
        if _same_constants(entries):
            dup.append((name, [l for l, _m, _h, _f in entries]))
    assert not dup, (
        f"同じ実装が写し取られている: {dup} ── "
        "ailine_core/primitives.py へ畳み、呼ぶ側は "
        "`from ailine_core.primitives import x as _x` で引くこと")


def test_every_same_name_pair_is_explained():
    """② 同名で実装が違うものは、理由つきで台帳に在ること。"""
    unlisted = []
    for name, entries in _toplevel_defs().items():
        if len(entries) < 2:
            continue
        if len({h for _l, _m, h, _f in entries}) == 1:
            continue
        if name not in KNOWN_DIFFERENT:
            unlisted.append((name, [l for l, _m, _h, _f in entries]))
    assert not unlisted, (
        f"同名だが実装が違うのに台帳に理由が無い: {unlisted} ── "
        "畳むか、なぜ違ってよいかを KNOWN_DIFFERENT に書くこと")


def test_the_ledger_does_not_keep_stale_entries():
    """③ 同名でなくなったものが台帳に残っていたら赤（古い不安を配らない）。"""
    now = {name for name, e in _toplevel_defs().items() if len(e) >= 2}
    stale = sorted(set(KNOWN_DIFFERENT) - now)
    assert not stale, f"もう同名ではない: {stale} ── 台帳から消すこと"


def test_every_reason_is_a_sentence():
    """★ 「まだ」「あとで」で済ませない ── 理由が文になっていること。"""
    for name, why in KNOWN_DIFFERENT.items():
        assert len(why) >= 30 and "。" in why, f"KNOWN_DIFFERENT[{name}] の理由が薄い"


def test_intentional_copies_still_agree():
    """★★ 意図的な写しは、**中身が一致していること**を機械が守る。

    ★ 写しを許した以上、片方だけ直る（＝片配線）を防ぐのは番人の仕事。
      ここが赤くなったら「どちらかを直してもう片方を忘れた」ということ。
    """
    import importlib
    import inspect as _inspect
    for (mod_a, name_a, mod_b, name_b), why in INTENTIONAL_COPIES.items():
        fa = getattr(importlib.import_module(mod_a), name_a)
        fb = getattr(importlib.import_module(mod_b), name_b)
        body = lambda f: [x for x in ast.parse(
            _inspect.getsource(f).lstrip()).body[0].body
            if not (isinstance(x, ast.Expr) and isinstance(x.value, ast.Constant))]
        assert ast.unparse(ast.Module(body=body(fa), type_ignores=[])) ==                ast.unparse(ast.Module(body=body(fb), type_ignores=[])), (
            f"意図的な写しがずれた: {mod_a}.{name_a} と {mod_b}.{name_b} ── "
            f"片方だけ直した疑い。理由: {why}")


def test_every_intentional_copy_states_its_contract():
    """★ 写しを黙って増やせない（なぜ畳めないかを書く）。"""
    for key, why in INTENTIONAL_COPIES.items():
        assert "契約" in why and len(why) >= 30, f"{key} の理由が薄い: {why!r}"


def test_the_folded_helpers_still_live_in_one_place():
    """★ 畳み先が生きていること（primitives が消えたら、また写し始める）。"""
    from ailine_core import primitives
    for fn in FOLDED_INTO_PRIMITIVES:
        assert callable(getattr(primitives, fn, None)), f"primitives.{fn} が無い"
