"""「次に打つもの・書くもの」を示す所は、全部台帳に載っていること（2026-10-01・分母を広げる）。

★★ 盲検で同じ形が 6 回（嘘の案内）。導線の台帳（hints）は在ったが、分母が
  「バッククォートの中の `ailine …` / `--…`」だけで、**例文（例:『…』）と「〜と書いて」は外**だった ──
  3 体目で直したはずの「例:『品目をキーに』」が 6・7 体目で再来したとき、台帳には 1 行も無かった。

★ 分母は tests/guidance_sites_core.py が製品のソースから出す（手書きの名簿にしない）。
  ここは台帳（typable_hints_register.json の guidance 欄）と**等号で**縛る:
    新しい案内を書いた → 台帳に載せるまで赤／案内を消した → 台帳の行を消すまで赤。
★ 歩けるかは見ない ── walked 欄を歩いた結果で縛るのは test_typable_hints_are_walked（guidance も歩く）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import guidance_sites_core as g  # noqa: E402
from test_every_typable_hint_is_accounted_for import _declared, occurrences  # noqa: E402

#: ★ 2026-10-01 の初回計測で未調査 151 件。歯止めは下げる向きにだけ動かす（歩いたら下げる）。
UNWALKED_AT_FIRST_COUNT = 151


def _guide() -> list:
    return g.register().get("guidance") or []


def _key(e: dict) -> tuple:
    return (e["file"], e["text"], e.get("nth", 1))


def test_the_register_says_how_the_denominator_is_made():
    m = g.register()["_meta"]
    assert "guidance_denominator" in m and "印" in m["guidance_denominator"], m.get("guidance_denominator")


def test_every_guidance_site_is_in_the_register():
    """★★ 本体 ── 案内を書いたら台帳に載る。"""
    known = {_key(e) for e in _guide()}
    missing = sorted(f"{f}:{t[:40]}#{n}" for f, t, n, _m in g.sites() if (f, t, n) not in known)
    assert not missing, (
        f"台帳に無い案内が {len(missing)} 件あります: {missing[:6]}" + chr(10)
        + "  ★ 『こう打て／こう書け』と言ったなら、その通りにして通ることを誰かが確かめる必要があります。"
          " tests/typable_hints_register.json の guidance に足してください（scripts/guidance_sites.py --list）")


def test_the_register_does_not_keep_stale_rows():
    now = {(f, t, n) for f, t, n, _m in g.sites()}
    stale = sorted(f"{e['file']}:{e['text'][:40]}" for e in _guide() if _key(e) not in now)
    assert not stale, f"実体に無い案内が台帳に残っています: {stale[:6]}"


def test_the_two_tiers_do_not_overlap():
    """★ 同じ案内を hints と guidance の両方に載せない（片方だけ直して食い違う形を作らない）。"""
    both = {(f, t) for f, t, _n in occurrences()} & {(e["file"], e["text"]) for e in _guide()}
    assert not both, sorted(both)[:6]


def test_every_row_says_whether_anyone_walked_it():
    ok = {"walked", "path_fails", "by_design", "未調査"}
    bad = [(e["file"], e["text"][:30], e.get("walked")) for e in _guide() if e.get("walked") not in ok]
    assert not bad, bad[:6]


def test_unwalked_only_shrinks():
    """★ 未調査は『大丈夫』ではなく『見ていない』── 増えたら止まる。0 になったらこの試験は消してよい。"""
    todo = [e for e in _guide() if e.get("walked") == "未調査"]
    assert todo, "★ 未調査が 0 件 ── 歩き終えたならこの試験は消してよい"
    assert len(todo) <= UNWALKED_AT_FIRST_COUNT, (
        f"未調査が増えています（{len(todo)} 件）── 新しい案内を足したなら歩いてから")


def test_the_lies_from_the_blind_runs_are_walked():
    """★★ 盲検 7 体目で事故った形が、歩いて確かめる側に在ること（記録が「直したから消す」にならない）。"""
    rows = [e for e in _guide() if "7 体目" in (e.get("note") or "")]
    assert rows, "7 体目の案内が台帳に無い"
    assert all(e["walked"] in ("walked", "by_design") for e in rows), [
        (e["file"], e["text"][:30], e["walked"]) for e in rows]


def test_the_marks_catch_the_shapes_that_lied():
    """★ 陽性対照: 盲検で嘘になった文の形を、印が拾うこと（拾えないなら分母に穴）。
    ★ 陰性対照: 入口の名札・argv の断片は案内でない。"""
    lied = ["依頼文に列名を含めて（例:『",
            "に』）もう一度実行してください。",
            "（`--header-row` は別のシートに掛かります）。--sheet も付けてください",
            "→ 言い換えるか、2 冊の突き合わせなら: ailine run 入金.xlsx 請求.xlsx \"…\""]
    for s in lied:
        assert g.marks_of(s), s
    for s in ["■ ailine run（2冊の照合）  A=", "--sheet", "ailine accounts"]:
        assert not g.marks_of(s), s


def test_a_mark_split_across_fstring_pieces_is_still_counted():
    """★ f-string・連結で印が断片の境目に割れても、丸ごとの型で分母に入る。"""
    split = g.split_sites()
    now = {(f, t) for f, t, _n, _m in g.sites()}
    assert all((f, t.strip()) in now for f, _ln, t in split), split


def _declared_elsewhere() -> tuple:
    """★ ailine の argparse の外で宣言された名前（宣言から引く・手で並べない）:
       GUI の入口（gui/server.py）のフラグ／冊に焼く作り手の印（CREATOR_MARK・『ailine match』等）。"""
    import ast
    from _product_source import product_files
    flags, marks = set(), set()
    for f in product_files():
        tree = ast.parse(f.read_bytes().decode("utf-8"))
        for node in ast.walk(tree):
            if (isinstance(node, ast.Call) and getattr(node.func, "attr", "") == "add_argument"):
                flags |= {a.value for a in node.args
                          if isinstance(a, ast.Constant) and str(a.value).startswith("--")}
            if (isinstance(node, ast.Assign) and isinstance(node.value, ast.Constant)
                    and any(getattr(t, "id", "") == "CREATOR_MARK" for t in node.targets)):
                marks.add(str(node.value.value).split(" ", 1)[-1])
    return flags, marks


def test_every_advertised_name_in_prose_exists():
    """★ 日本語の文の中で名指ししたフラグ・サブコマンドが宣言に在ること（安い側 ── 歩けるかは別）。

    ★ 初回（2026-10-01）に 2 件鳴った ── どちらも嘘ではなく、宣言の在り処が ailine の argparse の
      外だった（GUI の `--port`・冊の作り手の印『ailine match』）。名簿を書かず、宣言から引いて足す。
    """
    import re
    subs, flags = _declared()
    more_flags, marks = _declared_elsewhere()
    flags, subs = flags | more_flags, subs | marks
    bad = []
    for f, t, _n, _m in g.sites():
        for fl in re.findall(r"(?<![`\w-])(--[a-z][a-z-]+)", t):
            if fl not in flags:
                bad.append((f, fl, t[:40]))
        for sub in re.findall(r"(?<![`\w])ailine ([a-z][a-z-]+)", t):
            if sub not in subs:
                bad.append((f, sub, t[:40]))
    assert not bad, f"実在しないものを案内しています: {bad[:6]}"
