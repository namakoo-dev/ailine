"""案内は「その通りに打てば通る」ものだけ ── 依頼者が既に書いた言い方を例に返さない（2026-10-01）。

★★ 盲検で同じ形が 6 回（嘘の案内）。ここが縛るのは 7 体目までに出た 3 つ:

  ①「キーに と書け」（3・6・7 体目で同じ文が 3 度）
      $ ailine run 請求一覧.xlsx 取引先マスタ.xlsx "取引先コードをキーにして、取引先マスタの担当者を請求一覧に転記して"
      ？ 取引先マスタ.xlsx のキー列が依頼文から決まりません。候補: 取引先コード、担当者。
        依頼文に列名を含めて（例:『取引先コードをキーに』）もう一度実行してください。
    ★ 原因: マスタ側で依頼文に出てくる文字の列が 2 本（取引先コード・担当者）── 「名指しが 2 本」で
      決まらなかった。「をキーに」と役割まで付いた名指しを、ただの名指しと同じ重さで数えていた。
    ★ そのうえマスタには金額の列が 1 本も無く、例どおりに書き足しても**必ず止まる**回だった。

  ② 転記で「`--header-row 4` のように指定して再実行」→ その通りに打っても**同じ文**で止まる
    （--header-row は run が翻訳の前に選んだ参照表に掛かり、案内の列が在るのは別のシート）。

  ③ 試算表（見出しは 4 行目）に「`--header-row 3` のように」── 3 は決め打ちの例。

★ 歩くのは製品の入口から（tests/walk_refusals_core.walk_hint・本物の ~/.ailine には書かない）。
  ★ 示した道は**画面から拾って**打つ（follow / follow_quoted）── 正解を歩き方に手で書くと、
    案内が何と言っても通ってしまう（2026-09-23 に変異が素通りした形）。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from ailine_core import header_row_advice, match  # noqa: E402
import walk_refusals_core as walk  # noqa: E402

# ---- ① 照合の「キーに と書け」 ---------------------------------------------------------

INVOICE = ["請求番号", "請求日", "取引先コード", "取引先名", "税抜金額", "消費税", "税込金額", "入金状況"]
INVOICE_ROWS = [(2, ["INV-0901", "2026/09/01", "0012", "丸山工業", 150000, 15000, 165000, "入金済"], []),
                (3, ["INV-0902", "2026/09/03", "0045", "さくら商店", 48000, 4800, 52800, "未入金"], [])]
MASTER = ["取引先コード", "取引先名", "担当者", "締日", "支払サイト"]
MASTER_ROWS = [(2, ["0012", "丸山工業", "山田", "月末", "翌月末"], []),
               (3, ["0045", "さくら商店", "佐藤", "20日", "翌月末"], [])]
TASK7 = "取引先コードをキーにして、取引先マスタの担当者を請求一覧に転記して"


def _rows(rows):
    return [(r, v, ["General"] * len(v)) for r, v, _f in rows]


def test_a_column_named_as_the_key_is_the_key():
    """★★ 根: 「取引先コード**をキーに**」と役割まで書いた名指しは、同じ依頼文の別の列（担当者）より強い。"""
    res = match.resolve_columns(TASK7, INVOICE, _rows(INVOICE_ROWS), MASTER, _rows(MASTER_ROWS))
    assert res.key_a == "取引先コード" and res.key_b == "取引先コード", res
    assert not any(role == "key" for _s, role, _c in res.unresolved), res.unresolved


@pytest.mark.parametrize("task,want", [
    ("取引先コードがキーで、担当者を見て", "取引先コード"),
    ("取引先コード をキーに担当者も", "取引先コード"),
])
def test_other_ways_of_saying_the_key(task, want):
    got, _c = match.resolve_role(task, MASTER, set(), "key", MASTER)
    assert got == want


def test_a_bare_mention_is_still_not_enough():
    """★ 陰性対照: 役割の語が付いていない名指しが 2 本なら、従来どおり決めない（勝手に選ばない）。"""
    got, cands = match.resolve_role("取引先コードと担当者を見て", MASTER, set(), "key", MASTER)
    assert got is None and set(cands) == {"取引先コード", "担当者"}


def test_a_role_word_inside_a_longer_header_is_not_a_mark():
    """★ 陰性対照: 「税込金額を金額に」の中の『金額』は『税込金額』の一部 ── 『金額』列の名指しにしない。"""
    assert match.named_with_role("税込金額を金額に", ["金額", "税込金額"], "amount",
                                 ["金額", "税込金額"]) == ["税込金額"]


def _match_walk(task, a_rows, b_rows, expect, path=None):
    return {"books": {"A.xlsx": {"表": a_rows}, "B.xlsx": {"表": b_rows}},
            "argv": ["run", "{book:A.xlsx}", "{book:B.xlsx}", task], "expect": expect,
            "path": path or {"argv": ["doctor"]}}


def test_the_seventh_request_does_not_get_its_own_words_back(tmp_path):
    """★★ 7 体目の依頼そのもの: 依頼者が書いた『取引先コードをキーに』を例として返さない。
    ★ 金額の列が無い冊なので、言い直しの例は出さずに**何が無いか**を言う（手段が無いのに例を出さない）。"""
    a = [INVOICE] + [v for _r, v, _f in INVOICE_ROWS]
    b = [MASTER] + [v for _r, v, _f in MASTER_ROWS]
    w = _match_walk(TASK7, a, b, "言い直しても決まりません")
    walk._prepare(tmp_path, w)
    with walk.isolated_home(walk.ailine, tmp_path / "_home"):
        rc, out = walk._run(walk._subst(w["argv"], tmp_path), None)
    assert rc == 3, out
    assert "取引先コードをキーに』" not in out, "★ 依頼者が既に書いた言い方を例として返している:\n" + out
    assert "例:" not in out, out
    assert "『B.xlsx』に金額に使える列（数値だけの列）がありません" in out, out


def test_the_rewording_hint_walks(tmp_path):
    """★★ 例を出す回は、画面の『…』をそのまま書き足すと通る（全部の役割ぶんが 1 行に並ぶ）。"""
    a = [["請求番号", "取引先", "税抜", "税込"], ["INV-1", "丸山", 100, 110], ["INV-2", "さくら", 200, 220]]
    b = [["請求番号", "入金額"], ["INV-1", 110], ["INV-2", 200]]
    w = _match_walk("請求番号をキーに突き合わせて", a, b, "を書き足して、もう一度実行してください",
                    {"argv": ["run", "{book:A.xlsx}", "{book:B.xlsx}", "請求番号をキーに突き合わせて"],
                     "follow_quoted": 3, "expect_rc": 0, "expect": "キー: 請求番号(A) / 請求番号(B)"})
    r = walk.walk_hint({"file": "t", "text": "t", "walk": w}, tmp_path)
    assert r["verdict"] == "walked", r


def test_the_hint_is_never_already_in_the_request():
    """★ 性質: 案内に出す言い方は、依頼文に既に在るものを含まない（在るなら案内は出さない）。"""
    res = match.resolve_columns("税抜を金額に突き合わせて", ["請求番号", "取引先", "税抜", "税込"],
                                _rows([(2, ["I1", "丸", 1, 2], [])]), ["請求番号", "入金額", "手数料"],
                                _rows([(2, ["I1", 1, 0], [])]))
    phrases = match.rewording_that_resolves("税抜を金額に突き合わせて",
                                            ["請求番号", "取引先", "税抜", "税込"],
                                            _rows([(2, ["I1", "丸", 1, 2], [])]),
                                            ["請求番号", "入金額", "手数料"],
                                            _rows([(2, ["I1", 1, 0], [])]), res.unresolved)
    assert phrases, "★ 決まらない役割が在るのに例が出ない（この検体は例を出せる形）"
    for p in phrases:
        assert p not in "税抜を金額に突き合わせて", p
    # ★ 依頼文に既に在る言い方しか出せない回は、案内そのものを出さない
    assert match.rewording_that_resolves("請求番号をキーに", ["請求番号", "取引先"], [], ["請求番号"], [],
                                         [("A", "key", ["請求番号", "取引先"])]) is None


def test_a_hint_that_would_move_a_decided_column_is_not_given():
    """★★ 書き足せば「決まる」だけでは足りない ── もう決まっていた役割を黙って別の列へ動かす言い方は勧めない。

    A 側のキーは依頼文の『取引先』で決まっている。B 側のために『請求番号をキーに』を勧めると、
    A 側も『請求番号』に替わる（通るが、依頼者の言ったキーではない）。
    """
    ha, ra = ["請求番号", "取引先", "金額"], _rows([(2, ["I1", "丸山", 100], [])])
    hb, rb = ["請求番号", "取引先名", "入金額"], _rows([(2, ["I1", "丸山", 100], [])])
    task = "取引先で突き合わせて"
    res = match.resolve_columns(task, ha, ra, hb, rb)
    assert res.key_a == "取引先" and res.key_b is None, res
    assert match.rewording_that_resolves(task, ha, ra, hb, rb, res.unresolved) is None
    # ★ 陽性対照: 動かさない言い方なら勧める（B 側に『取引先名』を勧めても A 側は動かない）
    hb2 = ["取引先名", "備考", "入金額"]
    rb2 = _rows([(2, ["丸山", "x", 100], [])])
    res2 = match.resolve_columns(task, ha, ra, hb2, rb2)
    assert match.rewording_that_resolves(task, ha, ra, hb2, rb2, res2.unresolved) == ["取引先名をキーに"]


# ---- ②③ --header-row の案内 -----------------------------------------------------------

#: 7 体目の試算表の形（StructDump の行の特徴）: 1・2 行目はタイトル、3 行目は空、4 行目が見出し
TRIAL_BALANCE = {1: {"str": 1, "nonempty": 1}, 2: {"str": 1, "nonempty": 1},
                 4: {"str": 5, "nonempty": 5}, 5: {"str": 1, "nonempty": 1},
                 6: {"str": 1, "nonempty": 5}, 9: {"str": 2, "nonempty": 2},
                 12: {"str": 2, "nonempty": 2}}


def test_the_example_row_is_the_row_the_machine_saw():
    """★★ ③: 例の行番号は決め打ちの 3 でなく、見出しらしく見えた行（幅が最大）。"""
    q = header_row_advice.clarify_question(TRIAL_BALANCE)
    assert "`--header-row 4`" in q, q
    assert "--header-row 3" not in q, q


def test_no_row_number_is_guessed_when_nothing_looks_like_a_header():
    """★ 見出しらしい行が無いなら、行番号を当て推量で出さない。"""
    q = header_row_advice.clarify_question({1: {"str": 1, "nonempty": 2}})
    assert q == header_row_advice.NO_CANDIDATE_QUESTION
    assert not any(ch.isdigit() for ch in q.split("`")[1]), q


def test_the_sheet_is_named_when_the_header_row_would_land_elsewhere():
    """★★ ②: --header-row が掛かるシートと列の在るシートが違えば、--sheet も 1 行に並べる。"""
    h = header_row_advice.missing_column_hint("取引先コード", 4, "2026年9月", "取引先マスタ")
    assert "`--sheet 2026年9月 --header-row 4`" in h, h
    same = header_row_advice.missing_column_hint("取引先コード", 4, "2026年9月", "2026年9月")
    assert "--sheet" not in same and "`--header-row 4`" in same, same


def test_a_sheet_name_with_a_space_is_typable():
    h = header_row_advice.missing_column_hint("コード", 2, "9月 請求", "マスタ")
    assert '`--sheet "9月 請求" --header-row 2`' in h, h


def test_the_seventh_transfer_hint_walks(tmp_path):
    """★★ ② を入口から: 案内の `--sheet … --header-row …` を画面から拾って打つと通る。"""
    w = {"books": {"請求.xlsx": {
            "請求": [["請求一覧（9月分）"], ["作成：山田"], [], ["請求番号", "取引先コード", "金額"],
                     ["INV-1", "0012", 100], ["INV-2", "0045", 200]],
            "取引先マスタ": [["取引先コード", "担当者"], ["0012", "山田"], ["0045", "佐藤"]]}},
         "argv": ["run", "{book:請求.xlsx}", "取引先マスタのシートから取引先コードで担当者を引いてきて", "--copy"],
         "plan": [{"op": "LOOKUP_FILL", "args": {"target_sheet": "請求", "target_col": "担当者",
                                                  "source_sheet": "取引先マスタ", "key_col": "取引先コード"}}],
         "expect": "シートの4行目に見出しがあるようです。",
         "path": {"argv": ["run", "{book:請求.xlsx}", "取引先マスタのシートから取引先コードで担当者を引いてきて",
                           "--copy", "--dry"], "follow": "--sheet", "expect_rc": 0, "expect": "（--dry"}}
    r = walk.walk_hint({"file": "t", "text": "t", "walk": w}, tmp_path)
    assert r["verdict"] == "walked", r
