"""条件が 2 つある依頼を、**半分だけやって ✓ を出さない**こと（2026-09-06）。

★★ 出所（実測）: 「金額が1000以上**で部門が営業**の行の備考に『○』を付けて」で ──

    表   ボルト/営業/1500 ・ ナット/経理/9000 ・ ワッシャ/営業/2000 ・ ねじ/経理/300
    正   2 行目(営業/1500) と 4 行目(営業/2000) だけ
    実際 [2, 3, 4] ── **3 行目(経理/9000) にも ○ が付き、ok=True**

  ★ **「部門が営業」という条件が丸ごと無視され、しかも ✓ が出る。**
    宣言（金額の条件）と実体は一致するので、事後条件は正しく通る ──
    欠けていたのは「**依頼 vs 宣言**」で、今日の午前に否定で直したのと同じ形。

★★ 最初に書いた検体は**差が出なかった**（打ち消し合っていた）:
    ボルト/営業/1500・ナット/経理/800・ワッシャ/営業/2000
    → 経理の行は金額でも外れるので、部門を無視しても答えが同じになる。
  ★ **条件の片方だけを満たす行**を必ず入れる ── それが無い検体は、何も測っていない。

★ この束は**実装より先に**書いた（Namakoo の指示）。書いた時点では「半分だけやる」が
  再現し、実装した瞬間に xfail が XPASS で赤くなって印を外した。

★★ 直し方（2026-09-06・実装した形）:
  ・Basic の `SetColumnValueWhere` に **Optional で 2 組目**を足し、`RowMatches` を
    **2 回呼んで And** する ── 述語（Python の別実装・凍結した真理表）は**無傷**。
  ・2 組目は**残差**（1 組目が消費しなかった語）から採る。
    ★ これが要る理由は既存の検体が教えた: 「売上が700以上の行のチェック列に『◎』」で
      **閾値の 700 が原価列にも在る**ため「原価が700」を 2 組目に採ってしまった。
  ・値は**実表に在るもの**だけ（`task_names_real_values`）── だから比較は `eq` で足り、
    `contains` の曖昧さに踏み込まない。
  ・**1 列に絞れた時だけ**採る。3 条件の依頼は 2 列に当たるので自然に外れ、
    その回は「まず抽出で絞ってください」と道を案内する（行き止まりにしない）。
"""
import re
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402
from ailine_core import residue  # noqa: E402

HEADERS = ["品名", "部門", "金額", "備考"]
#: ★ 片方だけ満たす行（ナット=金額○/部門×、ねじ=金額×/部門×）を必ず含める
ROWS = [("ボルト", "営業", 1500, ""), ("ナット", "経理", 9000, ""),
        ("ワッシャ", "営業", 2000, ""), ("ねじ", "経理", 300, "")]
TASK = "金額が1000以上で部門が営業の行の備考に「○」を付けて"
#: 正しい答え（1 起点の行番号・見出しは 1 行目）
WANT_ROWS = [2, 4]


@pytest.fixture
def book(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r)
    p = tmp_path / "t.xlsx"
    wb.save(p)
    return p


def _resolve(book, task=TASK):
    bm = ailine.build_book_meta(book)
    return ailine.verify_dsl_args(
        "SET_WHERE", {"col": "備考", "cond_col": "金額", "cmp": "gte", "value": "○"},
        bm, task=task, vocab=ailine.load_vocab())


def test_the_specimen_can_tell_the_two_apart(book):
    """★ 検体の妥当性を先に確かめる ── 片方だけ満たす行が在ること。

    これが無いと、条件を無視しても答えが同じになり**何も測れない**（最初に踏んだ）。
    """
    only_amount = [i for i, (_n, d, a, _b) in enumerate(ROWS, start=2)
                   if a >= 1000 and d != "営業"]
    assert only_amount, "金額だけ満たす行が無い ── この検体では違いが出ない"
    only_dept = [i for i, (_n, d, a, _b) in enumerate(ROWS, start=2)
                 if a < 1000 and d == "営業"]
    del only_dept          # ★ 片側だけでも足りる（両方あればなお良い）


def test_the_machine_knows_the_words_it_did_not_use(book):
    """★ 材料は既に在る ── 使わなかった条件が**残差**に出ている。"""
    pool = [x for op in ailine.OP_META for x in ailine._op_match_pool(op) if x]
    used = {"col": "備考", "cond_col": "金額", "cmp": "gte",
            "cond_value": 1000, "value": "○"}
    left = residue.find_unconsumed_words(TASK, used, pool)
    assert "部門" in left and "営業" in left, left


def test_both_conditions_are_honoured(book):
    """★★ 本命 ── 2 つの条件が**両方**効くこと。

    ★ この試験は最初 `strict xfail`（未対応の印）で書いた。実装した瞬間 **XPASS で
      赤くなり**、印を外して期待値を正の向きにした ── 「番人にはいつ外してよいかを
      書く」が、書いた当人に同じ日に効いた実例。
    """
    ok, res, _inf, err = _resolve(book)
    assert ok, err
    assert res.get("_match_rows") == WANT_ROWS, res.get("_match_rows")


# --- 3 条件以上: 断って**道を案内する**（行き止まりにしない）-------------------

@pytest.fixture
def book3(tmp_path):
    """3 条件の依頼が当たる表（部門・担当の 2 列に値が当たる）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["品名", "部門", "担当", "金額", "備考"])
    for r in [("ボルト", "営業", "田中", 1500, ""), ("ナット", "経理", "鈴木", 9000, ""),
              ("ワッシャ", "営業", "佐藤", 2000, ""), ("ねじ", "営業", "田中", 300, "")]:
        ws.append(r)
    p = tmp_path / "t3.xlsx"
    wb.save(p)
    return p


TASK3 = "金額が1000以上で部門が営業で担当が田中の行の備考に「○」を付けて"


def test_three_conditions_are_refused_not_half_done(book3):
    """★ 半分だけやらない ── 3 つ目が決められないなら**断る**。"""
    bm = ailine.build_book_meta(book3)
    ok, _res, _inf, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "備考", "cond_col": "金額", "cmp": "gte", "value": "○"},
        bm, task=TASK3, vocab=ailine.load_vocab())
    assert not ok, "3 条件を黙って半分やっている"
    assert "3 つ以上" in err and "2 つまで" in err, err


def test_the_route_is_built_only_from_the_real_table(book3):
    """★★ Namakoo の懸念（「まず〜」にでたらめが入らないか）を機械で縛る。

    案内に出る **列名・値・行数がすべて実表と一致**すること ── 依頼文に在るだけの語や、
    実表に無い値は出せない。
    """
    bm = ailine.build_book_meta(book3)
    _ok, _res, _inf, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "備考", "cond_col": "金額", "cmp": "gte", "value": "○"},
        bm, task=TASK3, vocab=ailine.load_vocab())
    wb = openpyxl.load_workbook(book3)
    ws = wb["売上"]
    heads = [str(ws.cell(1, c).value or "") for c in range(1, ws.max_column + 1)]
    body = {h: [ws.cell(r, c + 1).value for r in range(2, ws.max_row + 1)]
            for c, h in enumerate(heads)}
    wb.close()
    for m in re.finditer(r"「(.+?)が(.+?)の行を抜き出して」（(\d+) 行）", err):
        col, val, cnt = m.group(1), m.group(2), int(m.group(3))
        assert col in heads, f"実表に無い列を案内している: {col}"
        assert val in [str(v) for v in body[col]], f"実表に無い値を案内している: {val}"
        assert cnt == sum(1 for v in body[col] if str(v) == val), (col, val, cnt)
    assert "の行を抜き出して" in err, err


@pytest.mark.local
def test_the_route_actually_runs(book3, tmp_path):
    """★ 案内した文が**そのまま通る**こと（昨日『示した例が通らない』を踏んだ形）。"""
    import os
    import subprocess
    bm = ailine.build_book_meta(book3)
    _ok, _res, _inf, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "備考", "cond_col": "金額", "cmp": "gte", "value": "○"},
        bm, task=TASK3, vocab=ailine.load_vocab())
    shown = re.findall(r"「(.+?の行を抜き出して)」", err)
    assert shown, err
    env = {**os.environ, "PYTHONPATH": str(REPO / "src")}
    got = subprocess.run([sys.executable, "-m", "ailine", "run", str(book3), shown[0],
                          "--copy", "--timeout", "200"],
                         capture_output=True, text=True, encoding="utf-8",
                         errors="replace", cwd=str(REPO), env=env)
    assert got.returncode == 0, f"案内した文が通らない: 「{shown[0]}」{chr(10)}{got.stdout[-500:]}"


# --- 誤爆させない（★ 既存の検体が捕まえた形を、自分の束にも持つ）---------------

def test_a_threshold_is_not_reused_as_a_second_condition(tmp_path):
    """★★ 2026-09-06 に実際に踏んだ: 「売上が700以上の行のチェック列に『◎』」で、
      **閾値の 700 が原価列にも在る**ため「原価が700」を 2 組目に採り、行が減った。
      ★ 1 組目が消費した語は 2 組目の材料にしない（残差だけを見る）。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for r in [["商品", "売上", "原価", "チェック"], ["りんご", 1200, 700, None],
              ["みかん", 800, 300, None]]:
        ws.append(r)
    p = tmp_path / "th.xlsx"
    wb.save(p)
    bm = ailine.build_book_meta(p)
    ok, res, _inf, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "チェック", "cond_col": "売上", "cmp": "gte"}, bm,
        task="売上が700以上の行のチェック列に「◎」を付けて", vocab=ailine.load_vocab())
    assert ok, err
    assert res.get("cond2_col") is None, res.get("cond2_col")
    assert res["_match_rows"] == [2, 3], res["_match_rows"]


def test_a_one_character_value_cannot_be_a_second_condition(tmp_path):
    """★ **既知の限界**を明示して持つ（2026-09-06 に測って分かった）。

    2 組目の材料は**残差**（1 組目が消費しなかった語）から採るが、残差の抽出は
    **2 文字以上**しか拾わない（1 文字を拾うと助詞や記号を大量に掴む設計）。
    だから『主』『甲』のような **1 文字の値**は 2 組目に採れない ── その回は
    従来どおり 1 条件で走る（★ 半分やることになるので、ここは**穴として残る**）。

    ★ 直すなら residue 側の意味論を変えることになり、4 つの呼び出し全部に影響する。
      いまは**扱えないと知っている**方を選ぶ（黙って持たない）。
    ★ この試験が赤くなったら「1 文字も採れるようになった」の意味 ── 期待値を反転して commit。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名簿"
    ws.append(["氏名", "所属", "担当", "メモ"])
    for r in [("田中", "営業", "主", ""), ("鈴木", "経理", "副", ""),
              ("佐藤", "営業", "副", ""), ("山田", "総務", "主", "")]:
        ws.append(r)
    p = tmp_path / "one.xlsx"
    wb.save(p)
    bm = ailine.build_book_meta(p)
    _ok, res, _inf, _err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "メモ", "cond_col": "所属", "cmp": "eq", "value": "○"},
        bm, task="所属が経理で担当が主の行のメモに「○」を付けて",
        vocab=ailine.load_vocab())
    assert res.get("cond2_col") is None, "1 文字の値が採れるようになった（限界が動いた）"

    # ★ 対で縛る: **2 文字**なら同じ形で採れる（限界が「1 文字」であることの証明）
    wb2 = openpyxl.load_workbook(p)
    ws2 = wb2["名簿"]
    for r, v in zip(range(2, 6), ("主任", "副任", "副任", "主任")):
        ws2.cell(r, 3).value = v
    q = tmp_path / "two.xlsx"
    wb2.save(q)
    wb2.close()
    bm2 = ailine.build_book_meta(q)
    _ok2, res2, _i2, _e2 = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "メモ", "cond_col": "所属", "cmp": "eq", "value": "○"},
        bm2, task="所属が経理で担当が主任の行のメモに「○」を付けて",
        vocab=ailine.load_vocab())
    assert res2.get("cond2_col") == "担当" and res2.get("cond2_value") == "主任", res2


def test_two_conditions_hold_together_with_the_negation(tmp_path):
    """★ 否定（〜以外）と 2 組目が**同時に**効くこと（今日の午前と午後の合流点）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名簿"
    ws.append(["氏名", "所属", "担当", "メモ"])
    for r in [("田中", "営業", "主任", ""), ("鈴木", "経理", "副任", ""),
              ("佐藤", "営業", "副任", ""), ("山田", "総務", "主任", "")]:
        ws.append(r)
    p = tmp_path / "neg2.xlsx"
    wb.save(p)
    bm = ailine.build_book_meta(p)
    ok, res, _inf, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "メモ", "cond_col": "所属", "cmp": "eq", "value": "○"},
        bm, task="所属が営業以外で担当が主任の行のメモに「○」を付けて",
        vocab=ailine.load_vocab())
    assert ok, err
    assert res["cmp"] == "nin" and res.get("cond2_col") == "担当"
    # 営業でない かつ 主任 ＝ 山田（5 行目）だけ
    assert res["_match_rows"] == [5], res["_match_rows"]


def test_no_row_matches_is_refused_before_writing(tmp_path):
    """★ 2 条件で該当が 0 行なら、**書かずに断る**こと。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名簿"
    ws.append(["氏名", "所属", "担当", "メモ"])
    for r in [("田中", "営業", "主任", ""), ("鈴木", "経理", "副任", "")]:
        ws.append(r)
    p = tmp_path / "none.xlsx"
    wb.save(p)
    ok, _res, _inf, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "メモ", "cond_col": "所属", "cmp": "eq", "value": "○"},
        ailine.build_book_meta(p),
        task="所属が経理で担当が主任の行のメモに「○」を付けて", vocab=ailine.load_vocab())
    assert not ok
    assert "当てはまる行がありません" in err and "何も書いていません" in err, err


def test_the_single_condition_path_is_untouched(book):
    """★ 対で縛る ── 条件が 1 つの依頼は 1 ビットも変わらないこと。"""
    ok, res, _inf, err = _resolve(book, "金額が1000以上の行の備考に「○」を付けて")
    assert ok, err
    assert res.get("cond2_col") is None
    assert res["_match_rows"] == [2, 3, 4], res["_match_rows"]
    code = ailine.codegen_dsl("SET_WHERE", res, book_meta=ailine.build_book_meta(book))
    # ★ 字面でなく**引数の個数**で見る（今日 4 本の番人が字面で落ちた）
    call = [ln for ln in code.splitlines() if "SetColumnValueWhere" in ln][0]
    assert call.count(",") == 7, f"1 条件なのに引数が増えている: {call}"


def test_the_postcondition_counts_both_conditions(tmp_path):
    """★★ 検算の側も 2 条件で数えること（変異試験がすり抜けて分かった）。

    ★ 解決（verify_dsl_args）までの検体では、事後条件の穴が見えない ──
      **前後 2 冊を作って直接叩く**。ここが 1 条件のままだと、Basic が正しく 2 行だけ
      書いても「当てはまるのに書かれていない行がある」と落ちる（または逆を通す）。
    ★ 本体（Basic）と検算（Python）は**別実装**のまま ── どちらも述語を 2 回呼ぶだけ。
    """
    from ailine_core.postconditions import shape

    def mk(path, marks):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "売上"
        ws.append(HEADERS)
        for (n, d, a, _b), m in zip(ROWS, marks):
            ws.append([n, d, a, m])
        wb.save(path)
        return path

    before = mk(tmp_path / "b.xlsx", ["", "", "", ""])
    # ★ 正しい適用: 営業 かつ 1000 以上 の 2 行だけに ○
    after = mk(tmp_path / "a.xlsx", ["○", "", "○", ""])
    args = {"col": "備考", "cond_col": "金額", "cmp": "gte", "cond_value": 1000,
            "value": "○", "cond2_col": "部門", "cond2_cmp": "eq", "cond2_value": "営業",
            "_target_sheet": "売上", "_header_row": 1}
    st, msg = shape.check_set_where(after, args, header_row=1, source_book=before)
    assert st == "pass", (st, msg)

    # ★ 逆に「金額だけ」で書いた（＝2 組目を無視した）回は落ちること
    half = mk(tmp_path / "h.xlsx", ["○", "○", "○", ""])
    st2, msg2 = shape.check_set_where(half, args, header_row=1, source_book=before)
    assert st2 == "fail", (st2, msg2)
    assert "変わっています" in msg2 or "広がった" in msg2, msg2


def test_the_second_condition_reaches_basic(book):
    """★ 2 組目が Basic に**渡る**こと（引数の個数で見る）。"""
    _ok, res, _inf, _err = _resolve(book)
    code = ailine.codegen_dsl("SET_WHERE", res, book_meta=ailine.build_book_meta(book))
    call = [ln for ln in code.splitlines() if "SetColumnValueWhere" in ln][0]
    assert call.count(",") == 10, f"2 組目が渡っていない: {call}"


def test_it_does_not_silently_do_half(book):
    """★ せめて「半分やった」と分かること（断るか・言うか・両方やるか）。

    ★ いまは**何も言わずに半分やる**ので、この試験も落ちる ── 直し方は 3 通りあるが、
      どれを選んでも「黙って半分」だけは無くなる。
    """
    ok, res, _inf, err = _resolve(book)
    rows = res.get("_match_rows")
    said = " ".join(res.get("_warnings") or [])
    honoured = rows == WANT_ROWS
    disclosed = ("部門" in said) or (not ok and "部門" in str(err))
    assert honoured or disclosed, (
        f"2 つ目の条件を黙って捨てている: 行={rows} 警告={said!r} err={err!r}")
