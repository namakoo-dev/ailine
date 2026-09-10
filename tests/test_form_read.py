# -*- coding: utf-8 -*-
"""帳票を読む規則の番人（2026-09-11）。

★★ ここに在る試験は**全部、実測で踏んだ事故**から来ている。
  検体 v2（87 冊）と、一度も測っていない実物の雛形 26 冊で出たものだけを置く。
  ★ 想像で書いた「請求書はこうだろう」は 1 本も入れない ── それは俺の頭の中の
    請求書を守るだけで、買い手の帳票を守らない。

★ 検体そのものは repo に入らない（他社の帳票から作っているため）。
  だからここでは openpyxl でその場に**最小の帳票**を組み立てて縛る。
"""
from __future__ import annotations

import openpyxl
import pytest

from ailine_core.field_record import CONFIRMED, NONE_FOUND, SINGLE, SPLIT, grade, value
from ailine_core.form_grid import Grid
from ailine_core.form_read import clean_org_name, read_form


def book(rows: dict, merges=()):
    """{"B3": 値, …} から 1 枚の帳票を組む。戻り値: (値のシート, 式のシート)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    for at, v in rows.items():
        ws[at] = v
    for m in merges:
        ws.merge_cells(m)
    return ws, ws


#: 使い回す最小の帳票 ── 上部の請求額・明細 2 行・帯（積み上げ型）。
def minimal(**over):
    rows = {
        "B2": "請求書",
        "B3": "ナギ商会株式会社",
        "B4": "経理部　御中",
        "G5": "あかね商事株式会社",
        "G6": "〒000-0000",
        "B11": " ご請求金額　", "C11": 3300, "D11": "（税込）",
        "B15": "品番・品名", "E15": "数量", "G15": "単価", "H15": "金額",
        "B16": "用紙代", "E16": 1, "G16": 1000, "H16": 1000,
        "B17": "運送費", "E17": 2, "G17": 1000, "H17": 2000,
        "E37": "小計", "H37": 3000,
        "E38": "消費税", "G38": 0.1, "H38": 300,
        "E39": "合計金額", "H39": 3300,
    }
    rows.update(over)
    return rows


def read(rows, merges=()):
    ws, wsf = book(rows, merges)
    return read_form(ws, ws_formula=wsf)


# ── ★ 名前を作る道は 1 本（片配線を塞ぐ） ──────────────────────
@pytest.mark.parametrize("field, cell", [("宛先", "B3"), ("請求元", "G5")])
def test_a_placeholder_name_is_refused_for_both_sides(field, cell):
    """★★ 1 本の試験で**宛先と請求元の両方**を縛る。

    ★ 2026-09-11、一度も測っていない実物の雛形 15 冊で踏んだ:
      プレースホルダ（`○○株式会社`）の番人を**請求元の側にだけ**書いていて、
      宛先は空の雛形に対して「宛先＝○○株式会社」を**値として出していた**。
    ★ 処方は「両方直す」ではなく「1 関数に畳んで呼び出し側に選ばせない」。
      だからこの試験も 1 本にする ── `clean_org_name` を壊せば**必ず両方赤くなる**。
    """
    recs = read(minimal(**{cell: "○○株式会社"}))
    r = recs[field]
    assert grade(r) == NONE_FOUND, f"★ 雛形のままの名前を {field} の値にした"
    assert "雛形" in r.blank_reason, f"★ 理由が雛形だと言っていない: {r.blank_reason}"
    assert "○○" in r.blank_reason, "★ 見つけたものを引用していない（人が直せない）"


def test_clean_org_name_is_the_only_door():
    """★ 名前を作る関数そのものの契約（両側がこれを通る）。"""
    assert clean_org_name("あかね商事株式会社") == ("あかね商事株式会社", "")
    assert clean_org_name("ナギ商会株式会社　御中")[0] == "ナギ商会株式会社"
    # ★ 1 セルに登録番号が改行で同居する（実物）
    assert clean_org_name("高梨産業株式会社\n(登録番号:T7010001234567)")[0] \
        == "高梨産業株式会社"
    # ★ 取次・担当の行は名前の行ではない
    assert clean_org_name("（取次）中央商事株式会社　担当 井上")[0] == ""
    assert clean_org_name("株式会社 〇〇〇")[0] == ""


# ── 宛先: 敬称の付いたセルが名前とは限らない ────────────────────
def test_the_honorific_cell_may_be_a_department_and_the_name_is_above():
    """★ `B3=社名` / `B4=経理部　御中` ── 御中が付くのは部署の側。

    ★ 2026-09-10 に「宛先 15/16 取れた」と誤報したのは、この御中セルを
      そのまま宛先として読んでいたから。
    """
    r = read(minimal())["宛先"]
    assert value(r) == "ナギ商会株式会社", f"★ 部署名を宛先にした: {value(r)!r}"


def test_the_honorific_may_live_in_the_name_cell_itself():
    """★ `B7=社名　御中` の骨もある。**両方の形が要る。**"""
    rows = minimal()
    del rows["B3"], rows["B4"]
    rows["B7"] = "ナギ商会株式会社　御中"
    r = read(rows)["宛先"]
    assert value(r) == "ナギ商会株式会社"


def test_a_gap_row_between_the_name_and_the_honorific_is_crossed():
    """★ 名前と御中のあいだが空くことがある（骨によって 1〜2 行）。"""
    rows = minimal()
    del rows["B4"]
    rows["B6"] = "総務課　御中"
    r = read(rows)["宛先"]
    assert value(r) == "ナギ商会株式会社"


# ── 請求額: ラベルの右をどう読むか ─────────────────────────────
def test_the_tax_rate_between_the_label_and_the_amount_is_not_the_amount():
    """★ `E38=消費税 | G38=0.1 | H38=300` ── いちばん近い右は**税率**。"""
    ws, wsf = book(minimal())
    g = Grid.read(ws)
    from ailine_core.form_read import _labelled_number, _LABEL_TAX
    got = _labelled_number(g, _LABEL_TAX, rows=(15, 60))
    assert got and got[0][1].value == 300, "★ 0.1（税率）を消費税額として拾った"


def test_a_money_amount_typed_as_text_is_not_silently_skipped():
    """★★ 答えの枠に文字で金額が入っていたら、**跨いで別の数を読まない**。

    ★ 2026-09-11 の実測: `¥1,320,000-` や `#REF!` を黙って跨ぎ、
      別のセルの数を「確」として出していた。
      黙って別のものを読むのが、この repo でいちばん高くつく失敗。
    """
    r = read(minimal(C11="¥1,320,000-"))["請求額"]
    assert grade(r) not in (CONFIRMED, SINGLE), "★ 別の数を黙って読んだ"
    assert "文字" in r.blank_reason or "読み取れ" in r.blank_reason


def test_a_broken_formula_in_the_amount_slot_is_reported():
    r = read(minimal(C11="#REF!"))["請求額"]
    assert grade(r) == NONE_FOUND
    assert "#REF!" in r.blank_reason


# ── ★ 一致だけで確と言わない・食い違いは区分に届く ────────────────
def test_the_detail_block_is_a_third_mouth():
    """明細の合計が帯と合っていれば、そこで初めて『確』が立つ。"""
    r = read(minimal())["請求額"]
    assert grade(r) == CONFIRMED
    assert value(r) == 3300
    assert "明細" in r.swept_how, f"★ 何を掃き出したか言っていない: {r.swept_how}"


def test_a_detail_that_disagrees_with_the_subtotal_blanks_the_amount():
    """★ 上部と帯が一致していても、明細と合わないなら値は出さない。

    ★ 集計範囲の外に明細が追記されている冊・SUMIF が行を落とす冊は、
      **上部と帯だけ見ていると一致していて気づけない**。
    """
    r = read(minimal(H17=9999))["請求額"]
    assert grade(r) == SPLIT, "★ 明細と小計の食い違いが区分に届いていない（片配線）"
    assert "明細" in r.blank_reason and "小計" in r.blank_reason
    assert value(r) is None


def test_a_row_whose_quantity_is_missing_is_caught_even_when_the_sum_agrees():
    """★★ 0 を足しても SUM は合う ── 満場一致で壊れた冊が通る形。

    ★ 数量の入力漏れで `単価 1,000 × 数量(空) = 金額 0`。
      小計・合計・上部の 3 つが揃って一致するので、行の中を見ないと分からない。
    """
    rows = minimal()
    rows["H17"] = 0
    del rows["E17"]                      # 数量が空
    rows["H37"] = 1000
    rows["H38"] = 100
    rows["H39"] = 1100
    rows["C11"] = 1100
    r = read(rows)["請求額"]
    assert grade(r) == SPLIT, "★ 数量が空で金額 0 の行を見逃した"
    assert "数量" in r.blank_reason, f"★ 何が悪いか言っていない: {r.blank_reason}"


def test_a_carried_forward_block_makes_the_amount_undecidable():
    """★ 繰越請求・源泉徴収がある帳票では、合計は**振り込む額ではない**。

    ★ 実物は `お振込金額　121,790 円` のように**ラベルと金額が同じセル**に入る。
    """
    r = read(minimal(B41="源泉所得税　▲10,210", B42="お振込金額　121,790 円"))["請求額"]
    assert grade(r) == SPLIT
    assert "源泉" in r.blank_reason, f"★ 見つけた欄を名指していない: {r.blank_reason}"


def test_the_workbooks_own_formula_is_not_an_independent_check():
    """★★ 合計が `=H37+H38` なら、小計＋消費税との一致は**検算ではない**。

    ★ 2026-09-10 の恒真（自分の写しと一致して裏が取れた）と同じ形。
    """
    ws = openpyxl.Workbook().active
    for at, v in minimal().items():
        ws[at] = v
    wsf = openpyxl.Workbook().active
    for at, v in minimal().items():
        wsf[at] = v
    wsf["H39"] = "=H37+H38"              # ★ 合計は小計と消費税の式そのもの
    recs = read_form(ws, ws_formula=wsf)
    r = recs["請求額"]
    ats = {e.at for e in r.evidences}
    assert "H37+H38" not in ats, "★ 帳票自身の式を 2 つ目の根拠として数えた（恒真）"
    assert any("検算になりません" in why for _at, _v, why in r.excluded), \
        "★ 数えなかったことを人に見せていない"


# ── 帯の形が 2 つある ────────────────────────────────────────
def test_the_band_can_be_a_table_where_the_tax_sits_in_another_column():
    """★ `C28=小計 | E28=金額` / `F26=消費税 | G26=…` ── 列が揃わない帯。

    ★ 初版は「消費税ラベルの右」を取り、**税区分の 10% 欄**を小計に足して
      偽の食い違いを出した（実測 R01/R02）。小計と同じ**行**の消費税列を見る。
    """
    rows = {
        "B1": "請求書",
        "B7": "ナギ商会株式会社　御中",
        "F2": "あかね商事株式会社", "F3": "〒000-0000",
        "B14": "ご請求金額", "E14": 153000,
        "C16": "品目", "E16": "単価", "F16": "数量", "G16": "金額",
        "B17": "用紙代", "E17": 90000, "F17": 1, "G17": 90000,
        "B18": "食品", "E18": 50000, "F18": 1, "G18": 50000,
        "C26": "10%対象", "D26": "対象額（税抜）", "E26": 90000,
        "F26": "消費税", "G26": 9000,
        "C27": "8%対象(※)", "E27": 50000, "G27": 4000,
        "C28": "小計", "E28": 140000, "G28": 13000,
    }
    r = read(rows)["請求額"]
    assert grade(r) != SPLIT, f"★ 表型の帯で偽の食い違いを出した: {r.blank_reason}"
    assert value(r) == 153000


def test_a_sheet_that_is_not_an_invoice_says_nothing():
    """★ 請求書でないものに答えない（実物の政府統計 6 冊で確認した振る舞い）。"""
    recs = read({"A1": "家計調査", "A3": "世帯数", "B3": 1234, "A4": "金額", "B4": 5678})
    for fld, r in recs.items():
        assert grade(r) == NONE_FOUND, f"★ 請求書でないのに {fld} を答えた: {value(r)!r}"
        assert r.blank_reason, "★ 空欄なのに理由が無い"
