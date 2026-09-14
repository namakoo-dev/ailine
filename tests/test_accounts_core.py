# -*- coding: utf-8 -*-
"""需要③『経費の勘定科目を先例から引く』の純関数の検体（2026-09-13）。

契約（設計 docs/DESIGN-20260913-経費の勘定科目を先例から引く.md ★ §6 が §2 を置き換える）:
  - 候補を出す行は「借方勘定科目が空 かつ 借方金額が在る」行だけ。複合仕訳の継続行
    （借方が丸ごと空）と埋まっている行は触らず、番号と理由を控える（§6.2）
  - 出所は**鍵ごと**（4 本: 借方取引先 / 貸方取引先 / 借方補助科目 / 摘要）── 同じ鍵の
    過去 N 件は **1 出所**（1 つの入力の N 度刷り・§6.1）
  - 区分は `field_record.grade_of` **だけ**が決める（この試験も語を `field_record` から引く）
  - 今回の冊の埋まった行は先例に数えない（自己先例は『裏が取れた』の最短路・§6.3）
  - 鍵が 1 件も当たらなければ断る（全行「無」の静かな成功をしない・§6.3）
  - 列は別名で当てる。0 列なら見た見出しを並べて断り、2 列に当たっても断る（§6.2）
  - 表記ゆれは `split_people.lookalike_pairs` で**名指しだけ**（併合しない）

★★ 変異で赤くなることを確かめた（2026-09-13・手で当てて確認・報告に記録）:
    単を確に格上げする版                     → test_one_key_stays_single が赤
    同じ鍵の N 件を出所 N つに数える版         → 同上（1 本の鍵で 確 になる）
    継続行にも候補を出す版                   → test_a_continuation_row_is_never_touched が赤
    今回の埋まった行を先例に数える版           → test_todays_filled_rows_are_not_precedents が赤

★ LLM も LibreOffice も openpyxl も要らない（純関数・値だけ）。
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core import accounts_core, field_record   # noqa: E402

#: マネーフォワードの書き出しを模した見出し（★ 実物の別名をそのまま使う）。
MF = ["取引No", "取引日", "借方勘定科目", "借方補助科目", "借方取引先", "借方金額(円)",
      "貸方勘定科目", "貸方取引先", "摘要"]


def _row(no="", date="", account="", sub="", partner="", amount="", credit="現金",
         credit_partner="", memo=""):
    return [no, date, account, sub, partner, amount, credit, credit_partner, memo]


def _grid(rows, headers=MF):
    """見出し 1 行 ＋ データ行の (行番号, 値) の並び（★ 行番号は物理行・1 起点）。"""
    out = [(1, list(headers))]
    for i, values in enumerate(rows, start=2):
        out.append((i, list(values)))
    return out


def _book(rows, headers=MF):
    """(header_map, データ行) ── 列の解決まで通す（断ったら AssertionError）。"""
    grid = _grid(rows, headers)
    head_row, _headers, header_map, refusal = accounts_core.resolve_accounts_columns(grid)
    assert refusal is None, refusal
    data = [(r, v) for r, v in grid if head_row is None or r > head_row]
    return header_map, data


def _past(rows, headers=MF, name="過去.csv"):
    header_map, data = _book(rows, headers)
    return {name: {"header_map": header_map, "rows": data}}


def _plan(today_rows, past_rows, headers=MF, past_headers=None):
    header_map, today = _book(today_rows, headers)
    return accounts_core.plan_accounts(today, header_map,
                                       _past(past_rows, past_headers or headers))


def _grade(plan, row):
    return field_record.grade(plan.records[row])


def _value(plan, row):
    return field_record.value(plan.records[row])


# --- 鍵そのもの ------------------------------------------------------------------

def test_the_four_keys_are_frozen():
    """★ 鍵は 4 本（貸方の勘定科目＝払い方は入っていない ── 設計 §6.1）。"""
    assert accounts_core.KEYS == ("借方取引先", "貸方取引先", "借方補助科目", "摘要")
    assert "貸方勘定科目" not in accounts_core.KEYS


# --- 4 つの区分 ------------------------------------------------------------------

def test_two_keys_from_two_different_past_rows_is_confirmed():
    """確: 2 本の鍵が**別々の過去の行**で当たり、科目が一致する。"""
    plan = _plan(
        [_row("11", "2026/08/03", "", "携帯", "甲通信", "8800", memo="8月分 電話代")],
        [_row("1", "2026/06/03", "通信費", "携帯", "乙通信", "8800", memo="6月分"),
         _row("2", "2026/07/03", "通信費", "", "甲通信", "8800", memo="7月分")])
    assert _grade(plan, 2) == field_record.CONFIRMED
    assert _value(plan, 2) == "通信費"
    keys = {k for k, _n, _r in plan.citations[2]}
    assert keys == {"借方取引先", "借方補助科目"}, keys


def test_one_key_stays_single_however_many_precedents_it_has():
    """★★ 単: 当たった鍵が 1 本なら、その鍵に先例が 5 件あっても 単（§6.1 の致命 1）。

    ★ 変異「同じ鍵の N 件を出所 N つに数える」「単を確に格上げする」は、ここで赤くなる。
    """
    past = [_row(str(i), f"2026/0{i}/01", "通信費", "", "甲通信", "8800", memo=f"{i}月分")
            for i in range(1, 6)]
    plan = _plan([_row("11", "2026/08/03", "", "", "甲通信", "8800", memo="8月分")], past)
    assert _grade(plan, 2) == field_record.SINGLE
    assert _value(plan, 2) == "通信費"
    reason = accounts_core.reason_of(plan.records[2])
    assert "過去 5 件すべて 通信費" in reason, reason
    assert "2026/05/01" in reason, f"最新の先例（並びの最後）が原文で出ていない: {reason}"


def test_one_key_with_two_accounts_is_split_with_counts_and_dates():
    """割（1 本の鍵の内訳が 2 科目以上・`conflict=True`）── 科目ごとの件数と最新の日付。"""
    plan = _plan(
        [_row("11", "2026/08/20", "", "", "丙タクシー", "2500")],
        [_row("1", "2026/06/20", "旅費交通費", "", "丙タクシー", "2300"),
         _row("2", "2026/07/20", "会議費", "", "丙タクシー", "4000"),
         _row("3", "2026/07/25", "会議費", "", "丙タクシー", "1000")])
    assert _grade(plan, 2) == field_record.SPLIT
    assert _value(plan, 2) is None, "★ 割 は値を出さない（凍結した判断）"
    reason = accounts_core.reason_of(plan.records[2])
    for token in ("旅費交通費 1 件", "会議費 2 件", "2026/06/20", "2026/07/25"):
        assert token in reason, f"{token} が根拠に無い: {reason}"


def test_keys_that_point_at_different_accounts_are_split_with_both_sides():
    """割（鍵どうしが違う科目を指す）── `both_sides` が両側を出せる形。"""
    plan = _plan(
        [_row("11", "2026/08/20", "", "", "甲通信", "8800", memo="打合せ")],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800", memo="6月分"),
         _row("2", "2026/07/10", "会議費", "", "乙商店", "3000", memo="打合せ")])
    record = plan.records[2]
    assert field_record.grade(record) == field_record.SPLIT
    assert field_record.value(record) is None
    sides = {at: value for at, value, _how in field_record.both_sides(record)}
    assert sides == {"借方取引先": "通信費", "摘要": "会議費"}, sides


def test_no_precedent_is_none_found_and_always_carries_a_reason():
    """無: 先例が無い行は値を出さず、理由を必ず持つ（`Record` が型で守っている側）。"""
    plan = _plan(
        [_row("11", "2026/08/03", "", "", "甲通信", "8800", memo="8月分"),
         _row("12", "2026/08/25", "", "", "未知商会", "900", memo="よく分からない")],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800", memo="6月分")])
    assert _grade(plan, 3) == field_record.NONE_FOUND
    assert _value(plan, 3) is None
    assert plan.records[3].blank_reason, "★ 空欄なのに理由が無い"
    assert "未知商会" in plan.records[3].blank_reason


def test_the_grade_words_come_only_from_field_record():
    """★ 区分の語は `field_record` の定数と一致する（出口で作り直していない）。"""
    counts = accounts_core.grade_tally(_plan(
        [_row("11", "2026/08/03", "", "", "甲通信", "8800")],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800")]))
    assert sorted(counts) == sorted(field_record.GRADE_ORDER)


# --- 触らない行 ------------------------------------------------------------------

def test_a_continuation_row_is_never_touched():
    """★★ 複合仕訳の継続行（借方が丸ごと空）に候補を出さない（設計 §6.2）。

    ★ 変異「継続行にも候補を出す」はここで赤くなる。
    """
    plan = _plan(
        [_row("11", "2026/08/03", "", "", "甲通信", "8800", memo="8月分"),
         ["", "", "", "", "", "", "未払金", "", "継続行"]],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800", memo="6月分")])
    assert plan.candidates == [2], plan.candidates
    numbers = [r for r, _why in plan.untouched]
    assert numbers == [3], plan.untouched
    assert "継続行" in dict(plan.untouched)[3]


def test_a_row_that_is_already_filled_is_counted_not_touched():
    plan = _plan(
        [_row("11", "2026/08/03", "", "", "甲通信", "8800"),
         _row("12", "2026/08/04", "旅費交通費", "", "丙タクシー", "1000")],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800")])
    assert plan.candidates == [2]
    assert "旅費交通費" in dict(plan.untouched)[3]


def test_a_row_without_an_amount_is_not_a_candidate():
    """借方勘定科目が空でも、金額が無い行には候補を出さない（設計 §6.2）。"""
    plan = _plan(
        [_row("11", "2026/08/03", "", "", "甲通信", "8800"),
         _row("12", "2026/08/04", "", "", "甲通信", "", memo="金額なし")],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800")])
    assert plan.candidates == [2]
    assert accounts_core.DEBIT_AMOUNT in dict(plan.untouched)[3]


def test_a_full_width_space_counts_as_empty():
    """空の判定は `form_read.norm` を通す（全角空白 1 文字は空・設計 §6.2）。"""
    plan = _plan(
        [_row("11", "2026/08/03", "　", "", "甲通信", "8800")],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800")])
    assert plan.candidates == [2], "★ 全角空白を『埋まっている』と読んだ"


# --- 自己先例 --------------------------------------------------------------------

def test_todays_filled_rows_are_not_precedents():
    """★★ 今回の冊の埋まった行は先例に数えない（自己先例は『裏が取れた』の最短路・§6.3）。

    ★ 変異「今回の埋まった行も先例に数える」はここで赤くなる。
    """
    plan = _plan(
        [_row("11", "2026/08/01", "通信費", "携帯", "甲通信", "8800", memo="8月分"),
         _row("12", "2026/08/02", "通信費", "携帯", "甲通信", "8800", memo="8月分その2"),
         _row("13", "2026/08/03", "", "携帯", "甲通信", "8800", memo="8月分その3"),
         _row("14", "2026/08/04", "", "", "乙商店", "1200", memo="鉛筆")],
        [_row("1", "2026/06/10", "消耗品費", "", "乙商店", "1200", memo="コピー用紙")])
    assert _grade(plan, 4) == field_record.NONE_FOUND, \
        "★ 今回の冊の埋まった行を先例に数えている"
    assert _grade(plan, 5) == field_record.SINGLE


def test_two_keys_landing_on_one_past_row_is_single_not_confirmed():
    """★★ 2 本の鍵が**同じ 1 行**を指した時は 確 ではなく 単 ── 裏は 1 つ。

    ★ 実装者が踏んで名指しした穴（2026-09-13）: §6.1 の表は「2 本以上の鍵が当たれば 確」
      としか書いておらず、同じ過去 1 行を 借方補助科目 と 借方取引先 の 2 通りに読んだだけで
      確が立った。`field_record` は「写し合いの 2 つを裏と数えるな」と書いている
      （結合セルを 2 度数えて 40 冊すべて確になった実測と同じ形）。契約側を直して 単 に落とした。
    ★ 何を見て 1 つに数えたかは根拠に残す（黙って弱めない）。
    """
    plan = _plan(
        [_row("11", "2026/08/10", "", "文具", "乙商店", "1500", memo="ボールペン")],
        [_row("1", "2026/06/10", "消耗品費", "文具", "乙商店", "1200", memo="コピー用紙")])
    assert _grade(plan, 2) == field_record.SINGLE, _grade(plan, 2)
    reason = accounts_core.reason_of(plan.records[2])
    assert accounts_core.SAME_ROW_CAVEAT in reason, reason
    assert "借方補助科目" in reason and "借方取引先" in reason, reason


def test_two_keys_backed_by_different_past_rows_are_confirmed():
    """★ 陽性対照（上の対）── 別々の過去の行が同じ科目を指せば 確。"""
    plan = _plan(
        [_row("11", "2026/08/10", "", "文具", "乙商店", "1500", memo="ボールペン")],
        [_row("1", "2026/06/10", "消耗品費", "文具", "", "1200", memo="コピー用紙"),
         _row("2", "2026/07/10", "消耗品費", "", "乙商店", "900", memo="付箋")])
    assert _grade(plan, 2) == field_record.CONFIRMED, _grade(plan, 2)


# --- 断る ------------------------------------------------------------------------

def test_no_key_column_at_all_is_refused():
    """★ 鍵に使える列が 1 本も無ければ断る（静かに全行『無』にしない）。"""
    headers = ["取引日", "借方勘定科目", "借方金額(円)", "貸方勘定科目"]
    header_map, today = _book([["2026/08/03", "", "8800", "現金"]], headers)
    plan = accounts_core.plan_accounts(today, header_map, {})
    assert plan.refused and "鍵" in plan.refused, plan.refused
    assert not plan.records


def test_a_book_where_no_key_ever_hits_is_refused():
    """★★ 全行『無』の静かな成功をしない（設計 §6.3）── 文字コード違い・列の取り違え。"""
    plan = _plan([_row("11", "2026/08/03", "", "", "甲通信", "8800", memo="8月分")],
                 [_row("1", "2026/06/03", "通信費", "", "別会社", "8800", memo="6月分")])
    assert plan.refused, "★ 1 件も当たらないのに黙って成功した"
    assert not plan.records
    for token in ("借方取引先", "文字コード"):
        assert token in plan.refused, plan.refused


def test_two_columns_hitting_one_role_are_refused_by_name():
    """★ 2 列に当たっても断る（split の D1 と同じ線）── 両方を名指しする。"""
    headers = MF + ["摘要"]
    grid = _grid([_row("11", "2026/08/03", "", "", "甲通信", "8800", memo="x") + ["y"]],
                 headers)
    _head, _headers, header_map, refusal = accounts_core.resolve_accounts_columns(grid)
    assert refusal and "摘要" in refusal and "2 つ" in refusal, refusal
    assert header_map == {}, "★ 断ったのに列を決めている"


def test_a_missing_required_column_lists_what_it_saw():
    headers = ["取引日", "借方取引先", "摘要"]
    grid = _grid([["2026/08/03", "甲通信", "8月分"]], headers)
    _head, _headers, header_map, refusal = accounts_core.resolve_accounts_columns(grid)
    assert refusal and accounts_core.DEBIT_ACCOUNT in refusal, refusal
    assert "借方取引先" in refusal, f"見た見出しを並べていない: {refusal}"
    assert header_map == {}


# --- 弥生（見出し行が無い・列位置） ------------------------------------------------

def _yayoi_row(date="", account="", sub="", amount="", memo=""):
    row = [""] * accounts_core.YAYOI_COLUMN_COUNT
    row[0] = "2000"                   # 識別フラグ（★ 一次資料どおり 4 桁 ── 空の冊は検体の欠けだった）
    row[1] = "1"                      # 伝票No.
    row[3] = date
    row[4] = account
    row[5] = sub
    row[8] = amount
    row[16] = memo
    row[10] = "現金"
    return row


def test_the_yayoi_shape_is_accepted_by_position_only():
    """★ 見出し行が無い 25 列の冊は**列位置**で受ける（設計 §6.2 が名指しした唯一の形）。"""
    grid = [(i, _yayoi_row("R08/06/04", "", "ミュー商会", "2031", "宅配便"))
            for i in (1, 2)]
    head_row, _headers, header_map, refusal = accounts_core.resolve_accounts_columns(grid)
    assert refusal is None, refusal
    assert head_row is None, "★ 弥生には見出し行が無い"
    assert header_map == accounts_core.YAYOI_POSITIONS
    assert header_map[accounts_core.DEBIT_ACCOUNT] == 5
    assert header_map["摘要"] == 17


def test_a_headerless_book_that_is_not_the_yayoi_shape_is_refused():
    """★ 陰性対照 ── 列数が違う／1 列目に文字が在る冊は受けない（形だけで受ける線を守る）。"""
    short = [(1, ["あ", "い", "う"]), (2, ["え", "お", "か"])]
    assert accounts_core.resolve_accounts_columns(short)[3], "★ 3 列の冊を受けてしまった"
    wrong = _yayoi_row("R08/06/04", "", "ミュー商会", "2031", "宅配便")
    wrong[0] = "仕訳"
    assert accounts_core.resolve_accounts_columns([(1, wrong), (2, list(wrong))])[3], \
        "★ 1 列目に文字が在る 25 列の冊を弥生として受けてしまった"


def test_yayoi_pulls_from_the_memo_because_there_is_no_partner_column():
    """★ 弥生は取引先の列が無い ── 摘要と補助科目だけになる（正直に 単 が増える姿）。"""
    today = [(1, _yayoi_row("R08/07/01", "", "", "2031", "宅配便定期便"))]
    _head, _h, header_map, refusal = accounts_core.resolve_accounts_columns(today)
    assert refusal is None
    past_rows = [(1, _yayoi_row("R08/06/04", "荷造運賃", "", "2031", "宅配便定期便"))]
    plan = accounts_core.plan_accounts(
        today, header_map, {"弥生_過去.csv": {"header_map": header_map, "rows": past_rows}})
    assert plan.keys_used == ("借方補助科目", "摘要"), plan.keys_used
    assert _grade(plan, 1) == field_record.SINGLE
    assert _value(plan, 1) == "荷造運賃"
    assert "R08/06/04" in accounts_core.reason_of(plan.records[1]), \
        "★ 和暦略記を原文のまま運んでいない（日付は読まないと決めた側）"


# --- 貸方だけ取引先 --------------------------------------------------------------

def test_the_credit_partner_alone_can_pull_a_precedent():
    """★ 借方に取引先が無く、貸方側にだけ在る冊でも引ける（鍵は 4 本ある）。"""
    plan = _plan(
        [_row("11", "2026/08/03", "", "", "", "8800", credit_partner="甲カード")],
        [_row("1", "2026/06/03", "支払手数料", "", "", "8800", credit_partner="甲カード")])
    assert _grade(plan, 2) == field_record.SINGLE
    assert _value(plan, 2) == "支払手数料"
    assert plan.citations[2] == [("貸方取引先", "過去.csv", 2)], plan.citations[2]


def test_the_credit_account_is_never_a_key_and_it_is_written_down():
    """★ 払い方（貸方の勘定科目）は鍵に入れない ── 入れていないことを検分に書く（§6.3）。"""
    plan = _plan(
        [_row("11", "2026/08/03", "", "", "甲通信", "8800", credit="未払金")],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800", credit="未払金"),
         _row("2", "2026/06/04", "会議費", "", "乙商店", "3000", credit="未払金")])
    assert _grade(plan, 2) == field_record.SINGLE, "★ 払い方で引いてしまった"
    assert any("貸方は鍵に入れていません" in note for note in plan.notes), plan.notes
    kinds = {row[0] for row in accounts_core.inspection_rows(plan)}
    assert accounts_core.NOTE_KIND in kinds


# --- 表記ゆれ（名指しだけ・併合しない） ---------------------------------------------

def test_lookalike_spellings_are_named_but_never_merged():
    """★ `split_people.lookalike_pairs` をそのまま使う ── 寄せずに名指しする（§6.2）。"""
    plan = _plan(
        [_row("11", "2026/08/03", "", "", "株式会社アルファ", "8800", memo="8月分"),
         _row("12", "2026/08/04", "", "", "乙商店", "1200", memo="コピー用紙")],
        [_row("1", "2026/06/03", "通信費", "", "㈱アルファ", "8800", memo="6月分"),
         _row("2", "2026/06/10", "消耗品費", "", "乙商店", "1200", memo="コピー用紙")])
    pairs = {frozenset((one, two)) for key, one, two in plan.lookalike
             if key == "借方取引先"}
    assert frozenset(("株式会社アルファ", "㈱アルファ")) in pairs, plan.lookalike
    assert _grade(plan, 2) == field_record.NONE_FOUND, \
        "★ 表記ゆれを併合して引いた（名指しだけのはず）"


# --- 出口の材料 ------------------------------------------------------------------

def test_the_inspection_rows_split_the_denominator_from_the_untouched():
    """★ 検分は「空欄の理由」と「触らない行」を**別の種類**で出す（§6.5 の 2 の分母）。"""
    plan = _plan(
        [_row("11", "2026/08/03", "", "", "甲通信", "8800", memo="8月分"),
         _row("12", "2026/08/25", "", "", "未知商会", "900", memo="謎"),
         _row("13", "2026/08/26", "通信費", "", "甲通信", "8800", memo="埋まっている")],
        [_row("1", "2026/06/03", "通信費", "", "甲通信", "8800", memo="6月分")])
    rows = accounts_core.inspection_rows(plan)
    blanks = [r for r in rows if r[0] == accounts_core.BLANK_KIND]
    untouched = [r for r in rows if r[0] == accounts_core.UNTOUCHED_KIND]
    assert [r[1] for r in blanks] == [3], blanks
    assert [r[1] for r in untouched] == [4], untouched


def test_the_citations_are_machine_readable_round_trip():
    """★ 先例の番地は機械可読（検算がこれを読んでセルを見に行く・§6.5 の 1）。"""
    items = [("借方取引先", "過去 1.csv", 12), ("摘要", "過去_2.xlsx", 3)]
    text = accounts_core.format_citations(items)
    assert accounts_core.parse_citations(text) == items, text


def test_a_number_and_a_string_are_different_keys():
    """★ 型の扱いは `match.normalize_key` の線（数値 123 と文字 "123" は別の鍵）。"""
    assert accounts_core.key_identity(123) != accounts_core.key_identity("123")
    assert accounts_core.key_identity("甲　社") == accounts_core.key_identity("甲社")
    assert accounts_core.key_identity("　") is None


def test_a_headerless_25_column_file_with_an_empty_first_cell_is_refused():
    """★ 陰性対照 ── 弥生の形は「25 列・1 列目が 4 桁の数字」まで。1 列目が空なら受けない。

    ★ 実装の途中で検体の 1 列目が空だったため受け入れを広げかけた（2026-09-13）。
      それは「検体の欠けに合わせて規則を書く」形なので、読み手は一次資料どおり狭いままにし、
      検体の側を直した。この番人はその線を守る（広げると赤くなる）。
    """
    row = _yayoi_row("R08/07/01", "", "", "2031", "宅配便")
    row[0] = ""
    _head, _headers, header_map, refused = accounts_core.resolve_accounts_columns([(1, row)])
    assert refused and "4 桁" in refused, refused
    assert not header_map, "★ 断ったのに列を決めている（推測で先へ進んでいる）"


# --- 割れた鍵は「拒否権」でなく「沈黙」（2026-09-13・実物の形で測って直した）------------
#
# ★★ なぜ在るか: Namakoo 提供の請求書 3 通から仕訳を起こして測ったら、到達が
#   合成検体の 75% から **30%** に落ちた。落ちた分は「鍵が当たらない」ではなく
#   **広い鍵（取引先）の割れが、完全に一致している狭い鍵（摘要）を道連れにしていた**。
#   タクシー 3 行はすべて摘要が過去 1 行と完全一致して正しい科目を指していたのに空欄だった。
#   ★ 実務では支払先が割れているのが普通（通販・タクシー・雑貨）で、狭い鍵で解くのが作法。
# ★ この欠陥は**合成検体では永久に見つからない**（過去の行が疎で、割れた鍵と当たる鍵が
#   同じ行に同居しないため）── 実測: 規則を変えても合成 155 行の点数は 155/155 のまま動かない。

def test_a_split_key_does_not_veto_a_key_that_matches():
    """★★ 取引先が割れていても、摘要が過去 1 行と一致していれば引ける（値が出る）。"""
    plan = _plan(
        [_row("11", "2026/04/10", "", "", "丙タクシー", "4500", memo="タクシー乗車 出張")],
        [_row("1", "2026/02/10", "旅費交通費", "", "丙タクシー", "4500", memo="タクシー乗車 出張"),
         _row("2", "2026/02/15", "接待交際費", "", "丙タクシー", "8200", memo="顧客送迎")])
    assert _grade(plan, 2) == field_record.SINGLE, _grade(plan, 2)
    assert _value(plan, 2) == "旅費交通費"


def test_the_split_is_still_named_even_when_a_value_comes_out():
    """★★ 沈黙させるのは**出所の数え方**だけ ── 「この支払先は割れている」は人に伝える。

    ★ 初版はここを落とした（`conflict` が真のときだけ理由に出していたので、値が出た行から
      内訳が消えた）。割れは出所に数えないだけで、伝えるべき事実は変わらない。
    """
    plan = _plan(
        [_row("11", "2026/04/10", "", "", "丙タクシー", "4500", memo="タクシー乗車 出張")],
        [_row("1", "2026/02/10", "旅費交通費", "", "丙タクシー", "4500", memo="タクシー乗車 出張"),
         _row("2", "2026/02/15", "接待交際費", "", "丙タクシー", "8200", memo="顧客送迎")])
    reason = accounts_core.reason_of(plan.records[2])
    for token in ("丙タクシー", "旅費交通費 1 件", "接待交際費 1 件", "出所に数えていません"):
        assert token in reason, f"{token} が根拠に無い: {reason}"


def test_only_a_split_key_is_split_not_none_found():
    """★★ 割れた鍵**しか**無い行は 割 ── 無 にすると「先例がありません」という嘘になる。

    ★ 凍結した検体の答えがここを守った（沈黙させすぎた初版で 16 行が 割 → 無 に落ちて赤くなった）。
    """
    plan = _plan(
        [_row("11", "2026/04/10", "", "", "丙タクシー", "4500", memo="はじめての摘要")],
        [_row("1", "2026/02/10", "旅費交通費", "", "丙タクシー", "4500", memo="別の摘要"),
         _row("2", "2026/02/15", "接待交際費", "", "丙タクシー", "8200", memo="また別の摘要")])
    assert _grade(plan, 2) == field_record.SPLIT, _grade(plan, 2)
    assert _value(plan, 2) is None
    reason = plan.records[2].blank_reason
    assert "割れて" in reason, reason
    # ★ 物差しを 1 度直した: 「先例がありません」で探すと『ほかの鍵に先例がありません』
    #   （正しい文）に当たって赤くなった。嘘なのは**どの鍵にも**と言い切る方。
    assert "どの鍵にも先例がありません" not in reason, f"★ 嘘（先例は在って割れている）: {reason}"
    assert "借方取引先" in reason, f"どの鍵が割れているか名指ししていない: {reason}"


def test_a_split_key_caps_the_grade_at_single():
    """★ 割れた鍵が在る行は 確 と名乗らない ── 掃けていない口が残っている。"""
    plan = _plan(
        [_row("11", "2026/04/10", "", "文具", "丙タクシー", "4500", memo="タクシー乗車 出張")],
        [_row("1", "2026/02/10", "旅費交通費", "", "丙タクシー", "4500", memo="タクシー乗車 出張"),
         _row("2", "2026/02/15", "接待交際費", "", "丙タクシー", "8200", memo="顧客送迎"),
         _row("3", "2026/03/01", "旅費交通費", "文具", "", "300", memo="別の行")])
    assert _grade(plan, 2) == field_record.SINGLE, _grade(plan, 2)
    assert accounts_core.SPLIT_KEY_CAP in accounts_core.reason_of(plan.records[2])


def test_two_keys_disagreeing_is_still_split():
    """★ 陰性対照 ── **本物の食い違い**（鍵どうしが違う科目）は今までどおり 割。"""
    plan = _plan(
        [_row("11", "2026/04/10", "", "", "甲通信", "8800", memo="打合せ")],
        [_row("1", "2026/02/03", "通信費", "", "甲通信", "8800", memo="2月分"),
         _row("2", "2026/03/10", "会議費", "", "乙商店", "3000", memo="打合せ")])
    assert _grade(plan, 2) == field_record.SPLIT
    assert _value(plan, 2) is None


def test_the_last_precedent_is_not_called_the_newest():
    """★ 2026-09-13（買い手役の初見）: `R08/08/25` が在るのに「最新 R08/07/25」と出た ── 日付を
    読まないと決めたのに「最新」と呼ぶのは、同じ文で開示していても画面の主語が嘘。"""
    past = [_row("1", "2026/08/25", "地代家賃", "駐車場", "丙パーキング", "8000", memo="8月分"),
            _row("2", "2026/07/25", "地代家賃", "駐車場", "丙パーキング", "8000", memo="7月分")]
    plan = _plan([_row("11", "2026/09/25", "", "駐車場", "丙パーキング", "8000", memo="9月分")], past)
    reason = accounts_core.reason_of(plan.records[2])
    assert "最新" not in reason, reason
    assert "並びで最後の先例 2026/07/25" in reason, reason


# --- 付けた科目が先例と違う行（2026-09-14・会計役が 2 回「一番期待した」所）------------------
#
# ★★ 借方勘定科目が既に埋まった行は「触らない行」として番号だけ控えて終わっていた ── 先例と
#   突き合わせていなかった。候補を出す行では候補＝先例なので「違う」は原理的に出ない。
# ★ 鳴らす条件は狭い: その鍵の先例が**全部同じ科目**で、付けた科目と違うときだけ。
#   割れた鍵・先例 0 件・同じ科目は**沈黙**（3 つとも陰性対照）。値は 1 文字も変えない。

_D_PAST = [_row("1", "2026/08/25", "通信費", "本社回線", "ＮＴＴ西日本", "8800", memo="8月分 回線"),
           _row("4", "2026/08/26", "消耗品費", "", "アスクル", "3000", memo="コピー用紙"),
           _row("5", "2026/08/27", "新聞図書費", "", "アスクル", "2000", memo="業界誌")]


def _differs(today_rows):
    return {r: w for r, w in _plan(today_rows, _D_PAST).differs}


def test_a_filled_account_that_contradicts_a_unanimous_precedent_is_named():
    got = _differs([_row("11", "2026/09/01", "仮払金", "本社回線", "ＮＴＴ西日本", "8800",
                         memo="9月分 回線")])
    assert list(got) == [2], got
    why = got[2]
    assert "付けた科目『仮払金』と先例が違います" in why, why
    assert "借方取引先『ＮＴＴ西日本』の先例 1 件はすべて『通信費』" in why, why
    assert "2 行目" in why, why
    assert "付け替えたのが正しいなら、このままで構いません" in why, why


def test_a_split_key_stays_silent():
    """★ 陰性対照 ── **割れた鍵だけが当たる行**では黙る（「違う」と言えない）。

    ★ 対照の作り方を 2 度間違えた（記録として残す）:
      ① 割れた側の 1 つ目と同じ科目を付けた → 沈黙の規則を外した版でも黙り、対照にならない
      ② どちらとも違う科目にしたら、**摘要**の鍵が一致して鳴った ── これは**正しい挙動**
         （鍵ごとに独立に見るので、狭い鍵が一致すればそちらで鳴る）。
      → だから対照は「割れた鍵しか当たらない行」にする（摘要も補助科目も先例に無い）。
    """
    got = _differs([_row("12", "2026/09/01", "雑費", "", "アスクル", "3000",
                         memo="初めて買うもの")])
    assert got == {}, got


def test_a_narrow_key_still_fires_even_when_a_wide_key_is_split():
    """★★ 鍵ごとに独立 ── 支払先が割れていても、摘要が全部同じなら**そちらで**鳴る。
    実務でいちばん多い形（通販・タクシー・雑貨は支払先が割れ、摘要で決まる）。"""
    got = _differs([_row("12", "2026/09/01", "雑費", "", "アスクル", "3000",
                         memo="コピー用紙")])
    assert list(got) == [2], got
    assert "摘要『コピー用紙』の先例 1 件はすべて『消耗品費』" in got[2], got[2]
    assert "借方取引先" not in got[2], f"割れた鍵は言わない: {got[2]}"


def test_a_new_partner_stays_silent():
    assert _differs([_row("13", "2026/09/01", "旅費交通費", "", "新しい会社", "1000",
                          memo="タクシー")]) == {}


def test_the_same_account_stays_silent():
    assert _differs([_row("14", "2026/09/01", "通信費", "本社回線", "ＮＴＴ西日本", "8800",
                          memo="9月分 回線")]) == {}


def test_a_candidate_row_is_not_in_this_list():
    """★ 空の行は候補の側（候補＝先例なので「違う」は出ない）── 分母を混ぜない。"""
    plan = _plan([_row("15", "2026/09/01", "", "本社回線", "ＮＴＴ西日本", "8800", memo="9月分")],
                 _D_PAST)
    assert plan.differs == [] and 2 in plan.records, (plan.differs, sorted(plan.records))


def test_the_report_carries_the_kind_and_the_blank_denominator_is_unchanged():
    plan = _plan([_row("11", "2026/09/01", "仮払金", "本社回線", "ＮＴＴ西日本", "8800",
                       memo="9月分 回線")], _D_PAST)
    rows = accounts_core.inspection_rows(plan)
    kinds = [r[0] for r in rows]
    assert accounts_core.DIFFERS_KIND in kinds, kinds
    assert accounts_core.BLANK_KIND not in kinds, kinds      # ★ 空欄の分母には入れない
