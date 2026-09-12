# -*- coding: utf-8 -*-
"""器官 `ailine_core/split_people.py` の番人 ── 担当者別に分ける／疑う（需要⑤・2026-09-12）。

設計: docs/DESIGN-20260912-担当者別に分けて配る.md（D1〜D6）。

契約（★ この器官は「代わりに決めない」）:
  - 束ねる鍵は**原本の文字そのまま**。`山田` と `山田　` は別の冊になり、組を**名指し**する
  - 合計・小計の語がある行は配らない ── **担当者の列そのものにラベルが入る表でも**
  - 担当者の列が空の行はどの冊にも入れず、行番号で名指しする
  - 複数担当（`山田/佐藤`）はどの冊にも入れない
  - 当たる見出しが 0 個／2 個以上なら分けない（理由に両方の見出しを出す）
  - 金額が文字（`10,000円`）の行は数えず名指しする

★ ここは純関数だけ（ファイルも openpyxl も LibreOffice も要らない）。
★ 変異で赤くなること（実測で確かめた変異）:
    ① 合計語の行を配る            → test_a_total_row_is_never_handed_to_anyone
    ② ゆれを併合して 1 冊にする     → test_lookalike_spellings_stay_separate_and_get_named
    ③ 空欄の行を誰かに配る          → test_a_blank_assignee_goes_to_nobody
    ④ 見出しが 2 列でも分ける       → test_two_matching_headers_refuse_to_split
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core import split_people   # noqa: E402

HEADERS = ["日付", "案件", "顧客", "担当者", "金額"]


def _rows(*data, header_row: int = 1, headers=None):
    """(行番号, 値の並び) の列を作る。見出し行は header_row に置く。"""
    grid = [(header_row, list(headers if headers is not None else HEADERS))]
    for i, values in enumerate(data, start=header_row + 1):
        grid.append((i, list(values)))
    return grid


def _plan(*data, header_row: int = 1, headers=None, by="担当者", amount="金額"):
    return split_people.plan_split(_rows(*data, header_row=header_row, headers=headers),
                                    header_row, by, amount)


# ── 陽性: 分ける ──────────────────────────────────────────────

def test_rows_are_grouped_by_the_exact_text_in_the_column():
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000],
                 ["2026-06-02", "印刷", "乙社", "佐藤", 2000],
                 ["2026-06-03", "保守", "丙社", "山田", 3000])
    assert plan.refused is None
    assert plan.parts == {"山田": [2, 4], "佐藤": [3]}
    assert plan.part_amounts == {"山田": 4000.0, "佐藤": 2000.0}
    assert plan.whole_rows == 3 and plan.whole_amount == 6000.0


def test_the_header_row_does_not_have_to_be_the_first_row():
    """★ 1 行目決め打ちにしない（設計 D1・09-06 に 2 回踏んだ形）。"""
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000], header_row=3)
    assert plan.header_row == 3 and plan.parts == {"山田": [4]}


# ── 疑う: 表記ゆれ（併合はしない）────────────────────────────────

def test_lookalike_spellings_stay_separate_and_get_named():
    """★★ 変異②「ゆれを併合して 1 冊にする」がここで赤くなる。

    併合は「同じ人だと決める」こと ── この製品はしない。別の冊のまま**名指し**する。
    """
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000],
                 ["2026-06-02", "印刷", "乙社", "山田　", 2000])
    assert sorted(plan.parts) == sorted(["山田", "山田　"]), "併合してはいけない"
    assert [sorted(p) for p in plan.lookalike] == [sorted(["山田", "山田　"])], \
        "norm が同じになる組を名指ししていない"


def test_kana_and_kanji_are_not_merged_and_not_even_named():
    """★ カナ↔漢字は「同じ人だと決める」ことなので寄せない ── 名指しもしない（D4 の★）。
       検体で `ヤマダ` が出たら、それは**取り逃し**として見えたまま残す（決められない
       ものを指ささない）。★ ここが変わるなら設計 D4 を先に書き換えること。"""
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000],
                 ["2026-06-02", "印刷", "乙社", "ヤマダ", 2000])
    assert sorted(plan.parts) == sorted(["山田", "ヤマダ"])
    assert plan.lookalike == []


# ── 疑う: 空欄・複数担当 ────────────────────────────────────────

def test_a_blank_assignee_goes_to_nobody():
    """★★ 変異③「空欄の行を誰かに配る」がここで赤くなる（空欄は誤配より安い）。"""
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000],
                 ["2026-06-02", "印刷", "乙社", None, 2000])
    assert plan.blank == [3]
    assert all(3 not in rows for rows in plan.parts.values()), "空欄の行を配ってはいけない"
    assert plan.blank_amount == 2000.0, "空欄の金額は証明の分母に残す"


def test_two_people_in_one_cell_are_handed_back_not_guessed():
    plan = _plan(["2026-06-01", "広告", "甲社", "山田/佐藤", 1000],
                 ["2026-06-02", "印刷", "乙社", "鈴木・田中", 2000])
    assert plan.multi == [(2, "山田/佐藤"), (3, "鈴木・田中")]
    assert plan.parts == {}
    assert plan.multi_amount == 3000.0


# ── 分けない行（★ 最悪の混入を止める側）────────────────────────────

def test_a_total_row_is_never_handed_to_anyone():
    """★★ 変異①「合計語の行を配る」がここで赤くなる。合計行の混入が最悪の事故（D2）。"""
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000],
                 ["小計", None, None, None, 1000],
                 ["合計", None, None, None, 1000])
    assert plan.excluded == [3, 4]
    assert plan.parts == {"山田": [2]}
    assert plan.whole_amount == 1000.0, "分けない行の金額を全体に入れてはいけない（倍になる）"


def test_a_total_label_inside_the_assignee_column_is_still_not_handed_out():
    """★★ 2026-09-12 の実測（検体 S05・★混入 2 件）の凍結。

    担当者の列が**1 列目**の表では、合計行のラベルがその列そのものに入る。初版は
    「担当者の列が空で」を条件に足していたため、**『合計』という名前の人の冊**が出来た。
    ★ 語が在る行は、担当者の列が埋まっていても配らない。
    """
    headers = ["記入者", "日付", "顧客", "金額"]
    plan = split_people.plan_split(
        _rows(["山田", "2026-07-02", "甲社", 5000],
              ["合計", None, None, 5000],
              headers=headers),
        1, "記入者", "金額")
    assert plan.parts == {"山田": [2]}, f"合計行を配った: {plan.parts}"
    assert plan.excluded == [3]


def test_a_note_row_with_a_single_cell_is_not_a_detail_row():
    """★ 備考・小見出しの行（1 セルだけの行）は明細ではないので配らない（D2『備考』）。"""
    headers = ["記入者", "日付", "顧客", "金額"]
    plan = split_people.plan_split(
        _rows(["山田", "2026-07-02", "甲社", 5000],
              ["※7月分は月末に集計", None, None, None],
              headers=headers),
        1, "記入者", "金額")
    assert plan.parts == {"山田": [2]}, f"備考の行を配った: {plan.parts}"
    assert plan.excluded == [3]


def test_a_completely_empty_row_is_not_a_finding_about_a_person():
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000],
                 [None, None, None, None, None])
    assert plan.excluded == [3] and plan.blank == []


# ── D1: 列が決まらないなら分けない ─────────────────────────────────

def test_two_matching_headers_refuse_to_split():
    """★★ 変異④「見出しが 2 列でも分ける」がここで赤くなる（設計 D1）。"""
    headers = ["日付", "顧客", "担当", "営業担当", "金額"]
    plan = split_people.plan_split(
        _rows(["2026-06-02", "甲社", "内藤", "山田", 30000], headers=headers),
        1, "担当", "金額")
    assert plan.refused, "『担当』と『営業担当』が同居する表を分けてはいけない"
    assert "担当" in plan.refused and "営業担当" in plan.refused, plan.refused
    assert plan.parts == {} and plan.by_column is None


def test_a_missing_header_refuses_and_shows_what_it_saw():
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000], by="記入者")
    assert plan.refused and "記入者" in plan.refused
    for name in HEADERS:
        assert name in plan.refused, f"見た見出しを並べていない: {plan.refused}"


def test_an_ambiguous_amount_header_also_refuses():
    """★ 金額の列が決まらないなら分けない ── 証明の分母を推測で作らない（D5）。"""
    headers = ["日付", "担当者", "金額", "金額(税込)"]
    plan = split_people.plan_split(
        _rows(["2026-06-01", "山田", 1000, 1100], headers=headers),
        1, "担当者", "金額")
    assert plan.refused and "金額(税込)" in plan.refused
    assert plan.parts == {}


def test_matching_columns_is_the_line_that_makes_two_hits_possible():
    assert split_people.matching_columns(["日付", "担当", "営業担当"], "担当") == [2, 3]
    assert split_people.matching_columns(["担当者", "金額"], "担当者") == [1]
    assert split_people.matching_columns(["担 当 者", "金額"], "担当者") == [1], \
        "空白は norm が落とす（『担 当 者』も同じ見出し）"
    assert split_people.matching_columns(["日付", "金額"], "担当") == []


# ── D5: 金額と証明 ─────────────────────────────────────────────

def test_a_text_amount_is_named_not_reinterpreted():
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", "10,000円"],
                 ["2026-06-02", "印刷", "乙社", "佐藤", 2000])
    assert plan.unparsed == [2], "文字の金額を名指ししていない"
    assert plan.part_amounts == {"佐藤": 2000.0}, "文字を数に読み替えてはいけない"
    assert plan.whole_amount == 2000.0


def test_without_an_amount_header_nothing_about_money_is_claimed():
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000], amount=None)
    assert plan.amount_counted is False
    assert plan.whole_amount == 0.0 and plan.unparsed == []


def test_proof_breaks_finds_the_equation_that_does_not_close():
    """★ 陽性対照: 破れを見つけられること（見つけられない検算は番人でない）。"""
    ok = {"rows": {"whole": 5, "parts": 3, "blank": 1, "excluded": 1, "multi": 0},
          "amount": {"counted": True, "whole": 100.0, "parts": 90.0, "blank": 10.0, "multi": 0}}
    assert split_people.proof_breaks(ok) == []
    bad_rows = {"rows": {"whole": 5, "parts": 2, "blank": 1, "excluded": 1, "multi": 0}}
    assert len(split_people.proof_breaks(bad_rows)) == 1
    bad_amount = {"rows": ok["rows"],
                  "amount": {"counted": True, "whole": 100.0, "parts": 80.0,
                             "blank": 10.0, "multi": 0}}
    assert len(split_people.proof_breaks(bad_amount)) == 1
    not_counted = {"rows": ok["rows"],
                   "amount": {"counted": False, "whole": 0, "parts": 999}}
    assert split_people.proof_breaks(not_counted) == [], \
        "数えていない金額で破れを名乗らない"


def test_the_parts_of_a_plan_add_up_to_the_whole():
    """★ 分母（D5 の行の等式）が、この器官の出す分類だけで閉じていること。"""
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000],
                 ["2026-06-02", "印刷", "乙社", None, 2000],
                 ["2026-06-03", "保守", "丙社", "山田/佐藤", 3000],
                 ["合計", None, None, None, 6000],
                 [None, None, None, None, None])
    parts = sum(len(rows) for rows in plan.parts.values())
    assert parts + len(plan.blank) + len(plan.multi) + len(plan.excluded) == plan.whole_rows


# ── ファイル名の安全化（シート内の値は触らない）─────────────────────

@pytest.mark.parametrize("value, want", [
    ("山田", "山田"),
    ("山田　", "山田_"),
    ("甲/乙", "甲_乙"),
    ("a:b*c?d", "a_b_c_d"),
])
def test_a_value_becomes_a_safe_file_name(value, want):
    assert split_people.safe_filenames([value])[value] == want


def test_two_values_never_share_one_file():
    got = split_people.safe_filenames(["甲/乙", "甲:乙"])
    assert len(set(got.values())) == 2, f"別の表記が同じ冊へ落ちる: {got}"


def test_the_inspection_book_name_is_not_stolen_by_a_person():
    got = split_people.safe_filenames([split_people.REPORT_STEM],
                                       reserved=(split_people.REPORT_STEM,))
    assert got[split_people.REPORT_STEM] != split_people.REPORT_STEM


def test_the_value_itself_is_never_rewritten():
    """★ ファイル名は安全化するが、束ねる鍵（＝シートに書く値）は原本のまま。"""
    plan = _plan(["2026-06-01", "広告", "甲社", "山田　", 1000])
    assert list(plan.parts) == ["山田　"], "値を書き換えてはいけない"


# ── 検分の行（分母つき・名指し）────────────────────────────────────

def test_the_inspection_rows_name_every_kind_of_finding():
    plan = _plan(["2026-06-01", "広告", "甲社", "山田", 1000],
                 ["2026-06-02", "印刷", "乙社", "山田　", 2000],
                 ["2026-06-03", "保守", "丙社", None, 3000],
                 ["2026-06-04", "保守", "丁社", "甲/乙", 4000],
                 ["2026-06-05", "保守", "戊社", "田中", "5,000円"],
                 ["合計", None, None, None, 10000])
    rows = split_people.report_rows(plan, {}, proof={"rows": {"whole": plan.whole_rows,
                                                              "parts": 3, "blank": 1,
                                                              "excluded": 1, "multi": 1},
                                                     "amount": {"counted": False}, "ok": True})
    kinds = {r[0] for r in rows}
    for kind in ("配った", "空欄", "複数担当", "分けない行", "表記ゆれ", "金額が文字",
                 "証明（行）"):
        assert kind in kinds, f"検分に『{kind}』が無い: {sorted(kinds)}"
    assert len(rows[0]) == len(split_people.REPORT_HEADERS), "見出しと列数が合わない"


def test_a_row_that_merely_contains_a_total_word_is_still_handed_out():
    """★★ 2026-09-12 の探針: 実装の初版は「合計の語が行のどこかに在れば配らない」に広げていて、
    `設計費／合計商事／山田／1000` が山田の冊から**消えた**。受け取った人は消えた行に気づけない。
    合計行と見るのは 2 つの形だけ ── 担当者が空で語が在る／担当者の列そのものが語。"""
    from ailine_core.split_people import plan_split
    hdr = ["日付", "案件", "顧客", "担当者", "金額"]
    rows = [hdr,
            ["2026-08-01", "設計費", "合計商事", "山田", 1000],
            ["2026-08-02", "7月合計分", "A社", "佐藤", 2000],
            [None, "小計", None, None, 3000],
            [None, "合計", None, None, 3000]]
    p = plan_split([(i + 1, r) for i, r in enumerate(rows)], 1, "担当者", "金額")
    assert p.parts == {"山田": [2], "佐藤": [3]}, p.parts
    assert p.excluded == [4, 5], p.excluded
    # ★ 担当者の列が 1 列目で、合計のラベルがその列に入る表（検体 S05 の形）は配らない
    hdr2 = ["記入者", "日付", "件名", "金額"]
    rows2 = [hdr2, ["山田", "2026-07-01", "対応", 100], ["合計", None, None, 100]]
    p2 = plan_split([(i + 1, r) for i, r in enumerate(rows2)], 1, "記入者", "金額")
    assert p2.parts == {"山田": [2]} and p2.excluded == [3], (p2.parts, p2.excluded)
