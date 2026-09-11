# 束で疑う（forms_suspect）の番人 ── 種類ごとの陽性・陰性対照・変異（2026-09-11）
#
# ★ 検体（bench/received_invoices/received_bundle・規則を知らない側が書いた 5 束 51 冊）での実測:
#   正 5 / 取り逃し 3 / 偽の疑い 1（令和の日付が読めず「空欄」に見えた）。
#   ここは器官を通さず、器官が出す「値の表」に対して束の論理だけを縛る。
from __future__ import annotations

import copy
import datetime as dt

from ailine_core import forms_suspect as fs


def _book(vendor, date, number, amount):
    return {"請求元": vendor, "宛先": "ナギ商会株式会社", "請求額": amount,
            "請求日": dt.date(*date) if date else None, "請求番号": number}


def _normal() -> dict:
    """毎月ふつうに来ているだけの束（陰性対照）。金額の変動は 2 倍強まで。"""
    t = {}
    for v, base in (("桜庭紙業株式会社", 101), ("蓮見工業株式会社", 201), ("白川運送有限会社", 301)):
        for m, amt in ((6, 36300), (7, 23100), (8, 39600)):
            t[f"{v}_2026-0{m}.xlsx"] = _book(v, (2026, m, 30), f"INV-2026-0{m}-{base}", amt)
    return t


def _kinds(found):
    return {(s["種類"], tuple(sorted(s["冊"]))) for s in found}


def test_a_normal_bundle_raises_no_suspicion():
    """★ 陰性対照 ── ここで 1 件でも出たらオオカミ少年。"""
    assert fs.suspect(_normal()) == []


def test_the_same_invoice_twice_is_a_duplicate():
    t = _normal()
    t["桜庭紙業株式会社_2026-08 (2).xlsx"] = copy.deepcopy(t["桜庭紙業株式会社_2026-08.xlsx"])
    found = fs.suspect(t)
    assert len(found) == 1, found
    s = found[0]
    assert s["種類"] == fs.DUPLICATE
    assert sorted(s["冊"]) == ["桜庭紙業株式会社_2026-08 (2).xlsx", "桜庭紙業株式会社_2026-08.xlsx"]
    assert "重複" in s["理由"] and "同じ" in s["理由"], s["理由"]


def test_a_duplicate_needs_a_date_or_a_number_not_just_the_amount():
    """★ 読めなかったもの（None）は根拠にしない ── 金額だけ同じでは重複と言わない。"""
    t = {"a.xlsx": _book("桜庭紙業株式会社", None, None, 36300),
         "b.xlsx": _book("桜庭紙業株式会社", None, None, 36300)}
    assert fs.suspect(t) == []


def test_same_month_different_amount_is_a_reissue():
    t = _normal()
    t["桜庭紙業株式会社_2026-07_再発行.xlsx"] = _book("桜庭紙業株式会社", (2026, 7, 31), "INV-2026-07-101-2", 46200)
    found = fs.suspect(t)
    assert _kinds(found) == {(fs.REISSUE, ("桜庭紙業株式会社_2026-07.xlsx", "桜庭紙業株式会社_2026-07_再発行.xlsx"))}
    why = found[0]["理由"]
    assert "同じ月" in why and "金額" in why and "枝番" in why, why


def test_mutating_the_amount_by_one_yen_turns_a_duplicate_into_a_reissue():
    """★ 変異: 1 円ずらすと 重複 が消えて 訂正再発行 になる（同じ組で）。"""
    t = _normal()
    dup = "桜庭紙業株式会社_2026-08 (2).xlsx"
    t[dup] = copy.deepcopy(t["桜庭紙業株式会社_2026-08.xlsx"])
    pair = (dup, "桜庭紙業株式会社_2026-08.xlsx")
    assert (fs.DUPLICATE, pair) in _kinds(fs.suspect(t))
    t[dup]["請求額"] += 1
    k = _kinds(fs.suspect(t))
    assert (fs.DUPLICATE, pair) not in k and (fs.REISSUE, pair) in k, k


def test_one_book_a_year_away_is_a_year_error():
    t = _normal()
    t["桜庭紙業株式会社_2026-07.xlsx"]["請求日"] = dt.date(2025, 7, 31)
    found = fs.suspect(t)
    assert _kinds(found) == {(fs.YEAR_OFF, ("桜庭紙業株式会社_2026-07.xlsx",))}
    assert "年" in found[0]["理由"] and "2025" in found[0]["理由"], found[0]["理由"]


def test_crossing_a_year_boundary_is_not_a_year_error():
    """★ 11 月・12 月・1 月と普通に年をまたぐ並びは疑わない。"""
    t = {}
    for i, (y, m) in enumerate(((2025, 11), (2025, 12), (2026, 1))):
        t[f"{i}.xlsx"] = _book("桜庭紙業株式会社", (y, m, 28), f"INV-{i}", 33000)
    assert fs.suspect(t) == []


def test_ten_times_the_other_months_is_a_magnitude_error():
    t = _normal()
    t["桜庭紙業株式会社_2026-07.xlsx"]["請求額"] = 231000
    found = fs.suspect(t)
    assert _kinds(found) == {(fs.MAGNITUDE, ("桜庭紙業株式会社_2026-07.xlsx",))}
    why = found[0]["理由"]
    assert "桁" in why and "10" in why and "他の月" in why, why


def test_a_price_rise_of_sixty_percent_is_not_a_magnitude_error():
    t = _normal()
    t["桜庭紙業株式会社_2026-08.xlsx"]["請求額"] = 47520     # 29,700 の 1.6 倍
    assert fs.suspect(t) == []


def test_the_same_number_on_two_different_invoices_clashes():
    t = _normal()
    t["桜庭紙業株式会社_2026-07.xlsx"]["請求番号"] = "INV-2026-06-101"
    found = fs.suspect(t)
    assert _kinds(found) == {(fs.NUMBER_CLASH, ("桜庭紙業株式会社_2026-06.xlsx", "桜庭紙業株式会社_2026-07.xlsx"))}
    why = found[0]["理由"]
    assert "番号" in why and "重なる" in why and "重複" in why, why


def test_a_blank_date_where_the_siblings_have_one_is_flagged_but_an_all_blank_vendor_is_not():
    t = _normal()
    t["桜庭紙業株式会社_2026-08.xlsx"]["請求日"] = None
    for m in (6, 7, 8):
        t[f"蓮見工業株式会社_2026-0{m}.xlsx"]["請求日"] = None      # 骨ごと読めない取引先
    found = fs.suspect(t)
    assert _kinds(found) == {(fs.BLANK_DATE, ("桜庭紙業株式会社_2026-08.xlsx",))}, found
    assert "請求日" in found[0]["理由"] and "空" in found[0]["理由"]


def test_books_whose_vendor_could_not_be_read_stay_out_of_vendor_checks():
    """★ 取引先が読めない冊は取引先ごとの照合に入れない（穴の値段を隠さず、誤って組ませない）。"""
    t = {"a.xlsx": _book(None, (2026, 7, 31), "X-1", 29700),
         "b.xlsx": _book(None, (2026, 7, 31), "X-2", 46200)}
    assert fs.suspect(t) == []


def test_the_kinds_are_the_only_words_the_findings_use():
    t = _normal()
    t["桜庭紙業株式会社_2026-08 (2).xlsx"] = copy.deepcopy(t["桜庭紙業株式会社_2026-08.xlsx"])
    t["蓮見工業株式会社_2026-07.xlsx"]["請求額"] = 231000
    for s in fs.suspect(t):
        assert s["種類"] in fs.KINDS and s["冊"] and s["理由"].strip()
