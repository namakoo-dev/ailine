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


# --- 請求番号の書き方の違いで重複を取り逃さない（2026-09-13・実物の検体から）------------
#
# ★★ なぜ在るか: Namakoo 提供の実物の形式の請求書で、請求番号が
#   `ＩＮＶー０９ー９９９`（全角英数 ＋ **長音符**）だった。原文がそうで、器官は忠実に読んでいる
#   ── かな入力のまま `-` を打つ、実務で一番多い打ち間違い。
#   同じ請求書が PDF と Excel で 2 通来て片方が `INV-09-999` だと、**重複が 1 件も鳴らなかった**。
#   ★ 二重払いは取り逃しの中で一番高くつく。
# ★ 値は原本のまま（全角は「文字で入っていた」証拠）。畳むのは**照合の鍵だけ**で、
#   畳んで一致しても原本が違えば**必ず名指しする**（併合はしない ── split の lookalike と同じ線）。

_SAME = {"請求元": "ラムダ株式会社", "宛先": "㈱デルタ", "請求額": 330000, "請求日": "2026-09-13"}


# ★ 既存の `_kinds(found)`（結果のリストを受ける）と**同名にしない** ── 上書きすると
#   古い試験 6 本が道連れで落ちた（実測）。この段の器は books を受けるので名前を分ける。
def _kinds_of(books):
    """★ 既存の `_kinds(found)` は（種類, 冊）の組を返す。ここで欲しいのは種類だけ
    ── 形を取り違えて 2 本赤くした（製品は正しく鳴っていた・実測）。"""
    return [s["種類"] for s in fs.suspect(books)]


def _why_of(books, kind):
    return next(s["理由"] for s in fs.suspect(books) if s["種類"] == kind)


def test_a_fullwidth_number_does_not_hide_a_duplicate():
    """★★ 全角＋長音符と半角＋ハイフンが同じ番号だと分かる（取り逃さない）。"""
    books = {"A.pdf": {**_SAME, "請求番号": "ＩＮＶー０９ー９９９"},
             "B.xlsx": {**_SAME, "請求番号": "INV-09-999"}}
    assert fs.DUPLICATE in _kinds_of(books), _kinds_of(books)


def test_the_two_spellings_are_named_so_a_person_can_judge():
    """★ 畳んだことを黙らない ── 原本の両方を見せる（人が「別物だ」と捨てられる形）。"""
    books = {"A.pdf": {**_SAME, "請求番号": "ＩＮＶー０９ー９９９"},
             "B.xlsx": {**_SAME, "請求番号": "INV-09-999"}}
    why = _why_of(books, fs.DUPLICATE)
    assert "ＩＮＶー０９ー９９９" in why and "INV-09-999" in why, why
    assert "書き方が違います" in why, why


def test_the_same_spelling_gets_no_extra_sentence():
    """★ 陰性対照 ── 書き方が同じなら余計な 1 行を足さない。"""
    books = {"A.pdf": {**_SAME, "請求番号": "INV-09-999"},
             "B.xlsx": {**_SAME, "請求番号": "INV-09-999"}}
    assert "書き方が違います" not in _why_of(books, fs.DUPLICATE)


def test_a_genuinely_different_number_still_does_not_fire():
    """★★ 陰性対照（偽の重複を作らない）── 末尾が 1 違うだけの別の請求書は鳴らさない。"""
    books = {"A.pdf": {**_SAME, "請求番号": "INV-09-999"},
             "B.xlsx": {**_SAME, "請求番号": "INV-09-998"}}
    assert fs.DUPLICATE not in _kinds_of(books), _kinds_of(books)


def test_a_number_that_really_contains_a_prolonged_mark_is_named_not_merged():
    """★ 一番外しそうだと凍結した所 ── 長音符を畳むと、**本当に長音符が入った番号**が
    別の番号と当たりうる。当たっても**原本を両方見せる**ので、人が捨てられる。"""
    books = {"A.pdf": {**_SAME, "請求番号": "ＮＯー１"},
             "B.xlsx": {**_SAME, "請求番号": "NO-1"}}
    assert fs.DUPLICATE in _kinds_of(books)
    why = _why_of(books, fs.DUPLICATE)
    assert "ＮＯー１" in why and "NO-1" in why, why


def test_the_folded_key_never_becomes_the_value():
    """★★ 畳むのは照合のときだけ ── 値（原本）は 1 文字も変えない。"""
    assert fs.number_key("ＩＮＶー０９ー９９９") == "INV-09-999"
    assert fs.number_key("INV-09-999") == "INV-09-999"
    assert fs.number_key("INV-09-998") != fs.number_key("INV-09-999")


# --- 疑いの条件を実務の分布で引き直す（2026-09-13・月末の束 7 束 125 冊で実測）----------
#
# ★★ なぜ変えたか（測った数字がそのまま理由）: 「同じ月・金額が違う」だけで訂正再発行を
#   鳴らしていたので、**同じ取引先が月に何通も出す商慣行**（建設の出来高・運送の便ごと・
#   資材の納品ごと）を疑っていた。C(5,2)=10 組が一斉に立ち
#
#       正 5 ／ ★誤報 156 ／ ★取り逃し 2（どちらも二重払いの典型）
#
#   だった。条件を引き直して **正 7 ／ ★誤報 0 ／ ★取り逃し 0**（同じ検体・同じ採点器）。
# ★ 引き直した形は 2 つ: ① 訂正再発行は**差し替えの徴候**（番号が同じ／枝番／ファイル名の語）
#   が在るときだけ鳴る ② 二重払い（同額・番号違い・日付が近い）という種類を足した。


def _many_in_one_month(n: int) -> dict:
    """同じ取引先が同じ月に n 通、金額は全部違う（★ 商慣行 ── 鳴ってはいけない）。"""
    return {f"梶原建設株式会社_{i}.xlsx":
            _book("梶原建設株式会社", (2026, 7, 5 + i), f"K-70{i}", 30000 + i * 1100)
            for i in range(1, n + 1)}


def test_many_invoices_in_one_month_without_a_sign_are_business_as_usual():
    """★★ 誤報 156 件の真因の番人 ── **分母つき**: 5 通なら組は 10 通りあり、そのどれも
    鳴らない（1 件でも鳴ればオオカミ少年が戻る）。"""
    books = _many_in_one_month(5)
    assert len(books) * (len(books) - 1) // 2 == 10, "分母（組の数）が変わった"
    assert fs.suspect(books) == []


def test_the_same_invoice_number_makes_it_a_reissue_and_leads_the_reason():
    """★ 徴候 a（決め手）── 番号が同じなら疑う。しかも理由文の**頭**で言う
    （良性の行と一字も違わない文にしない ── 買い手は 8 行しか読まない）。"""
    books = _many_in_one_month(2)
    books["梶原建設株式会社_2.xlsx"]["請求番号"] = "K-701"
    found = fs.suspect(books)
    assert _kinds(found) == {(fs.REISSUE, ("梶原建設株式会社_1.xlsx", "梶原建設株式会社_2.xlsx"))}
    why = found[0]["理由"]
    assert "請求番号が同じ" in why and "K-701" in why, why
    assert why.index("請求番号が同じ") < why.index("同じ月"), why


def test_every_resend_word_in_the_filename_counts_as_a_sign():
    """★ 徴候 c ── ファイル名の語は 1 箇所（`RESEND_WORDS`）にしか書かない。
    その全部が本当に効くことを分母つきで見る（表に足したのに効かない語を作らない）。"""
    checked = []
    for word in fs.RESEND_WORDS:
        books = _many_in_one_month(1)
        name = f"梶原建設株式会社_{word}.xlsx"
        books[name] = _book("梶原建設株式会社", (2026, 7, 20), "K-799", 44000)
        got = _kinds(fs.suspect(books))
        assert got == {(fs.REISSUE, ("梶原建設株式会社_1.xlsx", name))}, (word, got)
        checked.append(word)
    assert len(checked) == len(fs.RESEND_WORDS) >= 5, checked


def test_the_same_amount_with_a_different_number_in_one_month_is_a_double_payment():
    """★★ 取り逃し 2 件（一番高くつく形）の番人 ── 別の紙で同じ金額を 2 回。"""
    books = {"小林電機_A01.xlsx": _book("小林電機株式会社", (2026, 6, 10), "INV-3001", 55000),
             "小林電機_A02.xlsx": _book("小林電機株式会社", (2026, 6, 12), "INV-3002", 55000)}
    found = fs.suspect(books)
    assert _kinds(found) == {(fs.DOUBLE_PAY, ("小林電機_A01.xlsx", "小林電機_A02.xlsx"))}
    why = found[0]["理由"]
    assert all(w in why for w in ("同額", "55,000", "請求番号", "INV-3001", "INV-3002")), why


def test_a_monthly_fixed_fee_of_the_same_amount_is_not_a_double_payment():
    """★★ 一番外しそうだと凍結した所 ── 月額固定の保守料（別月・別番号・同額が 3 か月）で
    鳴ったら、買い手の毎月の請求が全部疑いになる。"""
    books = {f"三村保守株式会社_2026-0{m}.xlsx":
             _book("三村保守株式会社", (2026, m, 25), f"M-{m}01", 27500) for m in (6, 7, 8)}
    assert fs.suspect(books) == []


def test_the_double_payment_window_has_both_edges():
    """★ 窓の縁を両側から測る（片側だけだと「常に鳴る」版が素通りする）。"""
    def books(day):
        return {"A.xlsx": _book("福井興業株式会社", (2026, 6, 30), "F-1201", 61000),
                "B.xlsx": _book("福井興業株式会社", (2026, 7, day), "F-1202", 61000)}
    assert fs.DOUBLE_PAY in _kinds_of(books(7)), "7 日違いは鳴る"       # 6/30 → 7/7
    assert fs.suspect(books(8)) == []                                   # 6/30 → 7/8
    assert fs.DOUBLE_PAY_DAYS == 7, "窓は 1 箇所にしか書かない"


def test_the_same_number_is_never_called_a_double_payment():
    """★ 種類を取り違えない ── 番号が同じなら「同じ紙が 2 枚」か「番号の重なり」の話で、
    二重払い（別の紙で 2 回）ではない。買い手の次の一手が違う。"""
    books = {"A.xlsx": _book("福井興業株式会社", (2026, 6, 30), "F-1201", 61000),
             "B.xlsx": _book("福井興業株式会社", (2026, 7, 2), "F-1201", 61000)}
    assert fs.DOUBLE_PAY not in _kinds_of(books), _kinds_of(books)


# --- 請求額が 0 の冊（2026-09-13・買い手役の初見・経理）------------------------------------
#
# ★★ 白紙の雛形（明細 0 行・すべて ¥0）が請求額 0 の行として一覧に載り、合計にそのまま入った。
#   値は原本のまま（0 と書いてある・規則を知らない側の検体 6 冊が「値 0 が正」と宣言している）──
#   黙らせるのは束の所見の仕事。


def test_a_zero_amount_is_named_even_for_a_lone_book():
    books = {"白紙.pdf": _book("あかね商事株式会社", (2026, 9, 1), None, 0)}
    found = fs.suspect(books)
    assert _kinds(found) == {(fs.ZERO_AMOUNT, ("白紙.pdf",))}, found
    assert "0" in found[0]["理由"] and "白紙" in found[0]["理由"], found[0]["理由"]


def test_a_zero_amount_is_named_when_the_vendor_is_unreadable():
    """★ 取引先が読めない冊は取引先ごとの照合から外れるが、0 は 1 冊で分かる。"""
    books = {"白紙.pdf": _book(None, None, None, 0)}
    assert fs.ZERO_AMOUNT in _kinds_of(books)


def test_a_normal_amount_is_not_called_zero():
    assert fs.ZERO_AMOUNT not in _kinds_of(_normal())
    assert fs.suspect(_normal()) == []
