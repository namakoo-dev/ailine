# -*- coding: utf-8 -*-
"""台帳は画面と**同じ語**を言う（2026-09-18・盲検 4 体目 ①）。

★★ 起きたこと（記帳代行の担当者・初見・盲検）:

    画面:  △ 使い捨て2.out.xlsx は宣言どおりの変化を確認しました ── ただし ⚠ 1 件を…
    直後の ailine history:
           2026-09-18T08:43:20  ✓  1  qwen2.5-coder:7b  使い捨て2.xlsx  セル E1 に 借方金額 と入力して

  ★ 4 回とも再現。台帳の実体にも `"ok": true` としか残らず、**⚠ の痕跡が無い**。
  ★ 買い手の言葉:「顧問先の帳簿を触る道具で、**作業記録が嘘をつく**のは致命的です。
    税務調査で『この仕訳は何を根拠に入れたか』を説明する時、私はこの履歴を見ます。
    『機械検証済み』と書いてあるのに実際は警告つきだった、では使えません。
    **この 1 点だけで、私は所長に導入を提案できません**」。

★★ 根はまたこの形だった ── 器官は在るが配線が無い:
  `verdict` は `_finish_apply` が**1 箇所で正しく決めて**いる
  （verified / warned / unverified / unobservable）。コメントにも
  **「決めるのは 1 箇所・映すのは何箇所でも、という形にする」**と書いてある。
  ところが **build_history_entry が verdict を 1 つも載せていなかった** ──
  映す側は読みようが無く、別の真偽値（`ok` ＝適用が通ったか）で二値に潰していた。
  ★ `ok` と「機械検証済みか」は**意味が違う**。画面と台帳で別の判断をしていた。

★ 直しは決めた所を**運ぶ**こと（新しい判断を作らない）。
★ 古い行（verdict を持たない）は従来どおり ok で読む ── 過去の台帳を遡って書き換えない。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _row(**kw):
    base = {"ts": "2026-09-18T00:00:00+00:00", "ok": True, "attempts": 1,
            "model": "m", "book": "b.xlsx", "task": "なにか"}
    base.update(kw)
    return base


def test_a_warned_run_is_not_recorded_as_verified():
    """★ 事故そのもの: 画面が △ の run を、台帳が ✓ と書かない。

    ★★ 2026-09-18: 初版は `_history_mark` を**直に呼んで**いて、
      **表がそれを使っているか**を見ていなかった ── 表を元の
      `"✓" if e.get("ok") else "×"` に戻す変異が**緑のまま**通った。
      買い手が見たのは表そのもの。器官でなく**出てくる画面**で縛る。
    """
    assert ailine._history_mark(_row(verdict="warned", warning_count=1)) == "△"
    said = ailine.format_history_table([_row(verdict="warned", warning_count=1)])
    assert " △ " in said, f"表が △ を出していない:\n{said}"
    assert " ✓ " not in said, f"表が ✓ と書いている（事故そのもの）:\n{said}"


def test_every_verdict_has_its_own_mark():
    """★★ 2 値に潰さない ── 画面が出し分ける語を、台帳も出し分ける。

    ★ ここが無いと「△ も ⚠ も ✓ でないから ×」のような潰し方でも上の試験が通る。
      × は「適用していない」の印で、「疑わしい」の印ではない。
    """
    seen = {v: ailine._history_mark(_row(verdict=v)) for v in
            ("verified", "warned", "unverified", "unobservable", "not_applied")}
    assert seen == {"verified": "✓", "warned": "△", "unverified": "⚠",
                    "unobservable": "⚠", "not_applied": "×"}, seen


def test_an_old_row_without_a_verdict_still_reads():
    """★ 後方互換: verdict を持たない古い行は従来どおり ok で読む。

    ★ 過去の台帳を遡って書き換えない（記録は記録のまま）。判定材料が無い回に
      「分かる」の顔をしない、という線でもある。
    """
    assert ailine._history_mark(_row(ok=True)) == "✓"
    assert ailine._history_mark(_row(ok=False)) == "×"


def test_the_verdict_is_actually_carried_into_the_ledger():
    """★★ 決めた所が**台帳まで運ばれている**こと（ここが抜けていたのが事故の根）。

    ★ 表示だけ直して台帳に載せないと、`--json` や後から読む側が同じ嘘を読む。
    """
    entry = ailine.build_history_entry(
        {"ok": True, "verdict": "warned", "warning_count": 1}, Path("b.xlsx"),
        "なにか", "m", "none")
    assert entry["verdict"] == "warned", entry
    assert entry["warning_count"] == 1, entry


def test_the_screen_and_the_ledger_use_one_decision():
    """★★ 判断を 2 つ作らない ── 台帳の側で verdict を組み直していないこと。

    ★ `_finish_apply` のコメントが言っているとおり「決めるのは 1 箇所・映すのは何箇所でも」。
      ここで `warning_count > 0` などを**書き写す**と、それが 2 つ目の実装になる
      （この repo が何度も踏んでいる形）。
    """
    from _product_source import count_in_product, window_around
    assert count_in_product('result["verdict"] = (') == 1, (
        "★ verdict を決める場所が 2 つ以上ある")
    # ★★ 読む場所を**決め打ちしない** ── tests/test_guard_ledger.py が
    #   「本体を場所で決め打ちする番人が増えた」と鳴る（2026-09-18 に**この日 2 度目**踏んだ）。
    #   2026-09-03 に事後条件を ailine_core/ へ移した時、同じ形で番人 7 件が一斉に空振りした。
    body = window_around("def _history_mark(", after=1200)
    assert body, "★ 探す場所が空（この検査が空回りする）"
    assert "warning_count" not in body, (
        "★ 台帳の側で判断を組み直している（決めた verdict を運ぶだけにすること）")
