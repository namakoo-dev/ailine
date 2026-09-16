# -*- coding: utf-8 -*-
"""★★ 依頼文の**動詞**を、シートの名指しと読まない（2026-09-16・買い手 2 体目の ④⑤）。

★ 出所: ブックに『集計』シートが在ると、「担当者ごとに金額を集計して」の『集計』が
  シートの言及として採られ、**一度『集計』を作ったら二度と「集計して」と言えなく**なった。
  買い手の言葉:「事務の言葉づかいでは避けようがない」── 月末に 2 種類の集計は必ず作る。

★★ 片配線だった: 朝に器官（サ変動詞の判定）は作ったが、配線したのは
  **対象シートを決める側**（`target_sheet.resolve_target_sheet`）だけ。
  `subject.py` は『集計』を「依頼文が指している語」に数え続けたので、
  対象は正しく『受注』を選ぶのに **⚠ が立って `[y/N]` の関門で止まった**（実機で再現）。
  ★ 直しは「両方に書く」ではなく **`subject.designates_a_sheet` に 1 本畳んで両方が呼ぶ**。

★ 線は狭い（実測 13/13）:
    マーカー付き（「集計シートを」）      → 名指し（無条件で勝つ）
    助詞が続く（「集計を」「集計の」「集計に」）→ 名指し（動詞ではない）
    直後がサ変（「集計して」「集計する」）    → 名指しではない
"""
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core import subject, target_sheet  # noqa: E402

SHEETS = ["受注", "集計", "転記", "検分"]
HEADERS = {"受注": ["受注番号", "取引先", "担当者", "金額"],
           "集計": ["取引先", "合計 - 金額"],
           "転記": ["コード", "名前"],
           "検分": ["行", "枚数"]}

#: 事故の形 ── 名前がそのまま動詞。シートの名指しではない。
AS_VERB = [
    ("担当者ごとに金額を集計して", "集計"),
    ("取引先ごとに集計して", "集計"),
    ("部門ごとに金額を集計する", "集計"),
    ("担当者ごとに集計しよう", "集計"),
    ("コードで名前を転記して", "転記"),
]

#: 陰性対照 ── ここまで動詞と読んだら、シートを名指しできなくなる。
AS_NAME = [
    ("集計シートを並べ替えて", "集計"),
    ("集計シートの金額を太字にして", "集計"),
    ("集計を並べ替えて", "集計"),
    ("集計の金額を3桁区切りにして", "集計"),
    ("集計に合計行を追加して", "集計"),
    ("受注の金額を並べ替えて", "受注"),
    ("転記シートを見せて", "転記"),
    ("検分を確認して", "検分"),
]


def test_the_judgement_lives_in_one_place():
    """★★ 片配線の番人 ── 判定は 1 本で、両方の消費側が**同じ関数**を呼ぶこと。

    ★ 朝は同じ規則が 2 か所に在り、片方だけ直っていた。数を数えるのではなく
      **同一性**で縛る（書き写しが復活したら False になる）。
    """
    assert target_sheet._used_as_a_verb is subject.used_as_a_verb, (
        "動詞の判定が書き写されている（subject.py の 1 本を呼ぶこと）")
    assert target_sheet._mentioned_with_marker is subject.mentioned_with_marker


@pytest.mark.parametrize("task, name", AS_VERB)
def test_a_verb_does_not_designate_a_sheet(task, name):
    assert not subject.designates_a_sheet(task, name), task


@pytest.mark.parametrize("task, name", AS_NAME)
def test_a_real_mention_still_designates_a_sheet(task, name):
    assert subject.designates_a_sheet(task, name), task


@pytest.mark.parametrize("task, name", AS_VERB)
def test_the_target_sheet_is_not_stolen_by_the_verb(task, name):
    """★ 対象を決める側（朝に直した所）も、引き続き 1 枚目を選ぶこと。"""
    got, _src, err, _conflict = target_sheet.resolve_target_sheet(
        task, SHEETS, None, headers=HEADERS)
    assert err is None, err
    assert got == "受注", f"{task} → {got}"


@pytest.mark.parametrize("task, name", AS_VERB)
def test_the_verb_is_not_counted_as_something_the_task_points_at(task, name):
    """★★ ここが ④⑤ の本体 ── designator に数えると ⚠ が立ち `[y/N]` で止まる。

    ★ 朝の直しはこの経路に届いていなかった。対象は正しいのに実行できない、という
      いちばん質の悪い形（画面は正しく見えるので、原因が分からない）。
    """
    d = subject.task_designators(task, columns=HEADERS["受注"], sheets=SHEETS)
    assert name not in d.sheets, (
        f"依頼文の動詞『{name}』をシートの名指しとして数えている: {task} → {d.sheets}")


def test_a_named_sheet_is_still_counted_as_pointed_at():
    """★ 陰性対照 ── 本当に名指ししたシートは designator に残ること。"""
    d = subject.task_designators("集計シートを並べ替えて",
                                  columns=HEADERS["集計"], sheets=SHEETS)
    assert "集計" in d.sheets, d.sheets


# ---------------------------------------------------------------------------
# ★★ 変異試験で見つけた検体の穴（2026-09-16・その場で埋めた）
#
# 上の検体だけでは、次の 2 つの変異が**緑のまま**通った ──
#   M3 スロット判定（`_match_slot` の SHEET 枝）から配線を外す
#   M4 マーカーの枝を無条件で負けさせる
# どちらも「その経路を通る文が 1 つも無かった」のが理由。検体の側を直す。
# ---------------------------------------------------------------------------

def _verdict_for_sheet(task: str, sheet: str, columns):
    """SHEET スロット 1 件を仕分けて、①②③ のどれになったかを返す。"""
    slot = subject.Slot(key="_target_sheet", value=sheet, kind=subject.SHEET)
    got = subject.classify_slots([slot], task=task, columns=columns, sheets=SHEETS)
    return got[0]


@pytest.mark.parametrize("task, name", AS_VERB)
def test_the_verdict_does_not_claim_the_task_named_that_sheet(task, name):
    """★ M3 の穴 ── スロット判定（`_match_slot` の SHEET 枝）にも同じ線が要る。

    ★ ここが抜けると、**動詞を「そのシートを名指しした」と読む**。
      対象がたまたまその名前のシートだった回に、出るべき ⚠ を黙らせてしまう
      （断りを消す向きの壊れ方なので、画面には何も出ず気づけない）。
    """
    v = _verdict_for_sheet(task, name, HEADERS["受注"])
    assert v.tier != subject.MATCHED, (
        f"動詞『{name}』を「依頼文がこのシートを名指しした」と読んでいる: {task}")


@pytest.mark.parametrize("task, name", AS_VERB)
def test_no_warning_is_raised_about_the_real_target(task, name):
    """★★ 買い手の症状そのもの ── 対象『受注』に ⚠ が立たないこと。

    ★ 実機では、この ⚠ が `[y/N]` の関門を呼び、**2 つ目の集計が実行できなかった**。
    """
    v = _verdict_for_sheet(task, "受注", HEADERS["受注"])
    assert v.tier != subject.CONTRADICTED, (
        f"対象『受注』に矛盾ありと判定している（⚠ が立つ）: {task} → 残り {v.designators}")
    assert not subject.contradiction_lines([v]), subject.contradiction_lines([v])


def test_a_marked_mention_wins_even_when_the_same_word_is_also_a_verb():
    """★★ M4 の穴 ── マーカーと動詞が**同居**する文が検体に無かった。

    「集計シートで担当者ごとに集計して」── 人は『集計』シートを名指ししつつ、
    同じ語を動詞としても使う。マーカーの枝が無いと名指しが消え、
    **そのシートを対象にできなくなる**（動詞の側だけ見ると取りこぼす）。
    """
    task = "集計シートで担当者ごとに集計して"
    assert subject.used_as_a_verb(task, "集計"), "この文は動詞としても使っている（前提）"
    assert subject.mentioned_with_marker(task, "集計"), "マーカーも在る（前提）"
    assert subject.designates_a_sheet(task, "集計"), (
        "マーカー付きの名指しが動詞に負けている ── 『集計シート』を対象にできない")
    d = subject.task_designators(task, columns=HEADERS["集計"], sheets=SHEETS)
    assert "集計" in d.sheets, d.sheets
