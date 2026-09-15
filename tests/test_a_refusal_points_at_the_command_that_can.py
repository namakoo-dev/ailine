# -*- coding: utf-8 -*-
"""断りが「別のコマンドで出来る」と言うとき、そのコマンドは実在するか（2026-09-08）。

★★ 出所（盲検の検品が挙げた摩擦）:

    依頼 「この表を **PDF** にして」
    旧   `？ …照合できませんでした。**要望として記録します**。`
    実物 `ailine export-pdf` は**在る**

  ★ 持っている物を持っていないと言う形。断り自体は正しい（`run` の一覧には無い）ので、
    直すのは操作ではなく**導線**。

★ この番人は 2 つを縛る:
    ① 案内するコマンドが `ailine --help` の一覧に**実在する**こと
       ── 在らぬ道具へ人を送るのは、断るより悪い
    ② 案内が出る回は「要望として記録します」を**言わない**こと
"""
from __future__ import annotations

import ailine
from ailine_core import route
from ailine_core.cli_render import freeform_notice_reason, render_vocab_miss_refusal
import pytest


def _subcommands() -> set:
    """`ailine` が実際に持つサブコマンド名（parser から取る・手書きしない）。"""
    parser = ailine.build_parser()
    names = set()
    for act in parser._subparsers._group_actions:      # noqa: SLF001
        names |= set(getattr(act, "choices", {}) or {})
    return names


def test_every_route_points_at_a_command_that_exists():
    have = _subcommands()
    assert have, "サブコマンドが 1 つも取れなかった（parser の組み方が変わった？）"
    missing = [cmd for _w, cmd, _n in route.ROUTES
               if cmd.split()[-1] not in have]
    assert not missing, f"実在しないコマンドへ案内している: {missing}"


def test_the_refusal_names_the_command_instead_of_filing_a_wish():
    shown = "\n".join(render_vocab_miss_refusal("PDF出力", task="この表をPDFにして"))
    assert "ailine export-pdf" in shown, shown
    assert "要望として記録します" not in shown, shown


def test_a_request_no_command_can_do_still_gets_the_old_refusal():
    """★ 対の試験 ── 本当に出来ない依頼まで案内してしまわない。"""
    shown = "\n".join(render_vocab_miss_refusal("画像挿入", task="ロゴ画像を入れて"))
    assert "要望として記録します" in shown, shown
    assert "ailine export-pdf" not in shown, shown


def test_the_environment_failure_path_is_untouched():
    """★ ollama 不通の経路は語彙の話ではないので、案内を混ぜない（過去の致命①）。"""
    shown = "\n".join(render_vocab_miss_refusal(
        "", translate_error=True, task="この表をPDFにして"))
    assert "ailine doctor" in shown, shown
    assert "ailine export-pdf" not in shown, shown


# ── ★ 2026-09-16: 断りが「その機能は無い」と**断定しない**こと（盲検の買い手役⑦）──────
#
# ★★ 事故（買い手の引用）:
#     「発注点を割っている商品に「要発注」と印を付けて」
#       ？ この依頼（印を付ける）は、頼める操作の一覧に照合できませんでした。要望として記録します。
#     ところが `ailine ops` には、はっきりこう書いてある:
#       SET_WHERE  条件つき書換    こう書く: …／〜以上の行に印を付ける／…
#     ★ 買い手「**一覧に「印を付ける」と書いてあるのに、一覧に無いと断られた**」。
#
# ★ 実測して分かったこと: **目録は壊れていない**。「こう書く」の 90 本はそのまま打てば全部届く。
#   届かなかったのは言い方の**変種**だった。だから直すのはプールではなく**言葉**:
#   見たことと解釈を分ける（`_refuse_output_conflict` が 2026-08-26 に通ったのと同じ線）。
#   分かるのは「この依頼文からは操作を決められなかった」ことだけで、
#   「その操作が無い」は**こちらが足した解釈**であり、しかも嘘だった。

_ABSENCE_CLAIM = "一覧に照合できませんでした"


@pytest.mark.parametrize("about,task", [
    ("印を付ける", "発注点を割っている商品に「要発注」と印を付けて"),
    ("", "在庫の様子をいい感じにまとめて"),
    ("色分けする", "担当者ごとに色分けして"),
])
def test_a_refusal_never_claims_the_feature_is_absent(about, task):
    lines = render_vocab_miss_refusal(about=about, task=task)
    joined = "".join(lines)
    assert _ABSENCE_CLAIM not in joined, (
        f"『無い』と断定している（目録には在るかもしれない）:\n{joined}")
    assert "決められませんでした" in joined, joined


def test_the_refusal_still_shows_the_way_forward():
    """★ 断定をやめたぶん、**次の一手**は必ず残す（黙って弱くしない）。"""
    lines = render_vocab_miss_refusal(about="印を付ける", task="発注点を割っている商品に印を付けて")
    joined = "".join(lines)
    assert "ailine ops" in joined, joined
    assert "言い方を変えると" in joined, joined


def test_every_phrase_the_catalog_advertises_reaches_its_own_op():
    """★★ 目録と照合の片配線の番人: `ailine ops` が「こう書く」と見せた文を**そのまま打ったら**、
       その op に届くこと。★ 2026-09-16 の実測は 90/90 ── ここが割れた日に、
       買い手は「一覧に書いてあるのに通らない」を踏む（⑦はその**手前**の言葉の問題だった）。
    """
    misses = []
    for op, meta in ailine.OP_META.items():
        for phrase in meta.get("synonyms") or []:
            if op not in (ailine.suggest_ops(phrase) or []):
                misses.append((op, phrase))
    assert not misses, f"目録が見せている言い方が自分の op に届かない: {misses}"


@pytest.mark.parametrize("op,about", [
    ("OUT_OF_VOCAB", "印を付ける"),
    ("OUT_OF_VOCAB", ""),
])
def test_the_plan_path_refusal_also_never_claims_absence(op, about):
    """★ 同じ言葉を**複合計画の語彙外段**も使う（`freeform_notice_reason`）。
       ★ 初版はこちらに検体が無く、変異を当てても緑のままだった ── 片方だけ縛らない。"""
    line = freeform_notice_reason(op, about)
    assert _ABSENCE_CLAIM not in line, line
    assert "決められませんでした" in line, line


def test_a_translation_that_never_became_a_command_says_so_plainly():
    """★ 陰性対照 ── 翻訳がそもそも形にならなかった経路は、別の文のまま（言葉を混ぜない）。"""
    line = freeform_notice_reason("FREEFORM", "")
    assert "頼める操作の形になりませんでした" in line, line
