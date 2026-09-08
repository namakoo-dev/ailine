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
from ailine_core.cli_render import render_vocab_miss_refusal


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
