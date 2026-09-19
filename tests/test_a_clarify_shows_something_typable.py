# -*- coding: utf-8 -*-
"""聞き返しの行き止まりに、**そのまま打てる文**が在る（2026-09-20）。

★★ 出所（導通の盤が vague と判定した 3 件のうち 2 件 ── 本とフォルダの聞き返し）:
  行き止まりに置いてあったのは「（頼める操作の一覧: ailine ops）」だけだった。
  一覧は**操作の名前**の並びであって、**そのまま打てる依頼文**ではない。
  買い手は「一覧を見ても、この依頼が通る形が分からない」ところで止まる ──
  合格線の 3 条目（通る道を示す）を満たしていない。

★★ 盲検査定 A の実測がこの形を先に指していた: 語彙外の依頼を 4 回言い直して 4 回とも
  質問返しになり「普通の購入検討者ならここで評価を終える」。

★★ フォルダ側は `replace_examples_in_question` すら呼んでいなかった ── 本の経路は
  呼んでいたので**同じ判断の片配線**。だから 1 本の試験で**両方の経路**を縛る
  （片方だけ直す変異が緑で通らないように）。
"""
from __future__ import annotations

import contextlib
import io
import sys
import tempfile
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

import openpyxl  # noqa: E402

import ailine  # noqa: E402
import walk_refusals_core as walk  # noqa: E402

QUESTION = "何をしますか"


def _clarify_screen(kind: str) -> str:
    """CLARIFY を返す翻訳に固定して、本／フォルダ それぞれの行き止まりの画面を取る。"""
    plan = [{"op": "CLARIFY", "question": QUESTION}]
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        if kind == "folder":
            base = walk._make_folder(root)
            argv = ["run", str(base), "いい感じにして", "--out", str(root / "結果.xlsx")]
        else:
            base = walk._make_book(root)
            argv = ["run", str(base), "いい感じにして", "--copy"]
        rc, out = walk._run(argv, plan)
    assert rc == 3, f"★ 引き金が引けていない（exit {rc}）:\n{out}"
    return out


@pytest.mark.parametrize("kind", ["book", "folder"])
def test_the_dead_end_shows_a_request_you_can_type(kind):
    """★★ 事故そのもの: 行き止まりに**そのまま打てる文**が在ること（両経路）。"""
    out = _clarify_screen(kind)
    got = walk._example_in(out)
    assert got, f"★ そのまま打てる文が無い（一覧だけ）:\n{out}"


@pytest.mark.parametrize("kind", ["book", "folder"])
def test_the_dead_end_still_offers_the_way_to_tell_them_apart(kind):
    """★ 出口（一覧）は消さないこと ── 「言い方が悪い」と「対応していない」を分ける手段。"""
    assert "ailine ops" in _clarify_screen(kind)


def test_the_folder_dead_end_only_shows_what_folders_can_do():
    """★★ フォルダで通るのは抽出だけ ── 通らない例を見せない（導線が嘘なら無い方が悪い）。

    ★ 「けい線を引いて」のような 1 冊向けの例をフォルダで見せると、そのまま打って断られる。
    ★ 見せる op は `OP_META` の folder 宣言から導く（手書きの対応表を持たない）。
    """
    out = _clarify_screen("folder")
    from ailine_core.examples import example_task_for
    for op in ailine.OP_META:
        if (ailine.OP_META[op] or {}).get("folder"):
            continue
        ex = example_task_for(op)
        if ex:
            assert ex not in out, f"★ フォルダで通らない例を見せている（{op}）:\n{out}"


def test_the_folder_ops_come_from_the_declaration():
    """★ 対応表を手で書いていないこと ── 宣言が唯一の出どころ。"""
    assert ailine._folder_example_ops() == [
        op for op, m in ailine.OP_META.items() if (m or {}).get("folder")]
    assert ailine._folder_example_ops(), "★ フォルダで通る op が 1 つも無い（宣言が読めていない）"


def _clarify_branches() -> list:
    """製品の中で **CLARIFY を扱っている枝**を AST で取り出す（手で並べない）。

    ★ 数を凍結しない ── 枝が増えたらこの試験が自動でその枝も縛る。
    """
    import ast
    src = (REPO / "src" / "ailine" / "__init__.py").read_bytes().decode("utf-8")
    tree = ast.parse(src)
    out = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.If):
            continue
        test = ast.get_source_segment(src, node.test) or ""
        if '"CLARIFY"' in test and "op ==" in test:
            out.append((node.lineno, "\n".join(
                ast.get_source_segment(src, s) or "" for s in node.body)))
    return out


def test_every_clarify_dead_end_goes_through_the_one_exit():
    """★★ 片配線を作らない ── 出口の文面は 1 関数に置き、**質問を出す枝は全部**それを呼ぶ。

    ★ ここが緩むと、片方の経路だけ直して他方が古いまま残る（この repo が何度も踏んだ形。
      実際フォルダ側は `replace_examples_in_question` を呼んでいなかった）。
    ★ 枝は AST で数える ── 手で並べた表は、枝が増えた日に黙って古くなる。
    ★ 質問を画面に出さない枝（多段の計画のプレビュー）は対象外 ── 行き止まりでない。
    """
    branches = _clarify_branches()
    assert branches, "★ CLARIFY の枝が 1 つも見つからない（この試験が空回りしている）"
    bad = [ln for ln, body in branches
           if 'print(f"？ ' in body and "_print_clarify_exit(" not in body]
    assert not bad, f"★ 出口を呼んでいない聞き返しの枝がある（行 {bad}）"
    leaked = [ln for ln, body in branches
              if 'print("  （頼める操作の一覧: ailine ops）")' in body]
    assert not leaked, f"★ 出口の文面を枝に書き写している（行 {leaked}）── 1 関数に畳むこと"


def test_an_example_already_in_the_question_is_not_repeated():
    """★ 質問文が既に「（例: …）」を持つ回は足さない ── 同じものを 2 度見せない。"""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        ailine._print_clarify_exit("何をしますか（例: 「けい線を引いて」）")
    assert "そのまま打てます" not in buf.getvalue(), buf.getvalue()
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        ailine._print_clarify_exit("何をしますか")
    assert "そのまま打てます" in buf.getvalue(), buf.getvalue()
