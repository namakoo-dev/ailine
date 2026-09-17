# -*- coding: utf-8 -*-
"""名前が複数行に当たった削除は、断らずに**聞いてから全部消す**（2026-09-07）。

★★ 出所（外部の査定・Namakoo 決裁「削除は聞く。確認が必要だから」）:

    「ヤマノ食品の行を消して」→ ？ できませんでした（4・5 行目の 2 行に当たるため）
    「北斗精機の行を消して」  → 動く（1 行しか当たらない）

  ★ 断り方そのものは親切だった（行番号まで挙げていた）。だが**買い手の目には
    「消してと言ったのに何も起きない」＝失敗**で、しかも直し方が「行番号で言い直す」
    ＝ 2 行なら 2 回打つことになる。★ 言い回しで回避させない。

★ 直した形:
  ・決められないのではなく **聞けば決まる** ── 当たった行を全部挙げて確認を取る
  ・削除は取り返しがつかないので**必ず聞く**（既存の破壊の関所に載せる ──
    新しい関所も新しい exit code も作らない。聞く文だけ差し替える）
  ・生成は**下から順に**消す（上から消すと下の行番号がずれる）
  ・宣言も実体に合わせる（2 行消すなら「行数:2」── ここを 1 のままにすると、
    今日ずっと潰している「宣言と実体のずれ」を自分で作ることになる）
"""
from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl
import pytest

import ailine


@pytest.fixture
def book(tmp_path):
    """★ 『ヤマノ食品』が 2 行・『北斗精機』が 1 行（実物の請求書と同じ形）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "8月請求"
    ws.append(["取引先", "金額"])
    for r in [("丸和物流", 57600), ("ヤマノ食品", 42000),
              ("ヤマノ食品", 18000), ("北斗精機", 114000)]:
        ws.append(list(r))
    p = tmp_path / "b.xlsx"
    wb.save(p)
    return p


def _resolve(book, task):
    return ailine.verify_dsl_args(
        "DELETE_ROWS", {"at": 1, "count": 1}, ailine.build_book_meta(book),
        task=task, vocab=ailine.load_vocab())


def test_a_name_on_several_rows_is_resolved_not_refused(book):
    ok, res, _inf, err = _resolve(book, "ヤマノ食品の行を消して")
    assert ok, err
    assert res["_delete_rows"] == [3, 4], res.get("_delete_rows")
    assert res["count"] == 2, "宣言が実体とずれている（2 行消すのに行数が違う）"
    said = res.get("_confirm_delete") or ""
    assert "ヤマノ食品" in said and "2 行" in said and "3、4行目" in said, said


def test_a_name_on_one_row_does_not_take_the_ambiguity_path(book):
    """★ 対で縛る ── 1 行しか当たらない回は、**曖昧さの関所**には載せない。

    ★★ 2026-09-17（Namakoo 決裁 B）: 初版はここで「聞かずに進む」を縛っていた。
      当時それが正しかったのは、**削除に関所が 1 つも無かった**から ──「今までどおり」は
      「素通り」と同義だった。今日その土台が動いた: 盲検 3 体目が
      「上書きより削除の方が怖いのに、厳しい方が緩い」と指し、消えるものが在る削除は
      すべて関所に載せた（tests/test_deleting_data_asks_first.py）。
    ★ だからこの検体が縛るものを**言い直す** ──「聞くか聞かないか」ではなく
      **どちらの理由で聞くか**。関所は 1 つだが、鳴る理由は 2 つある:
        ① 名前が複数行に当たった（＝どれを消すか決まらない）  ← _delete_rows が立つ
        ② 消えるものが在る（＝取り返しがつかない）            ← 全部の削除に掛かる
      1 行しか当たらない回に ① の文を出したら、それは**嘘の理由**になる。
    ★ 名前で 1 行を指した削除だけ ② を素通りさせる案（A）は採らなかった ── 消える量は
      行番号で指した時と 1 バイトも変わらないのに、**指し方だけで守りが変わる**形になる。
    ★★ どの assert が効いているかを測ってある（変異が 1 本緑だったので追いかけた）:
      ①と②の分かれ目は**構造**で、条件式ではない ── 名前が 1 行に解けると
      `resolve_row_anchor` が行番号を返す（実測: 北斗精機→5 / ヤマノ食品→None）ので、
      ①の枝（`_at_anchor is None`）には**入りようがない**。
      よって効いている assert は `_delete_rows` が立たないこと。下の 2 つ
      （名前が出ない・件数の文である）は**通れない道の見張り**で、意図の記録として置く
      ── 「変異で赤くなる assert」と「読む人のための assert」を混ぜたまま数えない。
    """
    ok, res, _inf, err = _resolve(book, "北斗精機の行を消して")
    assert ok, err
    assert not res.get("_delete_rows"), (
        f"1 行しか当たらないのに『複数行に当たった』側の道に入った: {res.get('_delete_rows')}")
    said = res.get("_confirm_delete") or ""
    assert said, "値の入った行を、聞かずに消そうとしている（② が配線されていない）"
    assert "北斗精機" not in said, f"1 行なのに『名前が複数行に当たった』の文が出た: {said}"
    assert "件あります" in said, f"② の文（消えるものの件数）になっていない: {said}"


def test_the_generator_deletes_from_the_bottom_up(book):
    """★ 上から消すと、その下の行番号がずれる。"""
    bm = ailine.build_book_meta(book)
    _ok, res, _inf, _err = _resolve(book, "ヤマノ食品の行を消して")
    code = ailine.codegen_dsl("DELETE_ROWS", res, book_meta=bm)
    calls = [ln for ln in code.splitlines() if "DeleteRows" in ln]
    assert len(calls) == 2, code
    assert "oDoc, 3, 1" in calls[0] and "oDoc, 2, 1" in calls[1], calls


def test_the_delete_gate_asks_with_the_right_words(monkeypatch, capsys):
    """★ 削除は**必ず聞く**。関所は 1 つのまま、聞く文だけ差し替わること。"""
    a = argparse.Namespace(inplace=True, dry=False, ask=False, overwrite=False, json=False)
    seen = {}

    def _ask(p=""):
        seen["prompt"] = p
        return "y"

    monkeypatch.setattr("builtins.input", _ask)
    assert ailine._confirm_overwrite_or_gate(
        a, "『ヤマノ食品』に当てはまる 2 行（3、4行目）を削除します",
        prompt="削除しますか？") is None
    # ★ 聞く文は input() の引数なので、受け取った文そのものを見る
    assert "削除しますか？" in seen.get("prompt", ""), seen
    assert "上書き" not in seen.get("prompt", ""), seen


def test_saying_no_stops_without_writing(monkeypatch, capsys):
    """★ N と答えたら止まる（原本に触らない）。"""
    a = argparse.Namespace(inplace=True, dry=False, ask=False, overwrite=False, json=False)
    monkeypatch.setattr("builtins.input", lambda _p="": "n")
    assert ailine._confirm_overwrite_or_gate(a, "2 行を削除します",
                                             prompt="削除しますか？") == 1


def test_when_it_cannot_ask_it_shows_a_delete_flavoured_way_out(monkeypatch, capsys):
    """★ 聞けない場（非対話）では止まり、**削除の言葉で**逃げ道を示すこと。

    ★ ここが「上書きを承知して続行する」のままだと、削除なのに上書きの話をする。
    """
    def _eof(_p=""):
        raise EOFError
    a = argparse.Namespace(inplace=True, dry=False, ask=False, overwrite=False, json=False)
    monkeypatch.setattr("builtins.input", _eof)
    assert ailine._confirm_overwrite_or_gate(a, "2 行を削除します",
                                             prompt="削除しますか？") == 7
    said = capsys.readouterr().out
    assert "削除を承知して続行する" in said, said
    assert "上書きを承知して" not in said, said


def test_the_postcondition_honours_the_row_list(tmp_path, book):
    """★ 検算も宣言の一覧を読む（at+count だけを見ると、非連続で嘘になる）。"""
    from ailine_core.postconditions import move
    wb = openpyxl.load_workbook(book)
    ws = wb["8月請求"]
    ws.delete_rows(4)
    ws.delete_rows(3)
    after = tmp_path / "after.xlsx"
    wb.save(after)
    st, msg = move.check_delete_rows(
        after, {"at": 3, "count": 2, "_delete_rows": [3, 4], "_target_sheet": "8月請求"},
        header_row=1, source_book=book)
    assert st in ("pass", "warn"), (st, msg)
