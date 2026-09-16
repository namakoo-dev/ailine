# -*- coding: utf-8 -*-
"""見るだけの依頼を、行や列を減らす操作に写さない（2026-09-14・言い回し 120 件の盲検・家系②）。

★★ 「品番の重複がないか**見といて**」→ 重複除去／「〜がないか**調べて**」→ 重複除去。
  既定は新しい冊に書くので原本は壊れないが、出るのは「重複を除いた表」で、**どれが重複だったか
  分からない**（頼んだことの逆側の表）。
★ 一方、確認語が**抽出**（読むだけ）に着いた依頼は盲検で 9/9 正 ── 規則は狭く取る:
  確認語が在り、変更の動詞が 1 つも無い依頼を、減らす op（重複除去・行削除・列削除）に写さない。
"""
from __future__ import annotations

import inspect
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine   # noqa: E402
from ailine_core import intent   # noqa: E402

META = {"sheets": ["表"], "headers": {"表": ["品番", "品名", "区分", "在庫数"]},
        "header_rows": {"表": 1}}


def _dedup(task):
    return ailine.verify_dsl_args("DEDUP", {"keys": ["品番"]}, META, task=task)


# ── 器官 ──────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("task,word", [
    ("品番の重複がないか見といて", "ないか"),
    ("品名が同じなのに品番が違うものがないか調べて", "ないか"),
    ("在庫数がマイナスの行を教えて", "教えて"),
])
def test_a_check_only_request_is_recognised(task, word):
    assert intent.check_only_request(task) == word


@pytest.mark.parametrize("task", [
    "品番が同じ行を重複として除いて",
    "重複がないかチェックして、あったら消して",     # 変更の動詞が在れば見るだけではない
    "在庫数を確認して0の行を削除して",
])
def test_a_request_with_a_change_verb_is_not_check_only(task):
    assert intent.check_only_request(task) is None, task


# ── 関所（盲検の 2 件がそのまま検体）────────────────────────────────────────

def test_looking_for_duplicates_is_not_dedup():
    ok, _r, _i, err = _dedup("品番の重複がないか見といて")
    assert not ok, "見るだけの依頼を重複除去に写した"
    assert "見るだけの依頼" in err and "重複除去" in err, err
    assert "品番が同じ行を重複として除いて" in err, "★ 減らしてよいときの通る書き方を添える"


def test_asking_whether_duplicates_exist_is_not_dedup():
    ok, _r, _i, err = _dedup("品名が同じなのに品番が違うものがないか調べて")
    assert not ok and "見つけるだけの操作はありません" in err, err


def test_an_explicit_removal_still_passes():
    """陰性対照 ── 減らすと言っている依頼は今までどおり通る。"""
    ok, r, _i, err = _dedup("品番が同じ行を重複として除いて")
    assert ok, err
    assert r["keys"] == ["品番"]


def test_a_check_with_a_change_verb_still_passes():
    ok, _r, _i, err = _dedup("重複がないかチェックして、あったら消して")
    assert ok, err


def test_extract_is_not_gated():
    """★ 抽出は読むだけ ── 確認語が在っても止めない（盲検で 9/9 正だった経路を壊さない）。"""
    ok, r, _i, err = ailine.verify_dsl_args(
        "EXTRACT", {"col": "在庫数", "cmp": "lt", "value": 0}, META,
        task="在庫数がマイナスになってる行を教えて")
    assert ok, err
    assert (r["cmp"], r["value"]) == ("lt", 0.0)


def test_delete_rows_and_delete_column_are_gated_the_same_way():
    """★ 減らす op の 3 つに同じ関所（op ごとに書き分けない）。"""
    for op, args in (("DELETE_ROWS", {"at": 2}), ("DELETE_COLUMN", {"col": "区分"})):
        ok, _r, _i, err = ailine.verify_dsl_args(op, args, META, task="区分が空の行がないか調べて")
        assert not ok and "見るだけの依頼" in err, (op, err)


def test_the_gate_sits_where_both_routes_pass():
    """★ 片配線の番人 ── 単発も複合計画の段も通る `verify_dsl_args` の入口に 1 箇所。"""
    src = inspect.getsource(ailine.verify_dsl_args)
    assert src.count("refuse_reducing_a_check(") == 1, "関所は 1 箇所"
    assert "if task and" in src.split("refuse_reducing_a_check(")[0][-200:], \
        "★ 依頼文が無い経路（DSL 直渡し）は触らない"


def test_no_task_means_no_gate():
    ok, _r, _i, err = ailine.verify_dsl_args("DEDUP", {"keys": ["品番"]}, META, task="")
    assert ok, err


def test_every_reducing_op_has_its_own_way_to_ask():
    """★ 名簿の全 op に「減らしてよいときの通る書き方」が在ること。

    ★★ 2026-09-16 の掃き出しで見つけた**番人の芽**: `refuse_reducing_a_check` は
      `REDUCING_EXAMPLES.get(op, '…を削除して')` と**黙って一般文に落ちる**ので、
      `ROW_REDUCING_OPS` に op を足して文例を足し忘れても、どこも赤くならなかった。
      断り文だけが「…を削除して」という、その人の依頼と噛み合わない案内に痩せる。
    ★ 等号で縛る（⊇ ではなく ==）── 使われない文例が残るのも腐りなので両方向を見る。
    """
    from ailine_core import intent as intent_mismatch
    assert set(intent_mismatch.REDUCING_EXAMPLES) == set(intent_mismatch.ROW_REDUCING_OPS), (
        "減らす op の名簿と、その通る書き方の名簿がずれている: "
        f"文例だけ在る={sorted(set(intent_mismatch.REDUCING_EXAMPLES) - set(intent_mismatch.ROW_REDUCING_OPS))} / "
        f"文例が無い={sorted(set(intent_mismatch.ROW_REDUCING_OPS) - set(intent_mismatch.REDUCING_EXAMPLES))}")
