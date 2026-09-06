# -*- coding: utf-8 -*-
"""2 組目の条件は「条件の形で書かれている」ものだけ採る ── その番人（2026-09-06）。

★★ 出所: 自作 review 3 戦目が致命 3 件・重大 1 件として拾った。どれも根は 1 つで、
  2 組目を「残差に残った語 ∩ 他の列の実在値」だけで決めていて、**その語が条件として
  書かれているか**を見ていなかった。実測で再現した形:

    ① 否定 × 他列の同値   「所属が営業でない行のメモに『済』」（担当列にも『営業』）
                          → cond2=担当/営業 が勝手に付き、2 行 → **1 行**に縮んだ
                          ★ 根は残差が**リストの値を消費しない**こと（否定の cond_value は list）
                          ★ この欠陥は同じ日に**測定器の中で見つけて直していた**のに、
                            本体へ持ち帰らなかった（bench/residue_gate_probe.py）
    ② 列挙（OR）          「金額が1000以上で経理か営業の行…」
                          → どちらの枝にも入らず cond2 未設定のまま**1 条件で走った**
    ③ 事後条件            cond2_col が実表に無いと、2 組目が丸ごと無視されて **pass**
    ④ 付随的な固有名詞    「田中さんのように、所属が営業の行…」
                          → cond2=氏名/田中 が付き、3 行 → **1 行**（警告なし）

★ 直しは今日の否定と同じ形 ── 文に在るかではなく、**何に付いているか**を見る。
  採らなかった候補は**黙らない**（⚠ を出して ✓ を降ろす・直しはしない）。
"""
from __future__ import annotations

import openpyxl
import pytest

import ailine
from ailine_core import residue


def _book(tmp_path, name, headers, rows, sheet="名簿"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for r in rows:
        ws.append(r)
    p = tmp_path / f"{name}.xlsx"
    wb.save(p)
    return p


@pytest.fixture
def meibo(tmp_path):
    """★ 氏名にも所属にも『営業』が居る表（① と ④ を同時に踏める）。"""
    return _book(tmp_path, "meibo", ["氏名", "所属", "担当", "メモ"],
                 [("田中", "営業", "主任", ""), ("鈴木", "経理", "副任", ""),
                  ("佐藤", "営業", "副任", ""), ("山田", "営業", "主任", "")])


def _set_where(book, task, cmp_="eq"):
    return ailine.verify_dsl_args(
        "SET_WHERE", {"col": "メモ", "cond_col": "所属", "cmp": cmp_, "value": "○"},
        ailine.build_book_meta(book), task=task, vocab=ailine.load_vocab())


# --- 判定そのもの（純ロジック）---------------------------------------------

@pytest.mark.parametrize("task, col, val, want", [
    ("所属が営業で担当が主任の行のメモに「○」を付けて", "担当", "主任", True),
    ("担当が『主任』の行のメモに「○」を付けて", "担当", "主任", True),      # ★ 括弧を挟んでも
    ("田中さんのように、所属が営業の行のメモに「○」を付けて", "氏名", "田中", False),
    ("所属が営業でない行のメモに「済」を付けて", "担当", "営業", False),
    ("金額が1000以上で経理か営業の行の備考に「○」を付けて", "部門", "経理", False),
])
def test_a_value_counts_only_when_written_as_a_condition(task, col, val, want):
    assert residue.stated_as_condition(task, col, val) is want


def test_the_residue_consumes_values_inside_a_list():
    """★★ ① の根 ── 否定の `cond_value` は list になる。文字列しか消費しないと、
    否定した値が残差に残って別の列の条件として拾われる。

    ★ この欠陥は同じ日に測定器の中で直していた。**測定器と本体は必ず両方見る。**
    """
    args = {"cond_col": "所属", "cond_value": ["営業"], "col": "メモ", "value": "済"}
    assert residue.find_unconsumed_words("所属が営業でない行のメモに「済」を付けて",
                                         args, []) == []


# --- 配線（実表を通した挙動）------------------------------------------------

def test_a_negated_value_does_not_become_a_second_condition(tmp_path):
    """① 否定した値が、他の列に在っても 2 組目にならないこと。"""
    book = _book(tmp_path, "neg", ["氏名", "所属", "担当", "メモ"],
                 [("A", "営業", "営業", ""), ("B", "総務", "田中", ""),
                  ("C", "経理", "営業", "")])
    ok, res, _inf, err = _set_where(book, "所属が営業でない行のメモに「済」を付けて",
                                    cmp_="neq")
    assert ok, err
    assert res.get("cond2_col") is None, res
    assert res["_match_rows"] == [3, 4], res["_match_rows"]


def test_an_incidental_name_does_not_become_a_second_condition(meibo):
    """④ 条件として書かれていない固有名詞で、対象行を縮めないこと。"""
    ok, res, _inf, err = _set_where(
        meibo, "田中さんのように、所属が営業の行のメモに「○」を付けて")
    assert ok, err
    assert res.get("cond2_col") is None, res
    assert res["_match_rows"] == [2, 4, 5], res["_match_rows"]


def test_a_candidate_that_was_not_used_is_disclosed(meibo):
    """★ 採らなかった候補は**黙らない** ── 直さないが ✓ も出さない。"""
    _ok, res, _inf, _err = _set_where(
        meibo, "田中さんのように、所属が営業の行のメモに「○」を付けて")
    said = " ".join(res.get("_warnings") or [])
    assert "田中" in said and "氏名" in said, said
    assert "条件の形で書かれていない" in said, said


def test_a_real_second_condition_still_works(meibo):
    """★ 対で縛る ── 本物の 2 条件は今までどおり通ること（塞ぎすぎの検分）。"""
    ok, res, _inf, err = _set_where(meibo, "所属が営業で担当が主任の行のメモに「○」を付けて")
    assert ok, err
    assert res.get("cond2_col") == "担当" and res.get("cond2_value") == "主任", res
    assert res["_match_rows"] == [2, 5], res["_match_rows"]


def test_the_postcondition_refuses_a_second_column_that_is_not_in_the_table(tmp_path):
    """③ 宣言した 2 組目の列が実表に無いなら、黙って 1 条件に落として pass しないこと。

    ★ col / cond_col には同じ検査が在るのに 2 組目にだけ無かった＝片配線。
      見出しの綴りが 1 文字違うだけで、**2 組目が丸ごと無視されたまま ✓** が出ていた。
    """
    from ailine_core.postconditions import shape

    def mk(path, marks):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "売上"
        ws.append(["品名", "部門", "金額", "備考"])
        for (n, b, a), m in zip([("a", "経理", 1200), ("b", "営業", 1500),
                                 ("c", "営業", 1100), ("d", "総務", 900)], marks):
            ws.append([n, b, a, m])
        wb.save(path)
        return path

    before = mk(tmp_path / "b.xlsx", ["", "", "", ""])
    half = mk(tmp_path / "h.xlsx", ["○", "○", "○", ""])   # 金額条件だけを適用した回
    args = {"col": "備考", "cond_col": "金額", "cmp": "gte", "cond_value": 1000,
            "value": "○", "cond2_col": "部署", "cond2_cmp": "eq", "cond2_value": "営業",
            "_target_sheet": "売上", "_header_row": 1}     # ★ 実表の見出しは『部門』
    st, msg = shape.check_set_where(half, args, header_row=1, source_book=before)
    assert st == "fail", (st, msg)
    assert "部署" in msg and "ありません" in msg, msg


@pytest.mark.parametrize("task", [
    "金額が1000以上で部門が経理か営業の行の備考に「○」を付けて",
    "金額が1000以上で経理か営業の行の備考に「○」を付けて",
])
def test_an_enumeration_is_refused_with_a_real_route(tmp_path, task):
    """② 1 列に複数の値（OR）を、黙って 1 つに縮めないこと。"""
    book = _book(tmp_path, "or", ["品名", "部門", "金額", "備考"],
                 [("a", "経理", 1200, ""), ("b", "営業", 1500, ""),
                  ("c", "営業", 1100, ""), ("d", "総務", 1200, "")], sheet="売上")
    ok, _res, _inf, err = ailine.verify_dsl_args(
        "SET_WHERE", {"col": "備考", "cond_col": "金額", "cmp": "gte", "value": "○"},
        ailine.build_book_meta(book), task=task, vocab=ailine.load_vocab())
    assert not ok, "列挙を黙って半分やっている"
    assert "1 つの値まで" in err, err
    # ★ 案内は実表から作る ── 挙げた値が両方とも、行数つきで出ること
    assert "部門が経理の行を抜き出して" in err and "部門が営業の行を抜き出して" in err, err
    assert "（1 行）" in err and "（2 行）" in err, err
