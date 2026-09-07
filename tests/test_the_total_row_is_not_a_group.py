# -*- coding: utf-8 -*-
"""合計行を「取引先」として集計しないこと ── その番人（2026-09-07）。

★★ 出所（外部の査定が 2/2 で再現した ×）: 合計行のある請求書で
  「取引先ごとに金額を集計して」が `× 集計に含まれないグループがある` で落ちていた。

★★ そして**直し方を一度間違えた**。最初、検算側で合計行を除外して緑にした ──
  出来上がりを独立に読んだら、集計シートはこうなっていた:

      ('合計', 356400)   ← 元の合計行を「合計という取引先」として集計している
      ('合計', 712800)   ← そのうえで総計を足したので **2 倍**

  ★ 元の × は正しかった。壊れていたのは**生成側**で、私は**検算の目を潰して**緑にした。
    この repo で一番やってはいけない直し方だった。

★ 正しい直しは 3 つ:
  ① 生成が合計行を除く（除外の宣言 `_skip_rows` を Basic へ渡す・条件つき書換と同じ配線）
  ② 除外の宣言を**入口で 1 度だけ**作る（旧: 4 つの op が書き写し、集計だけ書き忘れ）
  ③ 検算が**総計そのもの**と**総計の下に何も無いこと**を見る
     ── 旧版は最初の『合計』で走査を止めており、**その先を誰も見ていなかった**
        （だから ② の悪い直しが緑になった）
"""
from __future__ import annotations

import openpyxl
import pytest

from ailine_core.postconditions import derive


def _src_and_out(tmp_path, out_rows):
    """元の表（合計行つき）と、集計シートを手で組んだブックを作る。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "8月請求"
    ws.append(["取引先", "金額"])
    for r in [("丸和物流", 57600), ("ヤマノ食品", 42000),
              ("ヤマノ食品", 18000), ("北斗精機", 114000)]:
        ws.append(list(r))
    ws.append(["合計", 231600])          # ★ データ行ではない
    out = wb.create_sheet("集計")
    out.append(["取引先", "合計 - 金額"])
    for r in out_rows:
        out.append(list(r))
    p = tmp_path / "b.xlsx"
    wb.save(p)
    return p


#: 正しい出来上がり（合計行を除いた 3 グループ + 総計）
GOOD = [("丸和物流", 57600), ("ヤマノ食品", 60000), ("北斗精機", 114000),
        ("合計", 231600)]

ARGS = {"group_col": "取引先", "value_col": "金額",
        "_target_sheet": "8月請求", "_skip_rows": [6]}


def test_the_declaration_is_made_once_at_the_entrance(tmp_path):
    """② 除外の宣言は**入口で 1 度だけ**作られ、どの op にも届くこと。

    ★ 旧: 4 つの op がそれぞれ同じ式を書き写しており、**集計だけ書き忘れていた**
      （＝片配線）。ここが空なら、生成も検算も合計行を混ぜる。
    """
    import ailine
    book = _src_and_out(tmp_path, GOOD)
    ok, res, _inf, err = ailine.verify_dsl_args(
        "AGGREGATE", {"group_col": "取引先", "value_col": "金額"},
        ailine.build_book_meta(book), task="取引先ごとに金額を集計して",
        vocab=ailine.load_vocab())
    assert ok, err
    assert res.get("_skip_rows") == [6], res.get("_skip_rows")


def test_the_generator_passes_the_declaration_to_basic(tmp_path):
    """① 宣言を**生成側が Basic へ渡す**こと ── 検算だけ直しても意味が無い。

    ★ ここを叩く検体が無かったため、生成側の配線を殺しても試験は緑のままだった
      （2026-09-07 の変異試験で発覚）。**番人でなく試験が見ていなかった**方の抜け。
    """
    import ailine
    book = _src_and_out(tmp_path, GOOD)
    bm = ailine.build_book_meta(book)
    _ok, res, _inf, _err = ailine.verify_dsl_args(
        "AGGREGATE", {"group_col": "取引先", "value_col": "金額"},
        bm, task="取引先ごとに金額を集計して", vocab=ailine.load_vocab())
    code = ailine.codegen_dsl("AGGREGATE", res, book_meta=bm)
    call = [ln for ln in code.splitlines() if "SummaryTable" in ln][0]
    assert '"5"' in call, f"除外の行が Basic に渡っていない: {call}"


def test_the_total_row_is_not_a_group(tmp_path):
    st, msg = derive.check_aggregate(_src_and_out(tmp_path, GOOD), ARGS, header_row=1)
    assert st == "pass", (st, msg)
    assert "3 グループ" in msg, msg


def test_without_the_declaration_the_total_row_leaks_in(tmp_path):
    """★ 宣言が届かなければ落ちること ── 生成側を戻した時に赤くなる縛り。"""
    args = {k: v for k, v in ARGS.items() if k != "_skip_rows"}
    st, _msg = derive.check_aggregate(_src_and_out(tmp_path, GOOD), args, header_row=1)
    assert st == "fail", "合計行を混ぜても通ってしまう"


@pytest.mark.parametrize("rows, why", [
    ([("丸和物流", 57600), ("ヤマノ食品", 60000), ("北斗精機", 114000),
      ("合計", 463200)], "総計が 2 倍"),
    ([("丸和物流", 57600), ("ヤマノ食品", 60000), ("北斗精機", 114000),
      ("合計", 231600), ("合計", 463200)], "総計の下に余分な行"),
])
def test_the_total_itself_is_verified(tmp_path, rows, why):
    """★ 旧版は最初の『合計』で止まり、その先も総計そのものも見ていなかった。"""
    st, msg = derive.check_aggregate(_src_and_out(tmp_path, rows), ARGS, header_row=1)
    assert st == "fail", f"{why} を見逃した: {msg}"


def test_a_missing_total_row_is_a_warning_not_a_failure(tmp_path):
    """★ 主張の強さを事実に合わせる ── 総計が見当たらないのは「間違い」ではなく
    「**確かめきれていない**」。✓ は出さないが × でもない。

    ★ 実測（2026-09-07）: ここを × にしたら、既存シートへ書こうとして**上書きの確認で
      止まる**検体が、確認の案内ごと消えて exit 7 → 1 に変わった。弱めて逃げたのではなく、
      ⚠ が事実に合っている（`warn` は ✓ を出さない）。
    """
    rows = [("丸和物流", 57600), ("ヤマノ食品", 60000), ("北斗精機", 114000)]
    st, msg = derive.check_aggregate(_src_and_out(tmp_path, rows), ARGS, header_row=1)
    assert st == "warn", (st, msg)
    assert "総計の行が" in msg, msg
