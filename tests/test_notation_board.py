# -*- coding: utf-8 -*-
"""依頼の項にも分母を持つ（2026-09-18・Namakoo「三項を満たしているかをチェックする機構はあるか？」）。

★★ 調べて分かったこと: 三項（依頼・宣言・実体）のうち **宣言**の項だけに分母があった。
  plan_writes_beyond_one_cell は OP_WRITE_TARGET から導かれ、
  test_every_value_writing_op_is_covered_by_the_gate が漏れを赤にする。
  ところが**依頼**の項には名簿も表も無く、`三項` という語は repo 中に 40 箇所あるのに
  **全部 docstring とコメント** ── 人が書いた物語で、機械が確かめている物は 1 つも無かった。

★★ なぜそれが致命か: task_points_at_one_row が None を返した時、
    (a) 依頼は本当に 1 か所を指していない  → 一括書換でよい
    (b) 指しているが、読めない記法だった    → 一括書換は事故
  この 2 つが**画面上で区別できない**。2026-09-18 の盲検 4 体目 ④ は (b) で、
  「セル E1 に 借方金額 と入力して」が D2〜D7 を潰して exit 0 になった。
  ★ 棚の線「出ないことは信号でない」── 致命の大半がこの形。

★ ここが縛るもの: 名簿（tests/notation_register.json）の宣言と、実物の器官の実測が
  1 検体ずつ一致すること。ずれたら、それが**直し**でも**退行**でも赤くなる。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tests"))

import notation_board_core as board  # noqa: E402


def test_the_register_and_the_product_agree():
    """★★ 事故そのもの: 名簿が「読める」と言う記法を、実物が読めていること。

    ★ 逆向きも縛る ── 「読まない」と書いた検体（品番 ABC123・2 か所 B2 と C3）を
      実物が読み始めたら、それは**拾いすぎ**でやはり赤。
    """
    bad = board.mismatches(board.survey())
    assert not bad, "名簿と実物がずれている:\n" + "\n".join(
        f"  {b['key']}  {b['kind']}  {b['task']}  宣言={b['declared']} 実測={b['measured']}"
        for b in bad)


def test_every_notation_carries_a_specimen():
    """★ 分母は検体で持つ ── 文だけ書いて検体の無い記法を名簿に置かない。"""
    reg = board.load_register()
    for key, rec in reg["notations"].items():
        assert rec["specimens"], f"{key} に検体が無い"
        assert rec["label"], f"{key} に人が読む名前が無い"
        assert rec["points_at"] in board.UNREAD_STATUS, f"{key} の points_at: {rec['points_at']}"


def test_the_readable_notations_have_a_negative_control():
    """★★ 恒真殺し: 「読める」と宣言した記法には、**読まない**検体を必ず持たせる。

    ★ 陽性だけ並べた名簿は「全部 YES を返す」実装でも緑になる ──
      2026-09-18 に A1 の境界（品番 ABC123・範囲 A1:C5）で実際に必要だった。
    """
    reg = board.load_register()
    have = [k for k, r in reg["notations"].items()
            if r["status"] == "reads" and any(not s["reads"] for s in r["specimens"])]
    assert "a1_cell" in have, (
        "★ A1 形式に陰性対照が無い（品番を座標と読む退行が緑で通る）")


def test_the_stock_of_unreadable_notations_is_named_not_hidden():
    """★ 読めない記法は**在庫**として名指しで残す（Namakoo:「断りは仕方なく断るにすぎない」）。

    ★ 在庫に unlock（何が変われば到達になるか）が無いと、未来の俺たちが仕様と読む
      ── 棚の線「断定で刻むと未来が拾えない」。
    """
    reg = board.load_register()
    stock = {k: r for k, r in reg["notations"].items()
             if r["status"] in ("gap", "narrower_than_column")}
    assert stock, "★ 在庫ゼロ ── 本当に全部読めるなら在庫を消してよいが、まず盤で確かめる"
    for key, rec in stock.items():
        assert rec.get("unlock"), f"{key} に unlock が無い（在庫が仕様に化ける）"


def test_the_status_is_derived_not_copied():
    """★★ 恒真殺し: status は名簿を写さず、points_at と**実測**から導かれること。

    ★ 初版の _derive_status は最後の分岐で名簿の status 自身を読んでいた
      ── 名簿を写して名簿と比べる形（この repo が何度も踏んだ恒真）。
    """
    import inspect
    body = inspect.getsource(board._derive_status)
    assert 'n["status"]' not in body, (
        "★ 導出が名簿の status を読んでいる（恒真に戻っている）")
    # ★ 実測を使っていること ── points_at だけで決めると器官を見ていない。
    assert 'reads_any' in body, "★ 導出が実測を見ていない"


def test_the_specimen_book_keeps_a_duplicate_name():
    """★ 冊の側の不変: 同名 2 行を必ず持つ。

    ★ 「指してはいるが決められない」を測れるのはこの形だけで、ここが落ちると
      同名 2 行あるだけで列が潰れる退行が緑で通る（実測した形）。
    """
    names = [n for n, _ in board.SPECIMEN_ROWS]
    assert len(names) != len(set(names)), "★ 検体の冊から同名の行が消えている"
