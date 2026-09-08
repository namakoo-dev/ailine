# -*- coding: utf-8 -*-
"""「依頼文にこの語が在る時だけ使う」を、散文でなく機械が縛る（2026-09-08）。

★★ 出所（Namakoo「使い方の広い操作に出る不具合が怖い」から辿った）:

    依頼   「分類ごとの売上を出して」        ← 『ピボット』とは一言も言っていない
    実行   操作:**ピボット**（実 CLI で 4/4）
    出力   **✓ 機械検証済み**（数字は正しいが、出るのは『ピボット』シート）

  ★ 禁止は**もう書いてあった** ── プロンプトに 5 行 + few-shot 1 例。守られていない。
    ★ 指示は意図、保証は機械。散文を宣言（requires_word）に移して機械が縛る。

★ これは「狭い op が広い op を奪う」形（同じ日の 3 件目）。集計は「〜ごとに」
  「まとめて」「別に」が流れ込む**広い口**なので、奪われると被害が頻度に比例する。

★ 同じ契約が 3 つある。実測で壊れていたのはピボットだけだが、**今は鳴らなくても縛る**
  （鳴る条件が将来できたときに黙って通るのを防ぐ）。
"""
from __future__ import annotations

import pytest

import ailine
from ailine_core import required_word


def _rules():
    return {op: (m["requires_word"], m.get("without_the_word"),
                 ailine.OP_LABELS.get(m.get("without_the_word"), ""))
            for op, m in ailine.OP_META.items() if m.get("requires_word")}


def test_the_permission_list_errs_wide_not_narrow():
    """★★ この表は**許可**なので、非対称である（2026-09-08 に実測して分かった）:

        広すぎる → モデルがその op を選んだ回だけ通る（元の分類精度に戻るだけ）
        狭すぎる → **正当な依頼を黙って別の op に変える**

    ★ だから広い側に倒す。ただし照合語(match_phrases)は流用しない ── PIVOT のそれには
      『入れ替え』が入っており、許可に使うと行の入れ替えが全部ピボットになりうる。
    """
    words = ailine.OP_META["PIVOT"]["requires_word"]
    # ★ op 自身が同義語として持つ言い方は、許可でも通ること
    for syn in ailine.OP_META["PIVOT"]["synonyms"]:
        assert any(w in syn for w in words), f"同義語『{syn}』が許可に通らない"
    # ★ 逆に、他の op と取り合いになる語は入れない
    assert "入れ替え" not in words and "集計" not in words, words


def test_the_contract_is_declared_not_written_in_prose():
    """★ 3 つの契約が宣言として在ること（プロンプトの散文に戻さない）。"""
    declared = {op for op, m in ailine.OP_META.items() if m.get("requires_word")}
    # ★★ 宣言するのは**読み替え先を持つ op だけ**（2026-09-08 に測って絞った）。
    #   DRAW_BORDERS/AUTOFIT にも同じ散文が在るが、依頼に無い段が湧く形は
    #   「捏造段」の機械が既に名指しして ✓ を降ろしている（test_operator9_fixes）。
    #   重ねた版は、複合計画の**正当な段まで巻き添えに断って**いた（番人が捕まえた）。
    assert declared == {"PIVOT"}, declared
    # ★ 読み替え先は「引数の形が同じ op」でなければならない（ピボット→集計）。
    assert ailine.OP_SCHEMA["PIVOT"] == ailine.OP_SCHEMA["AGGREGATE"]


@pytest.mark.parametrize("task, op, want", [
    # ★ 実測した事故そのもの ── 語が無いので集計へ読み替える
    ("分類ごとの売上を出して", "PIVOT", "AGGREGATE"),
    ("担当者別の売上をまとめて", "PIVOT", "AGGREGATE"),
    # ★ 対の試験: 語が在る回はそのまま通す（狭い op を殺さない）
    ("ピボットにして", "PIVOT", "PIVOT"),
    ("ピボットテーブルで集計して", "PIVOT", "PIVOT"),
    # ★★ 2026-09-08（Namakoo「ピボットで出してと言わない場合はどうやって出すの？」）:
    #   最初の許可リストは ("ピボット",) だけで、**op 自身の同義語『クロス集計』を
    #   見ていなかった**。実測で「クロス表にして」がピボット 4/4 → 集計に化けた
    #   ── 正当な依頼を黙って別の op に変える形で、直そうとした事故そのものを作った。
    ("クロス集計して", "PIVOT", "PIVOT"),
    ("クロス表にして", "PIVOT", "PIVOT"),
    ("分類と担当でクロス集計して", "PIVOT", "PIVOT"),
    ("縦に分類、横に担当で売上を出して", "PIVOT", "PIVOT"),
    # ★ 宣言の無い op は、この関所では触らない（既存の番人の受け持ち）
    ("表を整えて", "DRAW_BORDERS", "DRAW_BORDERS"),
    ("列幅を調整して", "AUTOFIT", "AUTOFIT"),
])
def test_an_op_without_its_word_is_rerouted_or_refused(task, op, want):
    plan, lines, refuse = required_word.enforce(
        [{"op": op, "args": {"group_col": "分類", "value_col": "売上"}}], task, _rules())
    if want is None:
        assert refuse and lines, (plan, lines, refuse)
        return
    assert not refuse, lines
    assert plan[0]["op"] == want, plan
    if want != op:
        assert lines and "読み直しました" in lines[0], lines
        # ★ 引数は落とさない（読み替えであって作り直しではない）
        assert plan[0]["args"] == {"group_col": "分類", "value_col": "売上"}
    else:
        assert lines == [], lines


def test_the_prose_and_the_declaration_are_kept_on_purpose():
    """★★ 散文と宣言が**二重に**在ることを、意図として固定する（2026-09-08・測って戻した）。

    はじめ「機械が縛るなら散文は要らない」と考えてピボットの説明を 5 行 → 1 行に削った。
    実 CLI の A/B で **別の op が壊れた**:

        「1行目のA列とB列を結合して」  セル結合 4/4 → **(空) 4/4**
        犯人はピボットの **-4 行**（列移動の +1 行ではない・切り分け済み）

    ★ 「説明を厚くする方が高くつく」は言い過ぎで、**削っても壊れる**。
      プロンプトは一枚の絡まった塊で、どちらの向きに触っても分類全体が組み変わる。
    ★ だから散文は据え置く。**authority は宣言（requires_word）の方**で、
      散文を直しても挙動は変わらない ── この試験がその約束を固定する。
    """
    doc = ailine.OPS_DOC
    pivot = [l for l in doc.splitlines() if l.startswith("PIVOT:")]
    assert len(pivot) == 1, pivot
    # ★ 散文は残っている（減らすと他の op が壊れると実測した）
    assert "明示された時だけ" in pivot[0], pivot[0]
    # ★ ただし挙動を決めるのは宣言の方
    assert ailine.OP_META["PIVOT"].get("requires_word"), "宣言が消えている"


def test_a_forced_op_is_left_alone():
    """★ 人が op を固定した回は触らない ── 画面に出した読みと実行を食い違わせない。

    ★ ここは配線の側の約束なので、呼び出し側の条件式を文字で縛る（実機を起こさない）。
    """
    import inspect
    src = inspect.getsource(ailine._reread_plan_from_the_table) if hasattr(
        ailine, "_reread_plan_from_the_table") else ""
    if not src:
        import pathlib
        src = pathlib.Path(ailine.__file__).read_text(encoding="utf-8")
    i = src.index("required_word.enforce")
    head = src[max(0, i - 400):i]
    assert "_forced_op" in head, head[-200:]


# --- 実機（本物の LLM を通して、画面と出来上がりまで見る）----------------------

@pytest.mark.local
def test_a_grouping_request_reaches_the_summary_not_the_pivot(tmp_path):
    """★ 実機 ── 「分類ごとの売上を出して」が『集計』シートを作ること。

    ★ 実測ではここが 4/4 で『ピボット』になり、✓ が出ていた（数字は正しいが
      依頼と違う形の表で、しかも開き直すたび書式が消える方）。
    """
    import os, subprocess, sys
    from pathlib import Path
    import openpyxl

    repo = Path(ailine.__file__).resolve().parents[2]
    book = tmp_path / "uriage.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["商品", "分類", "売上"])
    for r in (["りんご", "果物", 1200], ["みかん", "果物", 800],
              ["にんじん", "野菜", 2100]):
        ws.append(r)
    wb.save(book)

    r = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(book), "分類ごとの売上を出して",
         "--copy", "--sheet", "売上", "--timeout", "300"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900,
        env={**os.environ, "PYTHONPATH": str(repo / "src")})
    assert r.returncode == 0, r.stdout[-700:]
    out = openpyxl.load_workbook(book.with_name(book.stem + ".out.xlsx"))
    assert "集計" in out.sheetnames, (out.sheetnames, r.stdout[-700:])
    assert "ピボット" not in out.sheetnames, (out.sheetnames, r.stdout[-700:])
    rows = [tuple(x) for x in out["集計"].iter_rows(values_only=True)]
    assert ("果物", 2000) in rows and ("野菜", 2100) in rows, rows


@pytest.mark.local
def test_the_word_still_reaches_the_pivot(tmp_path):
    """★ 対の試験 ── 『ピボット』と言った回は今までどおりピボットが出る。"""
    import os, subprocess, sys
    from pathlib import Path
    import openpyxl

    repo = Path(ailine.__file__).resolve().parents[2]
    book = tmp_path / "uriage2.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["商品", "分類", "売上"])
    for r in (["りんご", "果物", 1200], ["にんじん", "野菜", 2100]):
        ws.append(r)
    wb.save(book)

    r = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(book),
         "分類ごとの売上をピボットで出して", "--copy", "--sheet", "売上", "--timeout", "300"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900,
        env={**os.environ, "PYTHONPATH": str(repo / "src")})
    assert r.returncode == 0, r.stdout[-700:]
    out = openpyxl.load_workbook(book.with_name(book.stem + ".out.xlsx"))
    assert "ピボット" in out.sheetnames, (out.sheetnames, r.stdout[-700:])
