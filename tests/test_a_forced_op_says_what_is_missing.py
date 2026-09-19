# -*- coding: utf-8 -*-
"""`--op` を固定して読み取れなかった回は、**何が足りないか**を画面で言う（2026-09-20）。

★★ 出所（導通の盤が `vague` と判定した 3 件の 1 つ）: 旧版はこうだった ──

    ？ 『1セル書換』として読み取れませんでした
      （依頼文に、対象の列や値が書かれているか確かめてください）

  何が足りないのかを**名指ししていない**ので、人は自分の書き方の何を直せばいいか
  分からない。合格線の 3 条目（通る道を示す）を満たしていない形。

★★ 機械は言える材料を持っていた:
  ・`OP_SCHEMA[op]` ── その操作に要る項目
  ・`_CONFIRM_FIELDS[op]` ── それを**人が読む名前**にする表（解釈行と同じ語）
  ・`ailine_core/examples.py` ── そのまま打てば通る例（文面はそこ 1 箇所）
  ★ 今日の他の直しと同じ形 ── 器官は在って、この断りだけが使っていなかった。

★★ この試験は**製品を実際に走らせて画面を読む**（2026-09-20・初版の直し）:
  初版は `_CONFIRM_FIELDS` と `OP_SCHEMA` から期待値を組み立て直して突き合わせていた。
  つまり**表を読んで表と比べていた** ── 断りの文面を鍵むき出しに戻す変異を 2 つとも
  緑で通した（変異試験が指した）。番人は製品の出口に立たせる。
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


def _refuse(op: str, task: str = "なんとかして") -> str:
    """`--op <op>` を固定したが読み取れなかった回の**画面**を返す。

    ★ 翻訳（LLM）は落ちた体で固定する ── 素の環境でも毎回走らせるため。
      引き金は「読み取れなかった」であって、どう読み取れなかったかではない。
    """
    real = ailine.translate_task_fixed_op
    ailine.translate_task_fixed_op = lambda *a, **k: None
    buf = io.StringIO()
    try:
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "在庫.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["商品", "金額"])
            ws.append(["ボルト", 120])
            wb.save(p)
            wb.close()
            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                try:
                    rc = ailine.main(["run", str(p), task, "--copy", "--op", op])
                except SystemExit as e:
                    rc = e.code
    finally:
        ailine.translate_task_fixed_op = real
    assert rc != 0, f"★ 引き金が引けていない（断りが出ずに exit {rc}）"
    return buf.getvalue()


@pytest.mark.parametrize("op, want", [
    ("SET_CELL_VALUE", ["対象の行", "対象列"]),
    ("EXTRACT", ["対象列", "条件", "値"]),
    ("SORT", ["対象", "順"]),
    ("AGGREGATE", ["分類列", "集計列"]),
])
def test_the_screen_names_the_missing_items_in_human_words(op, want):
    """★★ 事故そのもの: 足りない項目が**人の語で**画面に出ること。

    ★ 期待値はここに手で書く ── 製品の表から組み立てると、表を読んで表と比べる形に戻る。
    """
    out = _refuse(op)
    for w in want:
        assert w in out, f"★ 『{w}』が画面に無い:\n{out}"


@pytest.mark.parametrize("op, key", [
    ("SET_CELL_VALUE", "row"),
    ("EXTRACT", "cmp"),
    ("AGGREGATE", "group_by"),
])
def test_the_screen_never_shows_a_raw_schema_key(op, key):
    """★★ スキーマの鍵をそのまま見せないこと（`cmp` `group_by` は人に通じない）。"""
    out = _refuse(op)
    assert key not in out, f"★ 内部の鍵『{key}』が画面に出ている:\n{out}"


def test_an_op_without_display_names_says_nothing_rather_than_leaking():
    """★★ 表示名を持たない鍵は**出さない**（実測で該当は ADD_ROW の values だけ）。

    ★ 言えない回は項目を並べない ── 内部の語を見せるより黙る方がまし。
    ★ ここが緩むと `values` のような鍵が買い手の画面に出る。
    """
    out = _refuse("ADD_ROW")
    assert "values" not in out, f"★ 内部の鍵が漏れている:\n{out}"
    assert "が要ります" not in out, f"★ 全部言えないのに項目を並べている:\n{out}"


def test_the_screen_shows_a_way_that_works():
    """★ 通る例を見せること（合格線の 3 条目）── 文面は `examples.py` の 1 箇所が正。

    ★ 「例」という語を探さない ── 文面は最終調整で変わる（番人を字面で書かない）。
      不変なのは「そのまま打てる依頼文が画面に在る」こと。取り出しは導通の盤と同じ目。
    ★ その文が**実際に通るか**は導通の盤が歩く（`refusal_register.json` の walk 欄）。
      ここは LLM を回さないので、取り出せることまでを縛る。
    """
    import walk_refusals_core as walk
    out = _refuse("EXTRACT")
    assert walk._example_in(out), f"★ そのまま打てる依頼文が画面に無い:\n{out}"
    from _product_source import window_around
    body = window_around("として読み取れませんでした", before=700, after=400)
    assert "render_example_line(" in body, "★ 例の文面を書き写している疑い"


def test_most_ops_can_name_their_items():
    """★ 項目を名指しできる op が減っていないこと（実測 30 中 26）。

    ★ 減ったら `_CONFIRM_FIELDS` と `OP_SCHEMA` がずれた合図 ── 黙る op が増えて
      「何が足りないか分からない」断りに戻る。画面を数えるので製品の出口で測る。
    """
    can = [op for op in ailine.OP_SCHEMA if "が要ります" in _refuse(op)]
    assert len(can) >= 26, f"★ 項目を名指しできる op が {len(can)} に減った（26 以上のはず）"
