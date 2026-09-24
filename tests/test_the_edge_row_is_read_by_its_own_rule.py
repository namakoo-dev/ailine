"""表の端の 1 行（「最終行の担当」「一番上の商品」）を読む規則を、関数の docstring の例のとおりに縛る（2026-09-24）。

★ なぜ在るか: 位置の兄弟を ailine_core/anchor.py へ移した時、変異（この関数を「何も見つけない」にする）が
  **関係しそうな試験 943 本を全部すり抜けた** ── この関数を守っていたのは「呼び出しが 3 か所」を数える
  試験だけで、挙動を見る検体が 1 本も無かった（移す前から）。docstring に採る／採らないの例が書いてあるので、
  それをそのまま検体にする。
"""
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from ailine_core.anchor import task_names_a_table_edge_row  # noqa: E402

HEADS = ["商品", "担当", "金額"]


def _book(tmp_path, rows):
    p = tmp_path / "表.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"
    ws.append(HEADS)
    for r in rows:
        ws.append(r)
    wb.save(p)
    return {"headers": {"Sheet": HEADS}, "header_rows": {"Sheet": 1}, "path": str(p), "sheets": ["Sheet"]}


@pytest.fixture
def meta(tmp_path):
    return _book(tmp_path, [["りんご", "佐藤", 100], ["みかん", "鈴木", 200], ["ぶどう", "田中", 300]])


def test_the_last_row_named_with_a_real_column_is_taken(meta):
    got = task_names_a_table_edge_row("最終行の担当を「佐藤」にして", meta, "Sheet")
    assert got and got[0] == 4, got            # 見出し 1 行 + データ 3 行 → 最後の既存行は 4 行目
    assert "最終行" in got[1]


def test_the_top_row_is_the_one_after_the_header(meta):
    got = task_names_a_table_edge_row("一番上の商品を太字にして", meta, "Sheet")
    assert got and got[0] == 2, got


@pytest.mark.parametrize("task", [
    "一番下に行を足して",            # 「に」なので行を足す側の話（採らない）
    "税込み合計を一番下に出して",    # 同上
    "商品の列を末尾に移動して",      # 列の依頼
    "最後の列を「済」にして",        # 「列」は実表の列名ではない
])
def test_what_the_docstring_says_it_does_not_take(meta, task):
    assert task_names_a_table_edge_row(task, meta, "Sheet") is None


def test_an_empty_table_has_no_edge(tmp_path):
    empty = _book(tmp_path, [])
    assert task_names_a_table_edge_row("最終行の担当を「佐藤」にして", empty, "Sheet") is None
