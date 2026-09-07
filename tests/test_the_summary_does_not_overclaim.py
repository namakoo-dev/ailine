# -*- coding: utf-8 -*-
"""要約文が、検算より広いことを主張しないこと（2026-09-07）。

★★ 出所（外部の UX 検品）: 列を削除した回の画面がこうなっていた ──

    C2: 値 5→1000
    D2: 値 1000→'=C2-B2'          ← 20 セル近い変化が並ぶ
    …
    事後条件を確認（操作:列削除）: 列『在庫』を削除（**残りの列は 1 セルも変わらず**）

  **直上の差分と、直後の要約文が矛盾している。** 操作自体は正しい（右の列が左へ詰まり、
  式も書き直される）。嘘だったのは**文**の方で、検算が証明しているのは
  「残った列の中身が、その 1 列を抜いた並びと一致する」ことだけだった。

★ 同じ文が**列の挿入**にも在った（途中に挿せば右の列は動く）。片方だけ直さない。
★ 直し方は「文を検算に合わせる」── 検算を文に合わせない。
"""
from __future__ import annotations

import openpyxl
import pytest

from ailine_core.postconditions import move


def _book(path, headers, rows, formulas=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "在庫表"
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    for (cell, f) in (formulas or []):
        ws[cell] = f
    wb.save(path)
    return path


@pytest.fixture
def before(tmp_path):
    return _book(tmp_path / "b.xlsx", ["品名", "単価", "在庫", "売上", "利益"],
                 [("ボルト", 100, 5, 1000), ("ナット", 50, 8, 800)],
                 [("E2", "=D2-B2"), ("E3", "=D3-B3")])


def test_deleting_a_column_does_not_claim_nothing_moved(before, tmp_path):
    """★ 列を消せば右の列は左へ詰まる ── 「1 セルも変わらず」は事実でない。"""
    after = _book(tmp_path / "a.xlsx", ["品名", "単価", "売上", "利益"],
                  [("ボルト", 100, 1000), ("ナット", 50, 800)],
                  [("D2", "=C2-B2"), ("D3", "=C3-B3")])
    st, msg = move.check_delete_column(
        after, {"col": "在庫", "_target_sheet": "在庫表"},
        header_row=1, source_book=before)
    assert st == "pass", (st, msg)
    assert "1 セルも変わらず" not in msg, msg
    assert "詰まり" in msg or "ずれ" in msg, msg
    assert "中身" in msg, msg


def test_inserting_a_column_does_not_claim_nothing_moved(before, tmp_path):
    """★ 対で縛る ── 挿入側にも同じ文が在った（片方だけ直すのが事故の形）。"""
    after = _book(tmp_path / "i.xlsx", ["品名", "単価", "備考", "在庫", "売上", "利益"],
                  [("ボルト", 100, None, 5, 1000), ("ナット", 50, None, 8, 800)],
                  [("F2", "=E2-B2"), ("F3", "=E3-B3")])
    st, msg = move.check_add_column(
        after, {"name": "備考", "_at_col": 3, "_target_sheet": "在庫表"},
        header_row=1, source_book=before)
    assert st == "pass", (st, msg)
    assert "1 セルも変わらず" not in msg, msg
