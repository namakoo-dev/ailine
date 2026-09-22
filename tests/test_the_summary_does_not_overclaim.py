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


@pytest.mark.parametrize("which", ["add_row", "delete_rows"])
def test_row_operations_do_not_claim_nothing_moved(before, tmp_path, which):
    """★ 行でも同じ ── 挿せば下がり、消せば詰まる。「元のまま」はずれを隠している。

    ★ 正直な書き方は同じ repo に在った（check_insert_rows の「シフトを確認」）。
      片方だけ直さず、行と列の 4 箇所すべてを同じ作法へ揃えた。
    """
    if which == "add_row":
        after = _book(tmp_path / "ar.xlsx", ["品名", "単価", "在庫", "売上", "利益"],
                      [("ボルト", 100, 5, 1000), ("新品", None, None, None),
                       ("ナット", 50, 8, 800)],
                      [("E2", "=D2-B2"), ("E4", "=D4-B4")])
        st, msg = move.check_add_row(
            after, {"at": 3, "_target_sheet": "在庫表", "values": {"品名": "新品"}},
            header_row=1, source_book=before)
    else:
        after = _book(tmp_path / "dr.xlsx", ["品名", "単価", "在庫", "売上", "利益"],
                      [("ナット", 50, 8, 800)], [("E2", "=D2-B2")])
        st, msg = move.check_delete_rows(
            after, {"at": 2, "count": 1, "_target_sheet": "在庫表"},
            header_row=1, source_book=before)
    assert st in ("pass", "warn"), (st, msg)
    if st == "pass":
        assert "元のまま" not in msg, msg
        assert "ずれ" in msg or "詰まり" in msg, msg


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


# --- 事後条件の画面は 1 か所から出す（2026-09-22・盲検 3 体目 ⑧ / 4 体目 ⑦ の前段）------

def test_the_postcondition_screen_is_printed_from_one_place():
    """★★ 事後条件の結果を画面に出す 17 行は、**3 か所に写されていた**。

    ★★ しかもその中のコメントが「**1 実装・全経路**」と書いていた ── 宣言と実体の
      食い違いで、片方だけ直す事故の予約だった（この repo が何度も踏んだ形）。
    ★ 画面の語を直す（開発者語の言い換え）**前に**畳んだ。順番を逆にすると、
      直した人が「片方だけ直した人」になる。
    ★ 数え方: 画面に出す文が製品コードに 1 回だけ現れること
      ── `report_postcondition` の中だけ。
    """
    from _product_source import count_in_code
    for line in ("適用されたが事後条件を満たさない",
                 "事後条件を機械検証できなかった（操作",
                 "事後条件を確認（操作"):
        assert count_in_code(line) == 1, (
            f"『{line}』が製品コードに {count_in_code(line)} 回ある ── "
            "画面に出す所は report_postcondition 1 か所に畳むこと")


def test_the_folded_reporter_still_decides_the_exit_code():
    """★ 畳んだ先が**止める判断まで持っている**こと（呼び出し側に散らさない）。

    ★ 続行なら None・止めるなら終了コード、という 1 つの形にしてある。
    """
    from _product_source import window_around
    seg = window_around("def report_postcondition", before=0, after=2000)
    assert "return 1" in seg, "止める判断が畳んだ先に無い"
    assert "return None" in seg, "続行の合図が無い"
    assert 'result["ok"] = True' in seg, "成功の印を呼び出し側に残している"
