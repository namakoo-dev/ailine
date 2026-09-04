"""bench の検算（測定器）が、正しい出力を通し**壊れた出力を落とす**ことを確かめる。

★ なぜ要るか（2026-09-04）: 段2 の検算 6 本のうち **2 本が最初は間違っていた**。

  ① 合計行 ── 製品は `=SUM(...)` の**式**を書くのに「数値でなければ失敗」と書いて
     いた。正しい出力 6 件を全部 × と読んだ。
  ② 空行 ── 空セルは None で来るのに `str(x).strip()` で見ていたので
     `"None"` が真になり、空行を「空でない」と読んだ。正しい出力 8 件を全部 × と読んだ。

  どちらも**製品ではなく測定器の側の誤り**で、しかも「厳しすぎる」向きに壊れていた
  ── つまり見逃しではなく**正しいものを不合格と主張する**壊れ方。
  ★ ②は最初の変異試験を通していた。`""` は試したが **None を試していなかった**からで、
  「試験が在っても、その事故の形では鳴らない」の実例。

★ ここは LLM も LibreOffice も使わない（格子の計算だけ）ので一瞬で終わる。
"""
import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


def _bench():
    spec = importlib.util.spec_from_file_location(
        "basic_ops_matrix", ROOT / "bench" / "basic_ops_matrix.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


B = [["品名", "棚", "数量", "備考"],
     ["ボルト", "A-1", 120, None],
     ["ナット", "A-2", 80, None],
     ["ワッシャー", "B-1", 300, None]]


def _cp(grid):
    return [list(r) for r in grid]


def _run(check, before, after):
    return check(before, after)[0]


# --- ① 合計行 --------------------------------------------------------------
def test_total_row_check_accepts_both_a_number_and_a_live_formula():
    m = _bench()
    chk = m.total_row_appended(3)
    num = _cp(B) + [["合計", None, 500, None]]
    formula = _cp(B) + [["合計", None, "=SUM(C2:INDEX(C:C,ROW()-1))", None]]
    assert _run(chk, B, num), "数値の合計を落としてはいけない"
    assert _run(chk, B, formula), "★ 式の合計を落としてはいけない（製品は式を書く）"


@pytest.mark.parametrize("label, after", [
    ("合計が空", _cp(B) + [["合計", None, None, None]]),
    ("合計の値が違う", _cp(B) + [["合計", None, 499, None]]),
    ("式が別の列を指す", _cp(B) + [["合計", None, "=SUM(A2:A9)", None]]),
    ("式が SUM でない", _cp(B) + [["合計", None, "=AVERAGE(C2:C4)", None]]),
    ("行が増えていない", _cp(B)),
])
def test_total_row_check_rejects_broken_output(label, after):
    assert not _run(_bench().total_row_appended(3), B, after), label


def test_total_row_check_rejects_a_total_that_damaged_the_rows_above():
    after = _cp(B)
    after[1][2] = 999
    after.append(["合計", None, 1379, None])
    assert not _run(_bench().total_row_appended(3), B, after)


# --- ② 空行 ----------------------------------------------------------------
@pytest.mark.parametrize("filler", [None, "", "   "])
def test_blank_row_check_accepts_every_shape_of_empty(filler):
    """★ ここが最初に抜けた所 ── None を試していなかった。"""
    after = _cp(B)
    after.insert(2, [filler] * 4)
    assert _run(_bench().blank_rows_inserted(3), B, after), repr(filler)


@pytest.mark.parametrize("label, mutate", [
    ("挿さった行に値がある", lambda a: a.insert(2, ["新品", None, None, None])),
    ("位置が違う", lambda a: a.insert(3, [None] * 4)),
    ("そもそも増えていない", lambda a: None),
])
def test_blank_row_check_rejects_broken_output(label, mutate):
    after = _cp(B)
    mutate(after)
    assert not _run(_bench().blank_rows_inserted(3), B, after), label


def test_blank_row_check_rejects_an_insert_that_damaged_the_other_rows():
    after = _cp(B)
    after.insert(2, [None] * 4)
    after[3][2] = 1
    assert not _run(_bench().blank_rows_inserted(3), B, after)


# --- ③ 列を一律の値で埋める ------------------------------------------------
def test_column_set_check_watches_the_columns_it_was_not_asked_to_touch():
    m = _bench()
    chk = m.column_all_became(4, "確認済")
    ok = _cp(B)
    for r in ok[1:]:
        r[3] = "確認済"
    assert _run(chk, B, ok)
    broke = _cp(ok)
    broke[1][1] = "X"
    assert not _run(chk, B, broke), "★ 他の列を壊しても通ってはいけない"
    half = _cp(B)
    half[1][3] = "確認済"
    assert not _run(chk, B, half), "1 行だけでは通してはいけない"


# --- ④ 条件つき書換 --------------------------------------------------------
def test_conditional_check_rejects_both_too_wide_and_too_narrow():
    m = _bench()
    chk = m.rows_matching_changed(3, 100, 4, "○")
    ok = _cp(B)
    ok[1][3] = "○"
    ok[3][3] = "○"
    assert _run(chk, B, ok)
    wide = _cp(B)
    for r in wide[1:]:
        r[3] = "○"
    assert not _run(chk, B, wide), "★ 列全体を書き換えても通ってはいけない（本命の事故）"
    narrow = _cp(B)
    narrow[1][3] = "○"
    assert not _run(chk, B, narrow), "条件に合う行の片方だけでは通してはいけない"
    assert not _run(chk, B, _cp(B)), "何もしなくても通ってはいけない"


# --- ⑤ 重複除去（新しいシートを読む）---------------------------------------
D = [["取引先", "品名", "数量", "担当"],
     ["あかね商事", "机", 7, "佐藤"],
     ["うえだ物産", "椅子", 4, "鈴木"],
     ["あかね商事", "棚", 9, "佐藤"],
     ["うえだ物産", "椅子", 4, "鈴木"]]


class _FakeBooks:
    def __init__(self, grid, why="重複除去"):
        self._grid, self._why = grid, why

    def new_sheet_grid(self):
        if self._grid is None:
            return None, "新しいシートが 0 枚（1 枚のはず）: []"
        return self._grid, self._why


def test_dedup_check_reads_the_new_sheet_and_guards_the_original():
    m = _bench()
    chk = m.rows_deduped(2)
    good = D[:4]
    assert chk(D, D, _FakeBooks(good))[0]
    assert not chk(D, D, _FakeBooks(None))[0], "新しいシートが無ければ落とす"
    assert not chk(D, D, _FakeBooks(_cp(D)))[0], "減っていなければ落とす"
    assert not chk(D, D, _FakeBooks([D[0], D[2], D[4]]))[0], "重複が残っていれば落とす"
    assert not chk(D, D, _FakeBooks([D[0], D[1], ["ねつ造", "机", 7, "佐藤"]]))[0], \
        "元に無い行が現れたら落とす"
    changed = _cp(D)
    changed[1][0] = "書き換えた"
    assert not chk(D, changed, _FakeBooks(good))[0], \
        "★ 元のシートが変わっていたら落とす（新しいシートを作るのが契約）"


# --- ⑥ セル分割 ------------------------------------------------------------
S = [["氏名", "連絡先", "部署"],
     ["川口", "03-1、090-2", "営業"],
     ["森田", "03-5", "経理"]]


def test_split_check_wants_the_parts_and_the_original_column_together():
    m = _bench()
    chk = m.cell_split_into(2, 2, "、")
    ok = [S[0] + ["連絡先_1", "連絡先_2"],
          ["川口", "03-1、090-2", "営業", "03-1", "090-2"],
          ["森田", "03-5", "経理", "03-5", None]]
    assert _run(chk, S, ok)
    assert not _run(chk, S, _cp(S)), "何もしていなければ落とす"
    lost = [S[0] + ["連絡先_1", "連絡先_2"],
            ["川口", "03-1", "営業", "03-1", "090-2"],
            ["森田", "03-5", "経理", "03-5", None]]
    assert not _run(chk, S, lost), "★ 元の列が消えていたら落とす（SPLIT_CELL の契約）"
    partial = [S[0] + ["連絡先_1"],
               ["川口", "03-1、090-2", "営業", "03-1"],
               ["森田", "03-5", "経理", "03-5"]]
    assert not _run(chk, S, partial), "片方しか出ていなければ落とす"
