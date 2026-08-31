# セル 2 つの入れ替え ── 2026-08-31。
# Namakoo「これは間違った操作だが機械検証が通っている。やりたいのは丸和物流の単価セルの値と
# みどり建設の単価セルの値を入れ替えたいだけだが、それを行の入れ替えと判断している」
#
# ★★ 実測（画面）: 「丸和物流の単価とみどり建設の単価を入れ替えて」で
#   **行を丸ごと入れ替えて ✓ を出していた**（16 セルが動いた・頼んだのは 2 セル）。
#   ★ ✓ が出た理由: 番人は「**宣言どおり**行が入れ替わったか」を見る。
#     宣言そのものが違えば、検算は通る ── 三項（依頼・宣言・実体）の「依頼」が抜けた形。
#     この repo で何度も踏んでいるが、**✓ が出たまま気づけない**のはこれが一番危ない。
#
# ★ 依頼文には証拠が在った: **両側とも「〜の〈列名〉」と列を名指ししている**。
#   行の入れ替えなら列は出てこない。
# ★ 言い回しを数え上げない ── 見るのは**実表に在る見出し**と**実表に在る行**だけ。
#
# ★ 直しは 2 段階で入れた（安全側を先に）:
#   ① 検出して**断る**（✓ を出す道を先に塞ぐ）→ commit
#   ② セルの入れ替えとして**実装する**（SWAP の第 3 の軸。新しい op は作らない）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

HEADERS = ["取引先", "項目", "件数", "単価", "金額"]
ROWS = [["丸和物流", "配送", 12, 4800, 57600],
        ["近江スチール", "鋼材", 5, 12000, 60000],
        ["みどり建設", "内装", 9, 7200, 64800]]


@pytest.fixture()
def meta(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    return {"sheets": ["請求"], "headers": {"請求": list(HEADERS)},
            "header_rows": {"請求": 1}, "path": str(p)}


# --- ① 検出（誤爆させない）----------------------------------------------------------------

def test_two_named_cells_are_detected(meta):
    got = ailine.swap_targets_are_cells("丸和物流の単価とみどり建設の単価を入れ替えて",
                                          meta, "請求")
    assert got == [(2, 4), (4, 4)], got


def test_the_two_columns_may_differ(meta):
    got = ailine.swap_targets_are_cells("丸和物流の単価とみどり建設の金額を入れ替えて",
                                          meta, "請求")
    assert got == [(2, 4), (4, 5)], got


@pytest.mark.parametrize("task", [
    "丸和物流とみどり建設を入れ替えて",          # 行の入れ替え
    "金額と単価を入れ替えて",                   # 列の入れ替え
    "丸和物流の行とみどり建設の行を入れ替えて",   # 行と明示
])
def test_row_and_column_swaps_are_not_hijacked(meta, task):
    """★★ ここが緩むと、動いていた入れ替えを奪って壊す。"""
    assert ailine.swap_targets_are_cells(task, meta, "請求") is None


def test_the_same_cell_twice_is_not_a_swap(meta):
    assert ailine.swap_targets_are_cells("丸和物流の単価と丸和物流の単価を入れ替えて",
                                          meta, "請求") is None


# --- ② 解決（値は実表から読む・A' 原則）---------------------------------------------------

def test_the_values_come_from_the_table_not_the_model(meta):
    ok, r, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "丸和物流", "b": "みどり建設"}, meta,
        task="丸和物流の単価とみどり建設の単価を入れ替えて")
    assert ok, err
    assert r["_axis"] == "cell"
    assert r["_cells"] == [[2, 4], [4, 4]]
    assert r["_cell_values"] == [4800, 7200], r["_cell_values"]
    assert "単価" in r["_axis_label"]


def test_identical_contents_are_refused(meta):
    """★ 入れ替えても何も変わらない依頼は、✓ を出さずに断る。"""
    ok, _r, _i, err = ailine.verify_dsl_args(
        "SWAP", {"a": "丸和物流", "b": "近江スチール"}, meta,
        task="丸和物流の項目と近江スチールの項目を入れ替えて")
    if not ok:
        assert "同じ" in err or "決められません" in err, err


# --- ③ 番人（2 セル**だけ**が動いたことを証明する）----------------------------------------

def _book(tmp_path, name, rows):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(HEADERS)
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p


ARGS = {"_axis": "cell", "_cells": [[2, 4], [4, 4]], "_cell_values": [4800, 7200],
        "_target_sheet": "請求", "a": "丸和物流", "b": "みどり建設"}


def test_a_correct_cell_swap_passes(tmp_path):
    src = _book(tmp_path, "s.xlsx", ROWS)
    out = _book(tmp_path, "o.xlsx",
                 [["丸和物流", "配送", 12, 7200, 57600],
                  ["近江スチール", "鋼材", 5, 12000, 60000],
                  ["みどり建設", "内装", 9, 4800, 64800]])
    st, why = ailine.check_swap(out, ARGS, source_book=src)
    assert st == "pass", why


def test_swapping_the_whole_row_is_caught(tmp_path):
    """★★ これが今回の事故そのもの ── 行ごと入れ替えたら ✓ を出さない。"""
    src = _book(tmp_path, "s2.xlsx", ROWS)
    out = _book(tmp_path, "o2.xlsx",
                 [["みどり建設", "内装", 9, 7200, 64800],
                  ["近江スチール", "鋼材", 5, 12000, 60000],
                  ["丸和物流", "配送", 12, 4800, 57600]])
    st, why = ailine.check_swap(out, ARGS, source_book=src)
    assert st == "fail", why
    assert "セル変わっています" in why or "ではありません" in why, why


def test_doing_nothing_is_caught(tmp_path):
    src = _book(tmp_path, "s3.xlsx", ROWS)
    out = _book(tmp_path, "o3.xlsx", ROWS)
    st, _why = ailine.check_swap(out, ARGS, source_book=src)
    assert st == "fail"


def test_an_extra_cell_is_caught(tmp_path):
    """★ 「他は 1 セルも変わらず」── 余計に書いたら通さない。"""
    src = _book(tmp_path, "s4.xlsx", ROWS)
    out = _book(tmp_path, "o4.xlsx",
                 [["丸和物流", "配送", 12, 7200, 99999],
                  ["近江スチール", "鋼材", 5, 12000, 60000],
                  ["みどり建設", "内装", 9, 4800, 64800]])
    st, why = ailine.check_swap(out, ARGS, source_book=src)
    assert st == "fail", why


# --- ④ 実物で（LibreOffice を通す）--------------------------------------------------------

@pytest.mark.local
def test_only_two_cells_move_for_real(tmp_path):
    """★★ Namakoo の実例そのもの。"""
    import subprocess
    src = _book(tmp_path, "r.xlsx", ROWS)
    env = {**__import__("os").environ, "PYTHONPATH": str(REPO / "src")}
    r = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(src),
         "丸和物流の単価とみどり建設の単価を入れ替えて", "--copy", "--sheet", "請求"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=900, cwd=str(REPO), env=env)
    assert "✓" in r.stdout, r.stdout[-1500:]
    v = openpyxl.load_workbook(tmp_path / "r.out.xlsx", data_only=True)["請求"]
    assert v.cell(2, 4).value == 7200 and v.cell(4, 4).value == 4800
    assert v.cell(2, 1).value == "丸和物流" and v.cell(4, 1).value == "みどり建設"
    assert v.cell(2, 5).value == 57600, "金額まで動いている"
