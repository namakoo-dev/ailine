# 塊②（2026-08-25）── 分母を「走査」でなく「物理の使用範囲」から作る。
#
# ★ 中核 op の盲検が実測した 2 件（致命1・致命5）。検分者の処方をそのまま採る:
#   「除外そのものを ⚠ に昇格させ、分母を必ず入力側（物理行数・物理列数）から作れば、
#     致命 1・5・6 は同時に落ちる」
#
# 致命1: 見出しの無い列だけ置き去りになり、全行が入れ替わる
#   D1 が空で D2..D5 にデータがある表を並べ替えると、範囲が見出し由来（A〜C）なので
#   D 列だけ動かず、**全行で備考が別の商品の物に付け替わる**。
#   `check_sort` はキー列の単調性しか見ず、**行の同一性を一度も確かめない**ので ✓ が出る。
#   しかも画面の「変更点」も A〜C しか出さないため、**差分自体が壊れた列を隠す**。
#
# 致命5: 先頭列が空の末尾行が、処理からも分母からも消える
#   `_scan_last_row` は A 列を上から見て最初の空で止まる。末尾に A 列が空の行があると、
#   その行は処理されず、**分母にも数えられない**（「3行中1行が一致」＝真の分母は 5）。
#   ★ 表の**途中**の空きは ⚠ が出て △ に落ちるのに、**末尾だけ鳴らない**
#     （警告条件が「A 列に下方向の中身があるか」なので原理的に発火しない）。
#
# 契約:
#   ① 物理の使用範囲と、走査で得た範囲の食い違いを機械で出せる
#   ② 並べ替えは**行の同一性**を物理の列範囲で確かめる（他列が一緒に動いたか）
#   ③ 行がちぎれていたら ✓ でなく **×**（出力は既に壊れている）
#   ④ 末尾の欠けは分母の食い違いとして名指しされる
#   ⑤ 食い違いが無ければ 1 文字も増えない（誤爆しない）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import ailine  # noqa: E402


def _book(tmp_path, rows, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p


# --- ① 食い違いを出す器官 -------------------------------------------------------------

def test_extent_gap_sees_columns_beyond_the_headers(tmp_path):
    """D1 が空・D2 以下にデータ ── 見出し走査は 3 列、物理は 4 列。"""
    p = _book(tmp_path, [["商品", "単価", "数量", None],
                          ["a", 120, 3, "特売"], ["b", 200, 5, "通常"]])
    wb = openpyxl.load_workbook(p)
    gap = ailine.extent_gap(wb["売上"], header_row=1)
    assert gap["cols_scanned"] == 3, gap
    assert gap["cols_physical"] == 4, gap
    assert gap["cols_missing"] == 1, gap


def test_extent_gap_sees_rows_below_an_empty_first_cell(tmp_path):
    """末尾に A 列が空の行 ── A 列走査は 2 行、物理は 4 行。"""
    p = _book(tmp_path, [["商品", "単価"], ["a", 120], ["b", 200],
                          [None, 1500], [None, 2000]])
    wb = openpyxl.load_workbook(p)
    gap = ailine.extent_gap(wb["売上"], header_row=1)
    assert gap["rows_scanned"] == 2, gap
    assert gap["rows_physical"] == 4, gap
    assert gap["rows_missing"] == 2, gap


def test_extent_gap_is_silent_on_a_tidy_table(tmp_path):
    """⑤ 誤爆しない。"""
    p = _book(tmp_path, [["商品", "単価"], ["a", 120], ["b", 200]])
    wb = openpyxl.load_workbook(p)
    gap = ailine.extent_gap(wb["売上"], header_row=1)
    assert gap["rows_missing"] == 0 and gap["cols_missing"] == 0, gap


# --- ②③ 行がちぎれたら × ------------------------------------------------------------

def test_sort_fails_when_rows_were_torn(tmp_path):
    """★ 実測の再現: 見出しの無い列が置き去りになり、備考が別商品の物になった状態。
       今までは ✓ が出ていた。"""
    p = _book(tmp_path, [["商品", "単価", "数量", None],
                          ["ぶどう", 500, 2, "特売"],    # 本来は '高級'
                          ["かき", 200, 5, "通常"],      # 本来は '訳あり'
                          ["りんご", 120, 3, "高級"],    # 本来は '特売'
                          ["みかん", 80, 10, "訳あり"]]) # 本来は '通常'
    before = _book(tmp_path, [["商品", "単価", "数量", None],
                               ["りんご", 120, 3, "特売"],
                               ["みかん", 80, 10, "通常"],
                               ["ぶどう", 500, 2, "高級"],
                               ["かき", 200, 5, "訳あり"]], name="before.xlsx")
    status, reason = ailine.check_sort(p, {"col": "単価", "order": "desc"},
                                        source_book=before)
    assert status == "fail", f"行がちぎれているのに通した: {status} / {reason}"
    assert "行" in reason


def test_sort_passes_when_all_columns_moved_together(tmp_path):
    """⑤ 誤爆防止: 見出しの無い列も一緒に動いていれば通る。"""
    before = _book(tmp_path, [["商品", "単価", None],
                               ["りんご", 120, "特売"], ["ぶどう", 500, "高級"]],
                    name="before.xlsx")
    after = _book(tmp_path, [["商品", "単価", None],
                              ["ぶどう", 500, "高級"], ["りんご", 120, "特売"]])
    status, reason = ailine.check_sort(after, {"col": "単価", "order": "desc"},
                                        source_book=before)
    assert status == "pass", f"正しい並べ替えを落とした: {reason}"


# --- ④ 末尾の欠けを名指しする ---------------------------------------------------------

def test_trailing_rows_are_named_not_swallowed(tmp_path):
    """★ 表の途中の空きは鳴るのに、末尾だけ鳴らなかった。"""
    p = _book(tmp_path, [["商品", "単価"], ["a", 120], ["b", 200],
                          [None, 1500], [None, 2000]])
    args = {"col": "単価", "order": "desc"}
    ailine.check_sort(p, args)
    assert args.get("_unverified"), "末尾の 2 行が黙って消えた"
    assert any("2 行" in str(u["rows"]) or u["rows"] == 2 for u in args["_unverified"]), \
        args["_unverified"]
