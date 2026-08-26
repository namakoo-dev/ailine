# 複数ファイルの盲検・致命①⑨（2026-08-26）── 検算が書き手と同じ関数で除外し、
# 両方が同じ間違いをして一致する（恒真）。
#
# ★ 検分者の再現をそのまま検体にする:
#     区切りの空行がある表で、**3 列すべて埋まった**売上 1,000 円が
#     「合計行」として消え、stack も verify も exit 0。--json の mismatches も空。
#
# ★ 根: verify.py が stack.py と同一の total_row.split_total_rows_multi を呼ぶ。
#   docstring は「同じ検出でないと**別の恒真**（片方だけ取り逃がす）になる」と書いており、
#   共有は意図的な判断だった ── だが同じ罠の裏返しでしかない。
#
# 契約:
#   ① 検算が「書き手と同じ規則で」落とした行を、機械の値として持つ
#   ② その行をファイル名と行番号で名指しして人に見せる
#   ③ 判定（exit）は変えない ── 正しい合計行を持つ表を全部不合格にしない
#   ④ 落とす行が無ければ 1 文字も増えない（誤爆しない）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
from ailine_core import cli_render, verify  # noqa: E402


def _book(p: Path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for r in rows:
        ws.append(r)
    wb.save(p)
    return p


# --- ①② 落とした行を名指しする -------------------------------------------------------

def test_report_names_the_row_it_dropped():
    """② 実測の形: 『⚠ 01.xlsx の 4行目を…除外しています』が出ること。"""
    result = {
        "row_count": {"source": 3, "output": 3},
        "sums": {"金額": {"source": 6000, "output": 6000}},
        "mismatch": None, "mismatches": [],
        "unbacked_exclusions": [{"kind": "unbacked_exclusion", "file": "01.xlsx", "row": 4}],
    }
    lines = cli_render.render_verify_report("out.xlsx", "AA", result)
    named = [ln for ln in lines if "01.xlsx" in ln and "4行目" in ln]
    assert named, f"落とした行を名指ししていない: {lines}"
    assert named[0].startswith("⚠"), named[0]
    assert "裏取りしていません" in named[0], named[0]


def test_no_line_when_nothing_was_dropped():
    """④ 誤爆しない。"""
    result = {"row_count": {"source": 3, "output": 3},
               "sums": {"金額": {"source": 6000, "output": 6000}},
               "mismatch": None, "mismatches": [], "unbacked_exclusions": []}
    lines = cli_render.render_verify_report("out.xlsx", "AA", result)
    assert not [ln for ln in lines if "除外しています" in ln], lines


def test_missing_key_is_treated_as_nothing_dropped():
    """後方互換: このキーを持たない呼び出し元でも壊れない。"""
    result = {"row_count": {"source": 1, "output": 1}, "sums": {},
               "mismatch": None, "mismatches": []}
    cli_render.render_verify_report("out.xlsx", "AA", result)


# --- ① 機械の値として持つ（本番の経路で）---------------------------------------------

def test_the_dropped_row_is_recorded_as_data_not_prose(tmp_path):
    """① 表示文から読み取らせない ── 機械の値として在ること。

    ★ 検分者の検体そのもの: 見出し / 甲社1000 / (空行) / 乙社1000 / 丙社2000。
      直上が空行という理由だけで、3 列すべて埋まった乙社の行が落ちる。
    """
    src = _book(tmp_path / "01.xlsx",
                 [["日付", "得意先", "金額"],
                  ["2026-01-05", "甲社", 1000],
                  [None, None, None],
                  ["2026-01-07", "乙社", 1000],
                  ["2026-01-08", "丙社", 2000]])
    verify._LAST_DROPPED.clear()
    expected, values = verify._expected_rows_for_source(
        src, ["日付", "得意先", "金額"], "得意先", ["金額"], sheet_name="売上")
    dropped = verify._LAST_DROPPED.get(str(src), [])
    assert dropped, ("前提が崩れた: この検体で行が落ちなくなった（"
                     "total_row の規則が変わったなら検体を作り直すこと）")
    assert 4 in dropped, f"乙社の行(4行目)が落ちたのに記録されていない: {dropped}"
    assert 4 not in expected, "前提: 落ちた行は expected に入っていない"
