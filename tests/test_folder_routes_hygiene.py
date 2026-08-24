# 第三波 S1/S2/S5/S6 ── 修正より先に凍結した赤い検体。
#
# 契約:
#   S1 自己除外（ailine 産の出力を入力から外す）を **scan / stack / run の 3 経路すべて**が
#      行う。実測: 2 冊照合の出力が残ったフォルダに scan を掛けると「3 ファイル中 2」と
#      分母が汚れ、自分の出力を「取れなかった」と ⚠ で名指ししていた（stack は正しかった）。
#      ★ 実装は 1 つ（stack.split_own_outputs）── 三度目の書き写しを禁じる。
#   S2 out=<path> と宣言しておいてファイルを作らなかったら、作らなかったと言う。
#   S5 「A 186300 / B 0」の 0 が金額 0 か 1 行も無いのか読めた（『なし（0 行）』）。
#   S6 複数ファイルの入口が ops の一覧から辿れる（引数の形も argparse から生成）。

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import stack as stack_core  # noqa: E402
from ailine_core.cli_render import render_folder_routes  # noqa: E402
from ailine_core.match import KeyGroup, side_pair  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_golden_transcripts import _isolate, _run_main  # noqa: E402


def _book(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


# --- S1 自己除外が 3 経路すべてに在る --------------------------------------------------

def test_split_own_outputs_is_the_single_implementation():
    """★ 呼び出し側が自前のループを持たないこと（書き写しの再発防止）。"""
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    assert src.count("split_own_outputs") >= 3, "3 経路すべてが共通関数を呼んでいない"
    assert "is_own_output(p)" not in src, \
        "自己除外のループが手書きで残っている（1 箇所に畳んだはず）"


def _folder_with_own_output(tmp_path):
    folder = tmp_path / "f"
    folder.mkdir()
    _book(folder / "a.xlsx", [["取引先", "金額"], ["甲社", 1000]])
    _book(folder / "b.xlsx", [["取引先", "金額"], ["乙社", 2000]])
    # ailine 産の出力を模す（印は stack_core の CREATOR_MARKS が唯一の接点）
    own = folder / "out_照合.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["キー", "A側 合計", "B側 合計"])
    wb.save(own)
    return folder, own


def test_scan_excludes_ailine_own_output(tmp_path, monkeypatch, capsys):
    """S1 の本体: scan の分母が自分の出力で汚れない。"""
    _isolate(monkeypatch, tmp_path)
    folder, own = _folder_with_own_output(tmp_path)
    monkeypatch.setattr(stack_core, "is_own_output", lambda p: p.name == own.name)
    rc, out = _run_main(["scan", str(folder)], capsys)
    assert "2 ファイル中 2 照合できた" in out, f"分母が自分の出力で汚れている: {out}"
    assert "自分の出力" in out and own.name in out, f"黙って減らした（開示が無い）: {out}"


def test_stack_still_excludes_its_own_output(tmp_path, monkeypatch, capsys):
    """誤爆防止: 元から正しかった stack の挙動を壊していない。"""
    _isolate(monkeypatch, tmp_path)
    folder, own = _folder_with_own_output(tmp_path)
    monkeypatch.setattr(stack_core, "is_own_output", lambda p: p.name == own.name)
    out_path = tmp_path / "stacked.xlsx"
    rc, out = _run_main(["stack", str(folder), "--out", str(out_path)], capsys)
    assert "自分の出力" in out, out
    assert "2 ファイル中" in out, out


# --- S2 宣言と実体 -------------------------------------------------------------------

def test_stack_says_it_created_no_file_when_nothing_to_stack(tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    empty = tmp_path / "empty"
    empty.mkdir()
    out_path = tmp_path / "never.xlsx"
    rc, out = _run_main(["stack", str(empty), "--out", str(out_path)], capsys)
    assert not out_path.exists(), "前提: ファイルは作られないこと"
    assert "作っていません" in out, f"out= と宣言したのに、作らなかったと言っていない: {out}"


# --- S5 0 と『無い』の区別 -------------------------------------------------------------

def test_side_pair_distinguishes_zero_from_absent():
    absent = KeyGroup(key_display="甲社", a_count=2, a_sum=186300.0,
                      b_count=0, b_sum=0.0, diff=186300.0, state="A のみ",
                      a_rows=[], b_rows=[])
    assert "なし（0 行）" in side_pair(absent), f"0 と『無い』が区別できない: {side_pair(absent)}"
    real_zero = KeyGroup(key_display="乙社", a_count=1, a_sum=100.0,
                          b_count=1, b_sum=0.0, diff=100.0, state="+100",
                          a_rows=[], b_rows=[])
    assert "なし" not in side_pair(real_zero), \
        f"本物の金額 0 を『無い』と偽った: {side_pair(real_zero)}"


# --- S6 入口が辿れる -----------------------------------------------------------------

def test_ops_shows_the_multifile_routes(tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    rc, out = _run_main(["ops"], capsys)
    assert rc == 0
    for token in ("ailine scan", "ailine stack", "2 冊"):
        assert token in out, f"複数ファイルの入口が一覧に無い（{token}）: {out[-600:]}"


def test_folder_routes_use_argparse_shapes_not_a_template():
    """★ 雛形で書いた初版は即ずれた（verify はフォルダを取らないのに <フォルダ> と出た）。
       引数の形は登録簿から来ること ── ここを固定文字列に戻すと赤くなる。"""
    lines = render_folder_routes([("verify", "検算する", "<out> <sources>"),
                                   ("scan", "棚卸し", "<folder>")])
    joined = "\n".join(lines)
    assert "ailine verify <out> <sources>" in joined, joined
    assert "verify <フォルダ>" not in joined, "雛形に戻っている"
