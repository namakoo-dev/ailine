# 背景色は依頼文の色の語から決める（2026-10-01・依頼の項の台帳で D だった項目）。
#
# ★ 形: 色（color）は「COLOR_MAP に在るか」しか見ていなかった。「見出しを黄色にして」に
#   blue が返ると青く塗り、事後条件は宣言どおり青いかを確かめて ✓ を出す。
#
# 契約:
#   ① 依頼文の色の語が COLOR_MAP の鍵へ読めれば、食い違う LLM の色に勝つ（解釈行に出典）
#   ② 一致していれば何も足さない
#   ③ 読めない（語が無い・2 色・打ち消し・濃い/明るい・COLOR_MAP に無い薄い色）は何も変えない
#   ④ 列名と表の値の中の語（『青果』）は数えない
#   ⑤ 見出しが 1 行目でない表でも同じ
#   ⑥ 対応表の行き先はすべて COLOR_MAP の鍵（別の色の表を作らない）

import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import colors  # noqa: E402

HEAD = ["部門", "担当", "金額"]
DATA = [["青果", "白石", 300], ["鮮魚", "黒田", 500], ["精肉", "赤井", 700]]


def _book(tmp_path, header_row=1):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    for i in range(header_row - 1):
        ws.append(["売上一覧"] if i == 0 else [])
    ws.append(HEAD)
    for r in DATA:
        ws.append(r)
    wb.save(p)
    return p


def _meta(path, header_row=1):
    return {"sheets": ["売上"], "headers": {"売上": list(HEAD)},
            "header_rows": {"売上": header_row}, "path": str(path)}


def _fill(tmp_path, task, color, target="row:1", header_row=1):
    ok, r, inferred, err = ailine.verify_dsl_args(
        "FILL_COLOR", {"target": target, "color": color},
        _meta(_book(tmp_path, header_row), header_row), task=task)
    assert ok, err
    return r, inferred


# --- ① 食い違えば依頼文が勝つ -----------------------------------------------------------

def test_the_color_in_the_request_wins(tmp_path):
    r, inferred = _fill(tmp_path, "見出しを黄色にして", "blue")
    assert r["color"] == "yellow"
    assert r["_sources"]["color"] == "依頼文: 『黄』"
    assert "依頼文: 『黄』" in ailine.format_confirmation_line("FILL_COLOR", r, inferred)


def test_light_colors_go_to_the_light_key(tmp_path):
    assert _fill(tmp_path, "見出しを薄い青にして", "blue")[0]["color"] == "lightblue"
    assert _fill(tmp_path, "見出しを水色にして", "blue")[0]["color"] == "lightblue"


def test_the_request_wins_over_an_unsupported_llm_color(tmp_path):
    """LLM が COLOR_MAP に無い色を返しても、依頼文が読めればそちらで通る。"""
    assert _fill(tmp_path, "見出しを赤にして", "crimson")[0]["color"] == "red"


# --- ② 一致していれば黙る ----------------------------------------------------------------

def test_agreement_adds_nothing(tmp_path):
    r, _i = _fill(tmp_path, "見出しを黄色にして", "yellow")
    assert r["color"] == "yellow" and "color" not in r.get("_sources", {})


# --- ③ 読めない形は何も変えない -----------------------------------------------------------

def test_unreadable_requests_leave_the_llm_color(tmp_path):
    for task in ("見出しに色を付けて", "見出しを赤から青に変えて", "見出しを赤以外の色にして",
                 "見出しを濃い青にして", "見出しを薄い紫にして", "見出しを黄緑にして"):
        r, _i = _fill(tmp_path, task, "green")
        assert r["color"] == "green" and "color" not in r.get("_sources", {}), task


def test_words_that_only_contain_a_color_character(tmp_path):
    """赤字・空白は色ではない（「赤字の行を黄色に」は黄色）。"""
    r, _i = _fill(tmp_path, "赤字の行を黄色にして", "red")
    assert r["color"] == "yellow"


# --- ④ 列名と表の値の中の語は数えない ----------------------------------------------------

def test_values_in_the_table_are_not_colors(tmp_path):
    r, _i = _fill(tmp_path, "青果の行に色を付けて", "yellow", target="row:2")
    assert r["color"] == "yellow" and "color" not in r.get("_sources", {})
    r2, _i = _fill(tmp_path, "白石の行を黄色にして", "white", target="row:2")
    assert r2["color"] == "yellow"


# --- ⑤ 見出しが 1 行目でない表 -------------------------------------------------------------

def test_header_not_on_the_first_row(tmp_path):
    r, _i = _fill(tmp_path, "青果の行を黄色にして", "green", target="row:4", header_row=3)
    assert r["color"] == "yellow"
    r2, _i = _fill(tmp_path, "青果の行に色を付けて", "green", target="row:4", header_row=3)
    assert r2["color"] == "green"


# --- ⑥ 対応表の行き先は COLOR_MAP の鍵 -----------------------------------------------------

def test_every_color_word_lands_on_a_color_map_key():
    assert set(colors.COLOR_WORDS) <= set(colors.COLOR_MAP)
    for key, words in colors.COLOR_WORDS.items():
        for w in words:
            got = colors.read_color(f"{w}にして")
            assert got is not None and got[0] in colors.COLOR_MAP, (key, w, got)


# --- 変異: 配線を外すと LLM の色が黙って通る ---------------------------------------------

def test_without_the_wiring_the_llm_color_passes(tmp_path, monkeypatch):
    monkeypatch.setattr(colors, "read_color", lambda *a, **k: None)
    r, _i = _fill(tmp_path, "見出しを黄色にして", "blue")
    assert r["color"] == "blue"       # ← 直す前の挙動（嘘の ✓ の材料）
