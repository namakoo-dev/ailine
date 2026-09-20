# -*- coding: utf-8 -*-
"""`verify` が「二者が立っていない」一覧を咎める（2026-09-20・盲検 4 体目 ③）。

★★ 出所: 買い手の一覧は「請求元が全部空・宛先に仕入先名」だったのに、
  **`ailine verify` はそれを `✓ 破れはありません` で通していた**。

★★ なぜ通ったか（読んで確かめた）: この検算器は**抽出の規則を再現しない**設計で、
  測るのは「その値が出所の冊の中に在るか」＝**含有**。発行元の社名はその冊に確かに
  在るので、宛先の欄に入っていても含有は破れない。
  ★ 「値が**どの列に**在るか」は含有では測れない ── これは設計の穴ではなく、
    測っているものが違うという話。規則を再現すれば測れるが、それは**恒真**になる。

★★ 直し: 測るのは**道具自身の宣言の辻褄**。請求書は必ず**二者**なので、片方が
  「もう片方と同じ名前だから」という理由で空欄になっているなら、道具はどちらが発行元か
  決められていない ── その回の宛先も請求元も根拠にできない。
  ★ 道具は検分でそう**自白している**（`SAME_PARTY_MARK`）。読むのはその印だけで、
    どう抽出したかは一切見ない ── 独立性の線は 1 ミリも動かさない。
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from ailine_core import forms_collect, verify_forms  # noqa: E402


def _invoice(folder: Path, issuer: str, addressee: str) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求書"
    for at, v in {"A1": "請求書", "A3": issuer, "A5": addressee,
                  "A7": "下記のとおりご請求申し上げます",
                  "A9": "小計", "B9": 129000, "A10": "消費税", "B10": 12900,
                  "A11": "合計", "B11": 141900,
                  "A13": "請求日", "B13": "2026-09-18",
                  "A14": "請求番号", "B14": "A-1024"}.items():
        ws[at] = v
    p = folder / "請求書.xlsx"
    wb.save(p)
    wb.close()
    return p


def _forms_then_verify(d: Path, issuer: str, addressee: str) -> tuple:
    folder = d / "請求書群"
    folder.mkdir()
    _invoice(folder, issuer, addressee)
    out = d / "一覧.xlsx"
    env = {**os.environ, "PYTHONPATH": str(REPO / "src"),
           "AILINE_HOME": str(d / "home"), "PYTHONUTF8": "1"}

    def run(*args):
        return subprocess.run([sys.executable, "-m", "ailine", *args], cwd=str(REPO),
                              capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=900, env=env)

    made = run("forms", str(folder), "--out", str(out))
    assert made.returncode == 0, made.stdout[-500:]
    return run("verify", str(out), str(folder)), out, folder


def test_a_list_where_the_two_parties_collapsed_is_caught():
    """★★ 事故そのもの ── 二者が立っていない一覧を `✓` で通さないこと。"""
    with tempfile.TemporaryDirectory() as td:
        got, _out, _folder = _forms_then_verify(
            Path(td), "ナギ商会株式会社", "ナギ商会株式会社 御中")
    assert got.returncode != 0, "★ まだ ✓ で通している:\n" + got.stdout[-600:]
    assert "二者として立っていない" in got.stdout, got.stdout[-600:]
    # ★ 咎める時は**どの冊のどの項目**かを名指しすること（数だけにしない）。
    #   ★★ ここは**本物の画面**を見る ── 初版は判定を書き写した偽物（`_breaks_from`）を
    #     測っており、製品については何も言っていなかった（恒真）。消した。
    assert "請求書.xlsx" in got.stdout and "請求元" in got.stdout, got.stdout[-600:]


def test_a_healthy_list_is_still_passed():
    """★★ 陰性対照 ── 二者が立っている一覧は今までどおり通ること。

    ★ これが無いと「咎めた」のか「何でも咎める」のか分からない。
    """
    with tempfile.TemporaryDirectory() as td:
        got, _out, _folder = _forms_then_verify(
            Path(td), "大東金属 株式会社", "みどり商事 御中")
    assert got.returncode == 0, "★ 正しい一覧を咎めている:\n" + got.stdout[-600:]
    assert "破れはありません" in got.stdout, got.stdout[-600:]


def test_the_mark_is_declared_in_one_place():
    """★★ 印は**宣言**で、両側に字面を書かないこと。

    ★ 文言を直した日に片方が黙る形（今週 7 回踏んだ「番人を字面で書いた」の親戚）。
    ★ 検算が読むのは `forms_collect` 経由の印だけ ── `form_read` は import しない
      （抽出の規則を再現しないという線は不変・既存の番人が禁じている）。
    """
    from _product_source import count_in_code
    # ★★ 字面はコードの中に**1 回だけ**（＝宣言の所だけ）。
    #   ★ 変異試験が指した: 印を使わず字面に書き戻しても、値が同じなので何も壊れない
    #     ── 「宣言が在るか」ではなく「**字面が 2 箇所に無いか**」を縛る。
    #   ★ 数えるのはコードだけ（コメントの言及に当たらない ── 今日 2 度踏んだ形）。
    assert count_in_code("宛先と同じ名前") == 1, (
        "★ 印の字面がコードに 2 箇所以上ある（片方を直した日にもう片方が黙る）")
    src = (REPO / "src" / "ailine_core" / "verify_forms.py").read_bytes().decode("utf-8")
    assert "forms_collect.SAME_PARTY_MARK" in src, "★ 検算が印を使っていない"
    assert "宛先と同じ名前" not in src.replace("SAME_PARTY_MARK", ""), (
        "★ 検算の側に字面を書き写している")
    assert forms_collect.SAME_PARTY_MARK, "★ 登録簿が印を配っていない"


def test_only_declarations_travel_through_the_registry():
    """★★ 登録簿を**抜け道にしない**こと（2026-09-20）。

    ★ 既存の番人（`test_verify_forms`）は、検算器が**使っている名前**を AST で見て
      `form_read` を呼んでいないことを縛っている ── そこは正しいので作り直さない。
    ★ 今日、印を渡すために `forms_collect` を経由させた。同じ道で**関数**を通せば、
      既存の番人は気づかないまま規則が再現されうる ── そこを塞ぐのがこの試験。
    ★ 登録簿が `form_read` から受けるものは、**呼べないもの**（宣言）に限る。
    """
    import ast
    import inspect

    tree = ast.parse(inspect.getsource(forms_collect))
    names = []
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and "form_read" in (node.module or ""):
            names += [a.asname or a.name for a in node.names]
    assert names, "★ 登録簿が form_read から何も取っていない（この検査が空回りしている）"
    for n in names:
        got = getattr(forms_collect, n, None)
        assert not callable(got), (
            f"★ 登録簿経由で**呼べるもの**が渡っている: {n} ── 規則を再現する道になる")
