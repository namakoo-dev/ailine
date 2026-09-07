# -*- coding: utf-8 -*-
"""README の看板デモ（項目 9）が、書いてある通りに動くこと（2026-09-07）。

★★ 出所（外部の UX 検品）: README は項目 9 をこう約束していた ──
  「`✓` が出ない。『検証できていない行がある』と名指しで出て **`△`** に落ちる」。
  実物は **`×`**（「4行目: 式が期待形でない」）で止まり、原本は無変更だった。
  ★ しかも理由も違う ── 「検証できていない」ではなく「**書いていない行がある**」。

★ README 自身が「9 と 10 がこの道具の山場」と書いている。**その山場が手書きの約束**
  だったので腐った。実機で走らせて突き合わせる番人をここに置く。

★ この試験は**製品の意味を決めない** ── 「商品名の無い行に利益を書くべきか」は仕様の
  判断で、ここでは触らない。縛るのは「**README と実体が一致していること**」だけ。
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
DEMO = REPO / "demo" / "3_売上_欠けあり.xlsx"
TASK = "売上から原価を引いた利益の列を作って"


def _readme_row_9() -> str:
    text = (REPO / "README.md").read_text(encoding="utf-8")
    for line in text.splitlines():
        if "3_売上_欠けあり" in line and line.lstrip().startswith("|"):
            return line
    return ""


def test_the_readme_still_describes_this_scenario():
    """★ 先に「約束の在りか」を掴む ── 消えていたら試験は無意味になる（恒真を切る）。"""
    row = _readme_row_9()
    assert row, "README から項目 9 が消えている（番人が守る対象を失っている）"
    assert "✓" in row, row


@pytest.mark.local
def test_the_flagship_demo_behaves_as_the_readme_says(tmp_path):
    """★ 実機（LLM + LibreOffice）で走らせて、README の記述と突き合わせる。"""
    assert DEMO.is_file(), f"デモ検体が無い: {DEMO}"
    book = tmp_path / DEMO.name
    shutil.copy2(DEMO, book)
    before = book.read_bytes()

    got = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(book), TASK],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(REPO), timeout=900,
        env={**os.environ, "PYTHONPATH": str(REPO / "src")})
    said = (got.stdout or "") + (got.stderr or "")

    # ① ✓ を出さない（README の芯・ここだけは何があっても守る）
    assert "✓" not in said, said[-600:]
    # ② 書けなかった行を**名指し**する
    assert re.search(r"\d+行目", said), said[-600:]
    # ③ 原本は 1 バイトも変わらない
    assert book.read_bytes() == before, "原本が変わっている"
    # ④ 作業結果は .out.xlsx に残る
    assert (book.parent / (book.stem + ".out.xlsx")).is_file(), said[-400:]

    # ⑤ ★ README が言っている記号と、実際に出た記号が一致すること
    row = _readme_row_9()
    mark = "×" if "×" in said else ("△" if "△" in said else "?")
    assert mark in row, (
        f"実物は『{mark}』で終わったのに、README は違うことを書いている: {row}")
