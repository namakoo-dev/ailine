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
    # ★ 2026-09-16: 以前は「`✓` が出ない」と書いてあったので ✓ の出現を見ていた。
    #   いまは「3 行とも計算し ✓」なので、同じ検査が**逆の意味**で通ってしまう。
    #   記号だけでなく**結末**まで見る。
    assert "✓" in row, row
    assert "3 行とも" in row, f"README が結末を書いていない: {row}"


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

    # ★★ 2026-09-16: ここは長らく「✓ を出さない」を縛っていた。実体が変わったので
    #   期待値を入れ替えた ── この検体は冒頭で「製品の意味を決めない・縛るのは
    #   **README と実体の一致**だけ」と宣言しており、その線に従う。
    #   ★ 何が変わったか: 左端（商品名）が空の行に**届いていなかった**ので
    #     「4行目」を名指しして × で止まっていた。断りは正直だったが、
    #     計算そのものは可能だった（売上 1500 − 原価 900 は商品名が無くても引ける）。
    #     表の終わりを見誤らないよう直した結果、3 行とも計算して ✓ になった。
    #   ★ 芯（機械で確かめられないことに ✓ を出さない）は失われていない ──
    #     実機を含む 8 本以上の試験が別に縛っている（入れ替える前に数えた）。

    # ① 最後の行まで届くこと（届かないのが、そもそもの欠陥だった）
    assert "3 行を検証" in said, said[-600:]
    # ② 全部できたので ✓ で終わること
    assert "✓" in said, said[-600:]
    # ③ ★ 反対側の検算 ── 「できた」と言う以上、書けなかった行を名指しする必要はない。
    #   逆に **✓ と「〜行目」が同居したら**、それは「できたが一部できていない」という嘘。
    assert not re.search(r"\d+行目[:：]", said), (
        "✓ と『〜行目』の名指しが同居している（できたと言いながら書けていない）\n"
        + said[-600:])

    # ④ ★ README が言っている記号と、実際に出た記号が一致すること
    row = _readme_row_9()
    mark = "×" if "×" in said else ("△" if "△" in said else ("✓" if "✓" in said else "?"))
    assert mark in row, (
        f"実物は『{mark}』で終わったのに、README は違うことを書いている: {row}")


def test_the_screenshot_note_discloses_that_the_behaviour_changed():
    """★★ 番人が**片側にしか無かった**（2026-09-08 に盲検が見つけた）。

    README は同じ検体・同じ依頼を 2 箇所で語っていた:
      ・手順表 項番 9  → `✓` が出ず `×` で止まる  ← 上の実機の番人が守っている
      ・写真 ③ の説明  → `△` で通す              ← ★ 誰も守っていなかった

    写真は 2026-08 の実物で、2026-09-05 に「走査が表の終わりに届かなかった回は
    ✓ も △ も名乗らない」と決めたため、いまは `×` になる。★ 写真を隠さず、
    変わったことを**開示して持つ**と決めた ── この試験はその開示文を凍結する。
    """
    text = (REPO / "README.md").read_text(encoding="utf-8")
    i = text.find("3_partial_with_warning.jpg")
    assert i > 0, "③ の写真が README から消えている"
    block = text[max(0, i - 1200):i]
    assert "その後に挙動が変わっています" in block, (
        "写真 ③ の『挙動が変わった』開示が消えている ── "
        "消すなら、写真を撮り直して説明も現物に合わせること")
    assert "×" in block and "項番 9" in block, block[-300:]
