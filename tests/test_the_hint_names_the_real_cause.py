"""効かなかった時の「心当たり」が、**本当の原因**を先に言う（2026-09-22）。

★★ 事故（盲検 6 体目 ⑩・**致命**・こちらで再現）: タイトル行が結合された経費精算書で
  集計が失敗し、道具はこう案内した ──

    心当たり: シート『精算』に結合セルが 1 件あります（B1:F1）
    → …結合を解除してからお試しください

  買い手は解除を探して**完全に行き止まり**になった（この道具に解除は無い・①で断る側に
  直したばかり）。買い手:「タイトル付きの精算書は、ailine では一切集計できません」

★★ こちらで切り分けた**事実だけ**（原因はまだ分かっていない）:

    元のまま（結合あり・A 列が空）  → ×
    結合だけ外す（A 列は空のまま）  → **× 同じく失敗**（案内どおりにしても直らない）
    A 列を落とす（結合は無し）      → **✓ 通った**（3 グループを検証）

★★ つまり **案内が嘘**なのは確か ──「結合を解除してからお試しください」と言うが、
  解除しても直らないし、そもそもこの道具に解除は無い。
★ 一度「真因は 1 列目が空で走査が止まること」と書いたが、**それは私の誤りだった**
  （`header_row=1` で測ったから出た数字で、正しく 3 で測ると走査は止まっていない）。
  ★ ここに残すのは**確かめた事実と、嘘だった案内の処置**だけにする。

★ 直している最中に 3 回踏んだ（どれも今日の道具が捕まえた／捕まえられた）:
  ① 鍵の名を `rows_missed` と綴った（正は `rows_missing`）── 字面を確かめずに書いた
  ② 2 つの心当たりを 1 つの `try` に入れ、片方が転んで**両方消えた**
  ③ `table_scan.extent_gap` と書いて `NameError` を `except` に飲まれた
     ── ★ 今日入れた門（ruff F821 / pyright）は**これを名指しで捕まえる**。
       実機で試す前に門へ掛けていれば 1 分で済んだ。
"""
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import ailine  # noqa: E402


def _book(path, first_col_blank: bool, merge: bool):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "精算"
    off = 1 if first_col_blank else 0
    if merge:
        ws["B1"] = "2026年8月 経費精算書"
        ws.merge_cells("B1:F1")
    head = ["社員", "日付", "科目", "金額"]
    for i, h in enumerate(head):
        ws.cell(row=3, column=1 + off + i).value = h
    for r, row in enumerate([("佐藤", "2026-08-02", "旅費交通費", 1280),
                             ("鈴木", "2026-08-07", "会議費", 3500)], start=4):
        for i, v in enumerate(row):
            ws.cell(row=r, column=1 + off + i).value = v
    wb.save(path)
    return path


def test_the_header_row_is_not_guessed(tmp_path):
    """★ 見出し行は**渡されたものを使う**（推測で 1 行目と決めない）。

    ★ 渡さないと、見出しが 3 行目の冊で上の空行を「走査が止まった」と誤読する ──
      実際にそれで誤った心当たりを書きかけた（今日 1 回）。
    """
    p = _book(tmp_path / "a.xlsx", first_col_blank=True, merge=True)
    with_hr = ailine.likely_cause_of_no_change(p, "精算", header_row=3)
    assert all("走査" not in ln for ln in with_hr), (
        "★ 見出し行を渡しても走査の心当たりが出ている: " + repr(with_hr))


def test_the_merge_hint_does_not_send_them_to_a_dead_end(tmp_path):
    """★ この道具に結合の解除は無い ── 無い道へ送らない（導線が嘘なら無い方がまし）。"""
    p = _book(tmp_path / "b.xlsx", first_col_blank=False, merge=True)
    joined = "".join(ailine.likely_cause_of_no_change(p, "精算"))
    assert "結合セル" in joined, joined
    assert "この道具に結合を解除する操作はありません" in joined, (
        "★ 無い操作へ案内している: " + joined)


def test_a_healthy_table_says_nothing(tmp_path):
    """★ 陰性対照 ── 普通の表に心当たりを並べない（毎回出る助言は読まれなくなる）。

    ★ この冊は**見出しが 3 行目**（タイトル付きの精算書の形）。見出し行を渡さないと
      1 行目と決め打ちになり、上の空行を「走査が止まった」と誤爆する ── 実際にそうなり、
      呼び出し側から header_row を渡す形に直した。**推測で決めない。**
    """
    p = _book(tmp_path / "d.xlsx", first_col_blank=False, merge=False)
    assert ailine.likely_cause_of_no_change(p, "精算", header_row=3) == []


@pytest.mark.local
def test_the_real_specimen_now_names_the_column(tmp_path):
    """★ 実機 ── 買い手の冊と同じ形で、真因が画面に出ること。"""
    import os
    import subprocess
    p = _book(tmp_path / "e.xlsx", first_col_blank=True, merge=True)
    repo = Path(ailine.__file__).resolve().parents[2]
    r = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(p), "科目ごとに金額を集計して",
         "--copy", "--out", str(tmp_path / "out.xlsx"), "--timeout", "150"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(repo), env={**os.environ, "PYTHONPATH": str(repo / "src"),
                            "AILINE_HOME": str(tmp_path / "home")})
    assert "この道具に結合を解除する操作はありません" in r.stdout, r.stdout[-600:]
