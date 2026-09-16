# -*- coding: utf-8 -*-
"""★★ 左端の列が空の行があっても、表の終わりを見誤らない（2026-09-16）。

★ 出所: 買い手 2 体目の「3 桁区切りが合計行で効かない」を追ったら、根はもっと広かった。
  `helpers/AiLineHelpers.bas` の **11 か所**が、それぞれ独立に
  「**A 列が空になった行**で表は終わり」と書いていた。だから**左端が空の行**
  （合計行・続きの行）を持つ表は、そこから先が丸ごと見えない。

  実機で確かめた症状（どれも原本は壊れず × で止まる ── 事後条件は正直だった。
  ただし買い手は**その操作が一切できない**）:

      金額の列を3桁区切りにして  → 合計行だけ書式が付かず ×
      表全体を中央寄せにして      → 16 セル（合計行を除く）に掛けて ×
      表に罫線を引いて            → 同上 ×
      金額の多い順に並べ替えて    → 1 行しか対象にならず ×

★ 器官は在った: Python の走査は 2026-09-05 に「行の幅のどこかに値が在れば行」へ
  直してあった。直っていなかったのは **Basic 側だけ** ── 言語の境目で切れた片配線。

★ 直しは「11 か所を直す」でなく **1 関数（TableLastRow）に畳んで呼び出し側に持たせない**。
  Python から数を渡さない形にしたので、生成する .bas は 1 文字も変わらない
  （＝凍結した golden が動かない）。

★ この試験が守るのは **形**（左端が空の行を持つ表）であって、op ごとの結果ではない。
  検体の corpus にこの形が 1 つも無かったことが、そもそもの見落としの原因だった
  （消えたものは差分に出ない ── 負の被覆）。
"""
import re
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

HELPERS_BAS = REPO / "src" / "ailine" / "helpers" / "AiLineHelpers.bas"

#: A 列だけを見て表の終わりを決める書き方。1 つでも残っていたら同じ穴が開く。
_A_COLUMN_WALK = re.compile(r"Do While \w+\.getCellByPosition\(0,\s*\w+\)\.getString\(\) <> \"\"")


def _bas() -> str:
    return HELPERS_BAS.read_text(encoding="utf-8", errors="replace")


def test_the_single_place_that_decides_where_the_table_ends_exists():
    """★ 空回りの検出 ── 畳んだ先が消えていたら、下の試験は全部素通りする。"""
    src = _bas()
    assert "Function TableLastRow(" in src, "表の終わりを決める 1 本が無い"


def test_no_helper_decides_the_end_of_the_table_by_itself():
    """★★ 片配線の番人 ── A 列だけを見る書き方が**1 つも**残っていないこと。

    ★ 数を上げるだけの縛りにしない（>= N は 2 つ消しても緑になる）。
      ここは **0 でなければならない** ── 1 つ残れば、その op だけが同じ穴を持つ。
    """
    left = _A_COLUMN_WALK.findall(_bas())
    assert not left, (
        f"A 列だけで表の終わりを決めている箇所が {len(left)} 本残っている ── "
        "左端が空の行を持つ表で、その先が見えなくなる。TableLastRow を使うこと")


def test_the_end_is_decided_by_the_width_of_the_header_not_by_one_column():
    """★ 畳んだ先が「A 列を見るだけ」に退行していないこと（中身を縛る）。"""
    src = _bas()
    body = src.split("Function TableLastRow(")[1].split("End Function")[0]
    assert "getCellByPosition(lastCol, headerRow)" in body, "見出しの幅を測っていない"
    assert "For c = 0 To lastCol" in body, "行の幅を見ていない（A 列だけに戻っている）"


# --- 実機（ここからが本番。上の 3 本は形だけを守る）-----------------------------

def _book(tmp_path, rows, name="orders.xlsx"):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "受注台帳"
    ws.append(["受注番号", "取引先", "納期", "金額"])
    for r in rows:
        ws.append(list(r))
    p = tmp_path / name
    wb.save(p)
    return p


#: ★ 左端が空の行の**2 つの形**。どちらも実物の帳票に出る。
DATA_WITH_TOTAL = [("A-001", "丸山工業", "2026-09-20", 97500),
                    ("A-002", "北斗精機", "2026-10-05", 50000),
                    ("A-003", "西村工業", "2026-09-28", 1250000),
                    (None, None, "合計", 1397500)]          # 末尾の合計行
DATA_WITH_GAP = [("A-001", "丸山工業", "2026-09-20", 97500),
                  (None, "（続き）北斗精機", "2026-10-05", 50000),  # 途中の続き行
                  ("A-003", "西村工業", "2026-09-28", 1250000),
                  ("A-004", "東和商事", "2026-11-02", 384000)]


def _run(book, task, timeout=300):
    import os
    import subprocess
    return subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(book), task,
         "--sheet", "受注台帳", "--timeout", str(timeout)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout + 600,
        env={**os.environ, "PYTHONPATH": str(REPO / "src")})


@pytest.mark.local
@pytest.mark.parametrize("task, rows, want", [
    # ★ 期待するのは「**全部の行**に届いたこと」── 数を明示して縛る。
    ("金額の列を3桁区切りにして", DATA_WITH_TOTAL, "4 行に桁区切り"),
    ("表全体を中央寄せにして", DATA_WITH_TOTAL, "20 セルの中央揃え"),
    ("表に罫線を引いて", DATA_WITH_TOTAL, "20 セルの罫線"),
])
def test_it_reaches_the_row_whose_first_cell_is_empty(tmp_path, task, rows, want):
    book = _book(tmp_path, rows)
    r = _run(book, task)
    assert r.returncode == 0, r.stdout[-800:]
    assert want in r.stdout, f"全部の行に届いていない（期待: {want}）\n{r.stdout[-800:]}"


@pytest.mark.local
def test_a_gap_in_the_middle_no_longer_truncates_the_sort(tmp_path):
    """★ 途中に左端の空がある表でも、最後まで並べ替えること。"""
    import openpyxl
    book = _book(tmp_path, DATA_WITH_GAP)
    r = _run(book, "金額の多い順に並べ替えて")
    assert r.returncode == 0, r.stdout[-800:]
    ws = openpyxl.load_workbook(book)["受注台帳"]
    got = [ws.cell(row=i, column=4).value for i in range(2, 6)]
    assert got == sorted(got, reverse=True), (got, r.stdout[-600:])


@pytest.mark.local
def test_the_total_row_is_still_left_out_of_the_sort(tmp_path):
    """★★ 反対側の検算 ── 届くようにした代わりに、合計行まで並べ替えてはいけない。

    ★ ここが無いと「直した」の意味が半分になる。届く範囲を広げる変更は、
      **除外の仕組み（_sort_end_row / _skip_rows）がまだ効いているか**まで見て初めて完了する。
    """
    import openpyxl
    book = _book(tmp_path, DATA_WITH_TOTAL)
    r = _run(book, "金額の多い順に並べ替えて")
    assert r.returncode == 0, r.stdout[-800:]
    ws = openpyxl.load_workbook(book)["受注台帳"]
    assert ws.cell(row=5, column=3).value == "合計", (
        "合計行が並べ替えに巻き込まれた\n" + r.stdout[-800:])
    got = [ws.cell(row=i, column=4).value for i in range(2, 5)]
    assert got == sorted(got, reverse=True), (got, r.stdout[-600:])
