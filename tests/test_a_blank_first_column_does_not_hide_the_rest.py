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
    """★ 畳んだ先が「A 列を見るだけ」に退行していないこと（中身を縛る）。

    ★★ 2026-09-22: ここは**字面**（`getCellByPosition(lastCol, headerRow)`）で
      「幅を測っている」ことを確かめていた。列の走査を `HeaderLastCol` へ畳んだら、
      契約は満たしているのに**番人だけが落ちた** ── 今日 4 本目の字面の番人。
      ★ 契約で書き直す: 幅は**別の関数から受け取る**（自分で A 列だけ見ない）。
    """
    src = _bas()
    body = src.split("Function TableLastRow(")[1].split("End Function")[0]
    assert "HeaderLastCol(" in body, "見出しの幅を測っていない（幅を決める関数を呼んでいない）"
    assert "For c = 0 To lastCol" in body, "行の幅を見ていない（A 列だけに戻っている）"


def test_the_header_width_does_not_stop_at_the_first_empty_column():
    """★★ 2026-09-22（盲検 6 体目 ⑩の真因）── **列**の走査にも同じ罠が在った。

    ★ 旧版は 0 列目から「最初の空」で止めて幅を決めていた。**A 列が空の表**では
      1 周目で止まって幅が -1 になり、`TableLastRow` が見出し行を返す →
      呼び出し側の `If lastRow < headerRow + 1 Then Exit Sub` で**黙って何もしない**。
    ★ 09-16 に**行**から外した前提が、**列**に残っていた（同じ冊・同じコメント）。
    """
    src = _bas()
    body = src.split("Function HeaderLastCol(")[1].split("End Function")[0]
    assert "gapRun" in body, "空列で即座に止める形に戻っている"
    assert "SCAN_COLS_MAX" in body, "走査の上限が宣言から来ていない"
    # ★ 書き写しが復活していないこと（畳んだ意味が消える）
    assert "Do While oSheet.getCellByPosition(lastCol" not in src, (
        "列の走査が書き写されています ── HeaderLastCol に畳んでください")


def test_both_languages_use_the_same_numbers():
    """★★ 言語の境目で規則を割らない ── Basic と Python の定数が同じであること。

    ★ この族は「言語の境目で切れた片配線」を **3 回**踏んでいる
      （Basic の行 09-16 / Basic の列 09-22 / Python の列 09-22）。
    """
    import re
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
    from ailine_core import table_scan
    src = _bas()
    for name in ("SCAN_COLS_MAX", "GAP_COLS"):
        m = re.search(rf"Const {name} As Integer = (\d+)", src)
        assert m, f"Basic 側に {name} が無い"
        assert int(m.group(1)) == getattr(table_scan, name), (
            f"{name} が Basic と Python で違う: "
            f"{m.group(1)} / {getattr(table_scan, name)}")


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


# ---------------------------------------------------------------------------
# ★★ 掃き出しの続き（2026-09-16 夜・同じ日に踏んだ穴）
#
# helpers/*.bas の 11 か所を TableLastRow に畳んだが、**生成側が組み立てて埋め込む**
# 走査（_scan_last_row_basic）を数えていなかった。1 本から 7 箇所へ吐いていて、
# 当たる op は 6 つ（APPEND_TOTAL / BOLD / CENTER_ALIGN / COMPUTE_COLUMN /
# FILL_COLOR / SET_COLUMN_VALUE）。実測した症状:
#
#     合計行（左端が空）のある表で「取引先の列を太字にして」
#       → 見出しとデータ行だけ太字になり、合計行に届かず × （原本は無傷だが操作できない）
#
# ★ 朝に「1 つ直して満足しない・他の形を列挙してから閉じる」と書いた当人が、
#   その日のうちに 1 つ直して満足した。だから番人を**数える側**にも置く。
# ---------------------------------------------------------------------------

def test_the_generated_basic_also_asks_the_single_place():
    """★★ 生成側が吐く走査も、畳んだ 1 本を呼ぶこと。

    ★ `.bas` の中だけを見ていると、この経路は**丸ごと視界の外**に出る
      （走査は Python の文字列として組み立てられ、helpers には存在しない）。
    """
    import inspect
    import sys as _sys
    _sys.path.insert(0, str(REPO / "src"))
    import ailine
    src = inspect.getsource(ailine._scan_last_row_basic)
    assert "TableLastRow" in src, (
        "生成側の走査が畳んだ 1 本を呼んでいない ── "
        "左端が空の行を持つ表で、その先が見えなくなります")


def test_no_op_generates_its_own_a_column_walk():
    """★ 生成関数のどれも、A 列だけを見る走査を**自前で**書いていないこと。"""
    import ast
    import sys as _sys
    _sys.path.insert(0, str(REPO / "src"))
    import ailine
    src = (REPO / "src" / "ailine" / "__init__.py").read_bytes().decode("utf-8")
    bodies = {n.name: (ast.get_source_segment(src, n) or "")
              for n in ast.walk(ast.parse(src)) if isinstance(n, ast.FunctionDef)}
    bad = []
    for op in sorted(ailine.OP_SCHEMA):
        fn = ailine.CODEGEN_BY_OP.get(op)
        body = bodies.get(getattr(fn, "__name__", ""), "")
        if _A_COLUMN_WALK.search(body):
            bad.append(op)
    assert not bad, (
        f"生成関数が A 列だけの走査を自前で書いている: {bad} ── "
        "_scan_last_row_basic（畳んだ 1 本を呼ぶ）を使うこと")


@pytest.mark.local
@pytest.mark.parametrize("task, want", [
    ("取引先の列を太字にして", "B3"),          # ★ 合計行のセルまで届くこと
    ("取引先の列を黄色で塗って", "B3"),
])
def test_a_column_operation_reaches_the_total_row(tmp_path, task, want):
    """★ 実機 ── 列を対象にする操作が、左端の空いた合計行にも届くこと。"""
    book = _book(tmp_path, [("A-001", "丸山工業", "2026-09-20", 1250000),
                             (None, None, "合計", 1250000)])
    r = _run(book, task)
    assert r.returncode == 0, r.stdout[-800:]
    assert f"  {want}:" in r.stdout, (
        f"合計行のセル {want} に届いていない\n{r.stdout[-800:]}")
