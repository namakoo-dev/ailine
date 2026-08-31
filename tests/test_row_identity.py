# 計算で出している列が、直した値に付いていっているか ── 2026-08-31。
# Namakoo「金額が入れ替われば付随して関連するセルの内容も変えなければいけない。
# しかもそれが複数の内容に影響する場合はそれらも踏まえて変更しないといけない」
#
# ★★ 実測: 「丸和物流の単価とみどり建設の単価を入れ替えて」は**頼まれた 2 セルだけ**を
#   正しく入れ替える。だが 金額（＝件数×単価）は**直値**なので取り残され、
#       件数 12 × 単価 7200 = 86,400  なのに 金額 57,600 のまま
#   という**表として矛盾した状態**になる。それでも「2 セルだけ動いた」は真実なので ✓ が出た。
#
# ★ 式で書かれていれば LibreOffice が再計算するので起きない。**直値の派生列**だけの事故。
#   見た目は普通の数字なので、人は気づけない ── この道具が一番嫌う形。
# ★ 直さない・**言う**（どう直すかは人が決める ── 参照のズレと同じ線）。
# ★ 語も見出しも読まない ── **数だけ**を見る（sum_identity と同じ性質）。
# ★ op を問わず 5 つの助言の口すべてに同じ関数を配る（片配線を作らない）。

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import row_identity as ri  # noqa: E402

HEADERS = ["取引先", "件数", "単価", "金額"]
ROWS = [["丸和物流", 12, 4800, 57600],
        ["近江スチール", 5, 12000, 60000],
        ["ヤマノ食品", 28, 1500, 42000],
        ["みどり建設", 9, 7200, 64800]]


def test_a_product_identity_is_found():
    got = ri.identities(ROWS)
    assert (3, "×", 1, 2) in got, got


def test_the_commutative_twin_is_folded():
    """★ 掛け算は交換法則で 2 回当たる ── 人には同じことに見えるので 1 本にする。"""
    got = [x for x in ri.identities(ROWS) if x[0] == 3 and x[1] == "×"]
    assert len(got) == 1, got


def test_a_row_with_missing_numbers_is_ignored():
    """★★ 最初の実装はここで空振りした ── **合計行**（件数・単価が空）のせいで
       等式が全部捨てられていた。数が揃っていない行は無視する。"""
    with_total = ROWS + [["合計", None, None, 224400]]
    assert (3, "×", 1, 2) in ri.identities(with_total)


def test_swapping_an_input_breaks_it():
    after = [r[:] for r in ROWS]
    after[0][2], after[3][2] = after[3][2], after[0][2]     # 単価だけ入れ替え
    assert (3, "×", 1, 2) in ri.broken(ROWS, after)


def test_moving_whole_rows_does_not_break_it():
    """★ 鳴りすぎない: 行ごと入れ替えれば等式は保たれる ── 黙る。"""
    after = [ROWS[3], ROWS[1], ROWS[2], ROWS[0]]
    assert ri.broken(ROWS, after) == []


def test_changing_row_count_is_not_our_business():
    """★ 行を足す/消す操作で必ず鳴るのを避ける（別の番人の担当）。"""
    assert ri.broken(ROWS, ROWS + [["新規", 1, 100, 100]]) == []


def test_too_few_rows_decides_nothing():
    """★ 偶然の一致を等式と呼ばない。"""
    assert ri.identities(ROWS[:2]) == []


def test_the_note_names_the_columns_in_japanese():
    note = ri.describe([(3, "×", 1, 2)], HEADERS)
    assert "『金額』＝『件数』×『単価』" in note, note
    assert "直していません" in note


def test_the_advisory_is_wired_to_every_declared_path():
    """★ **宣言（op）が在る**助言の口すべてに配っていること（片配線を作らない）。

    ★★ FREEFORM の段だけは外してある ── あそこは op が決まっていないので
      `resolved`（宣言）が無い。**比べる相手が無いところでは黙る**のが正しい。
      （最初は 5 箇所すべてに配って NameError を 2 回出した。呼び出し側の変数名に
       頼る書き方をやめ、渡すものを resolved 1 つに減らして直した。）
    """
    src = (REPO / "src" / "ailine" / "__init__.py").read_text(encoding="utf-8")
    assert src.count("broken_identity_advisory(") == 5, "定義 1 + 呼び出し 4 のはず"
    assert "broken_identity_advisory(stepsource, out_book, resolved" not in src or         src.count("broken_identity_advisory(stepsource") == 1, "FREEFORM 段に配っている"


@pytest.mark.local
def test_it_fires_on_a_real_cell_swap(tmp_path):
    """★★ Namakoo の実例そのもの ── ✓ でなく △ になること。"""
    import subprocess

    import openpyxl
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r)
    wb.save(p)
    env = {**__import__("os").environ, "PYTHONPATH": str(REPO / "src")}
    r = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(p),
         "丸和物流の単価とみどり建設の単価を入れ替えて", "--copy", "--sheet", "請求"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=900, cwd=str(REPO), env=env)
    assert "『金額』＝『件数』×『単価』" in r.stdout, r.stdout[-1500:]
    assert "△" in r.stdout and "✓ " not in r.stdout.split("△")[0][-200:], r.stdout[-800:]
