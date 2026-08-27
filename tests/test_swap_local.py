"""SWAP の実機（LO・basrun）検体 ── 2026-08-27。

★★ なぜ実機の検体が要るか（今日この形で 3 時間払った）:
  生成した Basic が**コンパイルに失敗しても、basrun は「適用した」と言う**。
  マクロは 1 行も走らず、エラーも 1 行も出ない。sandbox の pytest は basrun をモック
  するので、この壊れ方は構造的に捕まえられない ── 実機でしか分からない。
  ★ 実際の犯人: 変数名 `oR`。Basic は大文字小文字を区別しないので **予約語 Or** と
    衝突してモジュールごと落ちていた。名前を変えるだけで直った。
  ★ 命綱は効いた: 事後条件が「変化なし」で × を出したので、嘘の ✓ にはならなかった。

★ もう 1 つの実測: **ヘルパの module を書き換えた直後の 1 回目は、古い module が走る**。
  切り分け中の「通ったり通らなかったり」はこれだった（同じ入力で 2 回目以降は安定）。

★ タイムアウトは短く（楔でスイート全体を道連れにしない）。
"""
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent


def _book(tmp_path, name="b.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["商品", "売上", "原価", "利益"])
    for i, (n, a, b) in enumerate([("りんご", 1200, 700), ("みかん", 800, 300),
                                    ("ぶどう", 1500, 900)], start=2):
        ws.cell(i, 1, n), ws.cell(i, 2, a), ws.cell(i, 3, b)
        ws.cell(i, 4, f"=B{i}-C{i}")
    wb.save(p)
    return p


def _run(book, task):
    return subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(book), task, "--copy", "--timeout", "90"],
        capture_output=True, text=True, timeout=420, encoding="utf-8", errors="replace")


def _out(book):
    return book.with_name(book.stem + ".out" + book.suffix)


@pytest.mark.local
def test_rows_really_swap_on_real_lo_and_formulas_follow(tmp_path):
    book = _book(tmp_path)
    p = _run(book, "みかんとぶどうを入れ替えて")
    assert p.returncode == 0, f"実機の行入れ替えが失敗:\n{p.stdout[-900:]}"
    raw = openpyxl.load_workbook(_out(book))["売上"]
    val = openpyxl.load_workbook(_out(book), data_only=True)["売上"]
    names = [raw.cell(r, 1).value for r in range(2, 5)]
    assert names == ["りんご", "ぶどう", "みかん"], f"入れ替わっていない: {names}"
    # ★ 芯: 式は**自分の行**を指し続ける（値を交換する実装だと、ここが他の行の値になる）
    assert val.cell(3, 4).value == 600, f"ぶどうの利益が自分の値でない: {val.cell(3, 4).value}"
    assert val.cell(4, 4).value == 500, f"みかんの利益が自分の値でない: {val.cell(4, 4).value}"


@pytest.mark.local
def test_columns_really_swap_on_real_lo_and_references_follow(tmp_path):
    book = _book(tmp_path, "c.xlsx")
    p = _run(book, "売上と原価の列を入れ替えて")
    assert p.returncode == 0, f"実機の列入れ替えが失敗:\n{p.stdout[-900:]}"
    raw = openpyxl.load_workbook(_out(book))["売上"]
    val = openpyxl.load_workbook(_out(book), data_only=True)["売上"]
    assert [raw.cell(1, c).value for c in range(1, 5)] == ["商品", "原価", "売上", "利益"]
    # 参照は LibreOffice が付け替える（=B2-C2 → =C2-B2）。**計算結果は変わらない**のが正しい。
    assert str(raw.cell(2, 4).value).replace(" ", "") == "=C2-B2", raw.cell(2, 4).value
    assert val.cell(2, 4).value == 500, val.cell(2, 4).value


@pytest.mark.local
def test_the_generated_basic_actually_runs(tmp_path):
    """★ 恒真殺し: 「適用した」と言われても走っていない、という今日の壊れ方を直接見る。
       入れ替えを頼んで**原本と 1 バイトも違わない**なら、それはマクロが走っていない。"""
    book = _book(tmp_path, "d.xlsx")
    before = book.read_bytes()
    p = _run(book, "みかんとぶどうを入れ替えて")
    assert p.returncode == 0, p.stdout[-600:]
    assert _out(book).read_bytes() != before, (
        "出力が原本と同一 ── マクロが 1 行も走っていない疑い"
        "（Basic のコンパイル失敗は basrun からは『適用した』に見える）")


# --- 列の追加（2026-08-27・Namakoo「列の追加はできないの？」）------------------------
#
# ★ 一段目は**同じ依頼文で回ごとに違う op** を返す（実測）。「原価列の右に列を追加して」で
#   INSERT_ROWS が返る回があり、そのまま走れば**列を頼まれて行を挿す**。
#   軸そのものを間違える形なので、実機で通しておく。

@pytest.mark.local
def test_a_column_really_gets_inserted_at_the_named_position(tmp_path):
    book = _book(tmp_path, "e.xlsx")
    p = _run(book, "原価の右に備考の列を追加して")
    assert p.returncode == 0, f"実機の列追加が失敗:\n{p.stdout[-900:]}"
    raw = openpyxl.load_workbook(_out(book))["売上"]
    val = openpyxl.load_workbook(_out(book), data_only=True)["売上"]
    assert [raw.cell(1, c).value for c in range(1, 6)] == ["商品", "売上", "原価", "備考", "利益"]
    # 利益は D → E へずれるが、参照する 売上/原価 は動いていないので式も値もそのまま
    assert val.cell(2, 5).value == 500, val.cell(2, 5).value
    assert raw.cell(2, 4).value is None, "挿した列に値が入っている"


@pytest.mark.local
def test_asking_for_a_column_never_inserts_a_row(tmp_path):
    """★ 軸の取り違えを実機で縛る ── 行数が増えたら、それは列の依頼に行で答えている。"""
    book = _book(tmp_path, "f.xlsx")
    p = _run(book, "原価列の右に列を追加して")
    assert p.returncode == 0, p.stdout[-600:]
    ws = openpyxl.load_workbook(_out(book))["売上"]
    assert ws.max_row == 4, f"行が増えた（列の依頼に行で答えた）: {ws.max_row}"
    assert ws.max_column == 5, f"列が増えていない: {ws.max_column}"
