# -*- coding: utf-8 -*-
"""断りは「決まらない」でなく**なぜ選べないか**を言う（2026-09-17・盲検 3 体目）。

★★ 起きたこと（製造業の購買担当・初見・盲検）:

    $ ailine run 発注台帳.xlsx 標準単価表.xlsx "品目をキーにして金額を突き合わせて"
    ？ 発注台帳.xlsx のキー列が依頼文から決まりません。候補: 品目、金額。
      依頼文に列名を含めて（例:『品目をキーに』）もう一度実行してください。
    ？ 発注台帳.xlsx の金額列が依頼文から決まりません。候補: 数量、単価。

  ★ 買い手は**既に「品目をキーに」と書いていた** ── 道具が勧めてきた例文と一字一句同じ。
  ★ そして「**なぜ金額が選べないのか**が最後まで分かりませんでした」。
    購買が突き合わせたい第一候補は金額なのに、候補に一度も現れない。

★★ 追ったら **1 つの原因で 2 つの症状**だった:
  『金額』は `=E2*F2` のままで計算結果が保存されていない → 照合の読み（data_only=True・
  計算結果が要るので正しい）では**空のセル**に見える → 数値の列と認められない →
  ① 金額の候補から外れる ② 代わりに**キーの候補に混ざる**（非数値なので）→
  「品目と金額のどちらがキーか決まらない」。
  ★ 道具は理由を**持っていたのに言っていなかった**（今日の①と同じ形）。

★★ 直すのは断る条件ではなく**断り方**。緩める方向（式の列を数値とみなす）は採らない ──
  計算結果の無いものを金額として足したら、それこそ「自信のある嘘の数字」を作ることになる
  （今日その形を照合の差額で直したばかり）。

★ 測って分かったこと（検体の側の話・記録として残す）: この詰まりは**検体の作りに依存する**。
  openpyxl で書いた冊は式にキャッシュ値を持たないので起きるが、Excel か LibreOffice が
  保存した冊なら値が入っていて素直に通る（実測: 計算値を入れた同じ表は key=品目・
  amount=金額 で解決）。実害が当たるのは**他の道具が書き出した冊**。
"""
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import match  # noqa: E402

HEADERS = ["発注番号", "仕入先", "品目", "数量", "単価", "金額"]


def _book(tmp_path, *, amount_cached: bool):
    p = tmp_path / ("cached.xlsx" if amount_cached else "formula.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(HEADERS)
    for i, (no, sup, item, q, price) in enumerate(
            [("P1", "大東金属", "SUS304", 40, 1850),
             ("P2", "高橋工機", "ボールねじ", 4, 28500)], start=2):
        ws.append([no, sup, item, q, price, (q * price) if amount_cached else f"=D{i}*E{i}"])
    wb.save(p)
    return p


def test_a_formula_column_without_values_is_found(tmp_path):
    """★ 事故の実体: 計算結果を持たない式の列を、名指しで挙げられること。"""
    p = _book(tmp_path, amount_cached=False)
    assert ailine._formula_columns_without_values(p, 1, HEADERS) == {"金額"}


def test_a_column_with_real_numbers_is_not_accused(tmp_path):
    """★★ 陰性対照: 値が入っている列を「式のまま」と言わない。

    ★ ここが無いと「数値列を全部そう呼ぶ」実装でも上の試験が通る ── それは
      通る回にまで嘘の注記を出す形。
    """
    p = _book(tmp_path, amount_cached=True)
    assert ailine._formula_columns_without_values(p, 1, HEADERS) == set()


def test_an_unreadable_book_is_quiet(tmp_path):
    """★ 読めない時は黙る（理由が言えないことを理由にして落ちない）。"""
    bad = tmp_path / "x.xlsx"
    bad.write_bytes(b"not a workbook")
    assert ailine._formula_columns_without_values(bad, 1, HEADERS) == set()


def test_the_same_table_resolves_once_the_values_are_there(tmp_path):
    """★★ 詰まりの原因が『式のまま』であることの裏取り ── 値が入れば素直に通る。

    ★ これが無いと「断り方を直した」だけで、**そもそも断るのが正しいのか**を
      確かめていないことになる（緩める方向へ倒さない根拠でもある）。
    """
    def rows_of(p):
        wb = openpyxl.load_workbook(p, data_only=True)
        return match.read_data_rows(wb.worksheets[0], 1, HEADERS)

    hb = ["仕入先", "品目", "標準単価"]
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.append(hb)
    ws2.append(["大東金属", "SUS304", 1850])
    ws2.append(["高橋工機", "ボールねじ", 28500])
    pb = tmp_path / "b.xlsx"
    wb2.save(pb)
    rows_b = rows_of(pb)
    task = "品目をキーにして金額を突き合わせて"

    bad = match.resolve_columns(task, HEADERS, rows_of(_book(tmp_path, amount_cached=False)),
                                hb, rows_b)
    good = match.resolve_columns(task, HEADERS, rows_of(_book(tmp_path, amount_cached=True)),
                                 hb, rows_b)
    assert not bad.ok, "★ 式のままの冊が通ってしまう（計算結果の無い値を足すことになる）"
    assert good.ok and good.key_a == "品目" and good.amount_a == "金額", (
        f"値が入っていても解決しない: {good}")


def test_the_reason_is_printed_outside_the_candidate_branch():
    """★★ 理由は候補の有無に関わらず言う ── if/else の**中**に入れたら片配線。

    ★ 2026-09-17 に実際に踏んだ: 最初の当て方では元の else が新しい if にくっつき、
      「使える列が見つかりません」が余計に出ていた。
    """
    from _product_source import count_in_product
    assert count_in_product("if _named_formula:") == 1
    assert count_in_product('say(f"？ {book_label} に{label}に使える列が見つかりません。")') == 1


def test_the_detector_lives_where_the_formulas_can_be_read():
    """★★ 器官を置く場所を間違えない ── 照合の読みは data_only=True で、式は見えない。

    ★ 2026-09-17 に実際に踏んだ: 検出器を match.py（行の値だけを見る側）に置いたら、
      式の列は**丸ごと空**に見えて 1 件も拾えなかった。冊を読み直す側に置く。
    ★ 効かなくなった器官は残さない（死んだ名前は次に読む人を騙す）。
    """
    from _product_source import count_in_product
    assert count_in_product("def _formula_columns_without_values(") == 1
    src = (REPO / "src" / "ailine_core" / "match.py").read_bytes().decode("utf-8")
    assert "formula_only_columns" not in src, (
        "★ 効かなくなった器官が match.py に残っている")


@pytest.mark.parametrize("role,label", [("key", "キー"), ("amount", "金額")])
def test_both_roles_can_carry_the_reason(role, label):
    """★ 理由はキー側・金額側の**どちらの断りでも**出せること（片方だけ直さない）。

    ★★ 2026-09-20: ここは**文面まるごと**（末尾の「Excel か LibreOffice で一度開いて
      保存すると値が入ります」を含む）で縛っていたので、助言の側を変えた日に赤くなった
      ── 道具が自分で値を入れるようになり、**既に試した回は同じことを人に頼まない**
      分岐を足したため。守っている不変（役割ごとに書き写さない）は変わっていない。
      ★ 今週 7 件目の「番人を字面で書いた」形。
    ★ だから**核の主張**に結び直す ──「式のままで計算結果が入っていない」は人が読む
      理由そのもので、ここが 2 箇所になったら片方が腐る。末尾の助言は場面で変わってよい
      （`_how` が 1 箇所で選ぶ）。
    """
    from _product_source import count_in_product
    assert count_in_product("式のままで計算結果が入っていない") == 1, (
        "理由の文が 1 箇所に畳まれていない（役割ごとに書き写すと片方が腐る）")
    assert count_in_product("{label}の列として使えません") == 1, (
        "役割の語を差し込む所が 2 箇所ある（キー側と金額側で別文になる）")
    assert role in ("key", "amount") and label in ("キー", "金額")
