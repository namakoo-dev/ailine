# -*- coding: utf-8 -*-
"""依頼が引用した値は、解釈のどこかに出ていること（2026-09-18）。

★★ 起きたこと（Namakoo に「範囲の op が要るか」を測っている途中で出た・実測）:

    依頼: A1:C5 を「済」にして
    解釈: 操作:けい線
    → A1〜C4 の 12 セルに罫線を引き、**✓ 機械検証済み**・exit 0・原本を書き換えた

  ★ 「済」と書けと言われて罫線を引き、✓ を出した。依頼に在る値が宣言のどこにも
    無いのに、誰も鳴らなかった。

★★ なぜ既存の関所が黙ったか ── 内容語の定義:
    residue.py は内容語を「漢字 2 文字以上／カタカナ 2 文字以上／英字 2 文字以上」
    と定義している。**1 文字の値は依頼文から丸ごと消える**:

        A1:C5 を「完了」にして  → 残差=['完了'] → 鳴る
        A1:C5 を「済」にして    → 残差=[]       → ★ 黙る

  ★ そして `task_quotes_a_value` は「済」を**正しく取れていた** ── 関所がそれを
    使っていないだけ。また「器官は在るが配線が無い」形。
  ★ 検体 313 件中 **68 件が 1 文字の値**（◎ ○ × x 5）── 帳簿でいちばん普通の印。

★ 直し方は語の長さを広げることではない（消費される語が減り、誤爆する側へ倒れる）。
  **依頼者自身が引用符で括った 1 つ**を項として渡す ── 三項（依頼・宣言・実体）の形。

★ 実測した誤爆率: 本物の走行記録 16,039 件の成功回のうち引用値を持つ 1,682 件で、
  鳴ったのは **1 件（0.06%）** ＝ 上の事故そのもの。
  ★ 過去に却下された「全ての値へ広げる」案（誤爆 21%）とは別物。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

import pytest  # noqa: E402
from ailine_core import residue  # noqa: E402
import ailine  # noqa: E402

HEADS = {"取引先", "金額", "担当"}


def test_the_accident_itself_is_caught():
    """★★ 事故そのもの: 罫線を引いた解釈に『済』が無いことを名指しする。"""
    assert residue.unaccounted_quoted_value("済", "解釈: 操作:けい線", HEADS) == "済"


def test_a_one_character_value_is_not_invisible():
    """★★ 根: 1 文字だから消える、が起きないこと。

    ★ 既存の内容語の関所は『完了』で鳴り『済』で黙る ── その非対称がこの事故の根。
      ここは長さに依らないことを縛る（帳簿の印は ◎ ○ × 済 可 のように 1 文字が普通）。
    """
    for v in ("済", "◎", "○", "×", "可", "完了", "対応済み"):
        assert residue.unaccounted_quoted_value(v, "解釈: 操作:けい線", HEADS) == v, v


def test_the_old_gate_really_is_blind_to_one_character():
    """★ 陽性対照つきの前提確認: 既存の関所が本当に 1 文字を見ていないこと。

    ★ ここが将来直ったら、この試験が**赤くなって教えてくれる**（前提が変わった合図）。
      出ないことを信号にしないため、2 文字側が鳴ることも同時に確かめる。
    """
    pool = ailine._op_match_pool("DRAW_BORDERS")
    assert residue.find_unconsumed_words("A1:C5 を「済」にして", {}, pool) == []
    assert residue.find_unconsumed_words("A1:C5 を「完了」にして", {}, pool) == ["完了"]


def test_a_value_that_is_written_does_not_ring():
    """★ 恒真殺し: 解釈に値が出ている回は黙る（鳴りっぱなしの関所でも上は通る）。"""
    assert residue.unaccounted_quoted_value(
        "佐藤", "解釈: 操作:1セル書換 行:3 列:担当 値:佐藤", HEADS) is None


def test_quoting_a_column_name_is_a_target_not_a_value():
    """★★ 実測で外した家系: 引用が**列の名前**なら、それは書く値ではなく対象。

    ★ 本物の走行記録に在った形 ──「「商品」セルに色を付けて」→「対象:cell:1,1」。
      正しい動作で、外さないと ✓ が △ に落ちる（オオカミ少年になる）。
    """
    assert residue.unaccounted_quoted_value(
        "商品", "解釈: 操作:背景色 対象:cell:1,1 色:yellow", {"商品", "金額"}) is None


def test_nothing_quoted_never_rings():
    """★ 引用の無い依頼を横取りしない（「利益を計算して」で鳴らない）。"""
    assert ailine.task_quotes_a_value("みかんの利益を計算して") is None
    assert residue.unaccounted_quoted_value(None, "解釈: 操作:計算列", HEADS) is None
    assert residue.unaccounted_quoted_value("", "解釈: 操作:計算列", HEADS) is None


def test_a_value_hidden_inside_another_word_still_rings():
    """★★ 監査で「本物の穴」と確定した 2 つ目（2026-09-19 に塞いだ）。

    ★ 初版は `v in declaration` だったので、宣言の別の語に紛れて黙っていた。
      未済・決済・返済・済み は帳簿の普通語なので、現実に起きる形。
    """
    assert residue.unaccounted_quoted_value(
        "済", "解釈: 操作:一括書換 列:担当 値:未済", HEADS) == "済"
    assert residue.unaccounted_quoted_value(
        "済", "解釈: 操作:一括書換 列:済み対応 値:x", HEADS) == "済"


@pytest.mark.parametrize("decl", [
    "解釈: 操作:1セル書換 行:3 列:担当 値:済",
    "解釈: 操作:行追加 挿入位置:3 入れる値:担当=済",
    "解釈: シート:『売上』(1枚目) 操作:1セル書換 書き込む値:済 入れる位置:7行目",
])
def test_a_value_written_as_its_own_field_stays_silent(decl):
    """★★ 拾いすぎ殺し: 値が**塊として**出ている回は黙る。

    ★ `入れる値:担当=済` のように `=` の右に出る形も「出ている」── ここを落とすと
      本物の走行 1,682 件で誤爆が 0.18% → 3.51% へ膨らむ（実測して捨てた案）。
    """
    assert residue.unaccounted_quoted_value("済", decl, HEADS) is None


@pytest.mark.parametrize("value, decl", [
    ("2026/08/31", "解釈: 操作:1セル書換 対象列:締め日 書き込む値:2026/08/31 入れる位置:7行目"),
    ("{{合計:税込金額}}", "解釈: 操作:1セル書換 対象列:御中 書き込む値:{{合計:税込金額}} 入れる位置:7行目"),
    ("売上 (2026)", "解釈: シート:『売上 (2026)』(2枚目) 操作:並べ替え 対象:単価 順:降順"),
])
def test_values_that_contain_separators_are_not_torn_apart(value, decl):
    """★★ 区切りを含む値（日付・差し込み語・空白入りのシート名）を壊さない。

    ★ 「空白と : と = で割って比べる」案は、この 3 形を壊して誤爆を 20 倍にした。
      だから**割らずに境界で見る**（実測で選んだ設計）。
    """
    assert residue.unaccounted_quoted_value(value, decl, HEADS) is None


def test_every_quoted_value_is_checked_not_just_the_first():
    """★★ 監査で「本物の穴」と確定した 3 つ目（2026-09-19 に塞いだ）。

    ★ 「チェック列の『◎』を『合格』に書き換えて」で 2 つ目が宣言から落ちても黙っていた。
    ★ 1 つ目を返す task_quotes_a_value は**そのまま残す** ── あちらは「1 セルのつもりの
      依頼か」を見分ける項で、値を 1 つに決める必要がある（別の仕事）。
    ★ 実測: 引用が 2 つ以上ある依頼は 16,067 件中 4 件（0.02%）。全部を関所に掛けても
      誤爆は 1 件も増えなかった（0.18% のまま）── 増えると測れていたら入れなかった。
    """
    assert ailine.task_quotes_values("チェック列の「◎」を「合格」に書き換えて") == ["◎", "合格"]
    assert ailine.task_quotes_values("担当を「佐藤」にして") == ["佐藤"]
    assert ailine.task_quotes_values("利益を計算して") == []
    # ★ 同じ値を 2 回書いても 1 つに畳む（同じ ⚠ を 2 行出さない）
    assert ailine.task_quotes_values("「済」を「済」に") == ["済"]


def test_the_second_value_is_wired_into_the_gate():
    """★ 配線: 本体が**全部を返す側**を呼んでいること（1 つ目だけに戻っていない）。"""
    from _product_source import product_text
    i = product_text().index("unaccounted_quoted_value(")
    seg = product_text()[max(0, i - 500):i + 500]
    assert "task_quotes_values(" in seg, (
        "★ 関所が引用値を 1 つしか受け取っていない（2 つ目の穴が戻る）")


def test_the_judgement_lives_in_one_place():
    """★★ 片配線を作らない ── 判定は residue.py に 1 つだけ、本体は材料を渡すだけ。

    ★ 隣の関所（列名の側）のコメントが「判定は ailine_core/residue.py に 1 つだけ置き、
      ここは**材料を渡すだけ**」と書いている。同じ形にする。
    ★ 場所で決め打ちしない（tests/test_guard_ledger.py が鳴る）。
    """
    from _product_source import count_in_product, product_text
    # ★ 数えるのは製品全体（ailine_core/ を含む）なので、定義 1 + 呼び出し 1 ＝ 2。
    assert count_in_product("unaccounted_quoted_value(") == 2, (
        "★ 定義 1 + 呼び出し 1 になっていない（判定を書き写したか、配線が外れた）")
    # ★ 本体側で長さや文字種を見直していないこと（2 つ目の実装になる）。
    i = product_text().index("unaccounted_quoted_value(")
    seg = product_text()[i - 400:i + 400]
    assert "len(" not in seg, "★ 本体が値の長さを見ている（判定が 2 つに割れた）"


def test_the_gate_is_wired_where_the_verdict_is_decided():
    """★★ 鳴っても ✓ のままなら意味が無い ── 警告数に足す所に居ること。

    ★ 今日の台帳の直しで、warning_count > 0 は verdict を warned（△）にする。
      ここが外れると「⚠ を出しながら ✓」という、まさに直したばかりの嘘が戻る。
    """
    from _product_source import product_text
    i = product_text().index("unaccounted_quoted_value(")
    seg = product_text()[i:i + 700]
    assert "warning_count += 1" in seg, (
        "★ 鳴った回が警告数に足されていない（✓ が下がらない）")
