"""グラフ段①(extract_chart_kind_from_task)の自分の境界検体 + verify_dsl_args(CHART) の
   kind/category_col 解決の単体検体（凍結済み tests/test_chart_kinds.py の兄弟・
   凍結対象ではない ── 自由に書き足してよい）。

★ 断片誤爆に注意（brief 指示）: 「円」は通貨表記（「500円」）と、「棒」は複合語
  （「相棒」）と衝突しうる。extract_chart_kind_from_task 側の断片ガード
  （_CHART_KIND_YEN_GUARD の数字近傍除外・「棒」でなく「棒グラフ」全体を語にする）を
  ここで実測する。
"""
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
import ailine  # noqa: E402


# --- 自分の境界検体（断片誤爆ガード）--------------------------------------------

def test_extract_chart_kind_ignores_yen_as_currency_amount():
    """「500円」の「円」は通貨表記であってグラフ種別の指定ではない。
       明示の「棒グラフ」だけが拾われて bar になる（円につられて pie にならない）。"""
    assert ailine.extract_chart_kind_from_task(
        "商品ごとの単価(500円)を棒グラフにして") == "bar"


def test_extract_chart_kind_ignores_bar_fragment_inside_compound_word():
    """「相棒」に含まれる「棒」はグラフ種別の指定ではない（「棒グラフ」全体でないと拾わない）。
       他に種別の手掛かりが無ければ None（機械は断定しない）。"""
    assert ailine.extract_chart_kind_from_task(
        "相棒と一緒に売上のグラフ化を進めて") is None


# --- verify_dsl_args(CHART): kind/category_col の解決 --------------------------

_META = {"sheets": ["Sheet"], "headers": {"Sheet": ["商品", "金額", "在庫"]},
         "header_rows": {"Sheet": 1}}


def test_verify_dsl_args_chart_kind_defaults_to_bar_when_silent():
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "CHART", {"value_col": "金額"}, _META, task="金額をグラフにして")
    assert ok, err
    assert resolved["kind"] == "bar"
    assert not resolved.get("_warnings")


def test_verify_dsl_args_chart_mechanical_kind_beats_llm_with_disclosure():
    """LLM が bar を返しても、依頼文の「折れ線」が line に上書きし、食い違いを開示する
       （EXTRACT の cmp と同じ作法）。"""
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "CHART", {"value_col": "金額", "kind": "bar"}, _META,
        task="売上の推移を折れ線で見せて")
    assert ok, err
    assert resolved["kind"] == "line", f"機械抽出が LLM に負けた: {resolved}"
    assert resolved.get("_warnings"), "食い違いの開示が無い"


def test_verify_dsl_args_chart_kind_stands_when_task_silent_and_agrees_without_noise():
    """依頼文に種別語が無ければ LLM の kind を従来どおり使う（開示なし・オオカミ少年防止）。"""
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "CHART", {"value_col": "金額", "kind": "pie"}, _META, task="金額をグラフにして")
    assert ok and resolved["kind"] == "pie"
    assert not resolved.get("_warnings"), f"一致なのに開示が出た: {resolved.get('_warnings')}"


def test_verify_dsl_args_chart_rejects_unknown_kind():
    ok, _resolved, _inf, err = ailine.verify_dsl_args(
        "CHART", {"value_col": "金額", "kind": "scatter"}, _META, task="金額をグラフにして")
    assert not ok
    assert "scatter" in err


def test_verify_dsl_args_chart_category_col_defaults_to_first_column():
    ok, resolved, inferred, err = ailine.verify_dsl_args(
        "CHART", {"value_col": "金額"}, _META, task="金額をグラフにして")
    assert ok, err
    assert resolved["category_col"] == "商品"
    assert "category_col" in inferred


def test_verify_dsl_args_chart_category_col_explicit_is_validated():
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "CHART", {"value_col": "金額", "category_col": "在庫"}, _META,
        task="在庫ごとの金額をグラフにして")
    assert ok, err
    assert resolved["category_col"] == "在庫"


def test_verify_dsl_args_chart_category_col_unknown_column_fails():
    ok, _resolved, _inf, err = ailine.verify_dsl_args(
        "CHART", {"value_col": "金額", "category_col": "存在しない列"}, _META,
        task="金額をグラフにして")
    assert not ok
    assert "存在しない列" in err


def test_category_mention_is_consumed_not_flagged(tmp_path, monkeypatch, capsys):
    """★ 検分の実機再現（2026-08-23）: 「商品ごとの構成比を円グラフにして」──『商品』は
       横軸列（category_col）に正しく吸われたのに、category_col が OP_SUBJECT_SLOTS に
       無いため ③ 誤爆（『売上』は依頼文と照合できません）で ✓ が消えた。
       operator8 ①（LOOKUP_FILL の参照シート）と同じ形 ── 言及は消費されるスロット
       （SUBJ_INPUT）として登録し、③ を出さず ✓ まで通す。"""
    import sys as _sys, io
    from test_golden_transcripts import _book, _isolate, _run_main
    import openpyxl
    _isolate(monkeypatch, tmp_path)
    # 治具: 中身は fixture（実 LO 産 pie.xlsx）と同一の表にする ── fake_apply が fixture を
    #   丸コピーするため、列名/値が違うと事後条件が別の理由で落ちて ③ の検査にならない
    book = _book(tmp_path, [["商品", "金額"], ["りんご", 100], ["みかん", 200], ["ぶどう", 150]])
    monkeypatch.setattr(
        ailine, "translate_task",
        lambda model, task, book_meta, temperature=0.1:
        {"op": "CHART", "args": {"value_col": "金額", "kind": "pie"}})

    def fake_apply(out_book, code, workdir, helper_files=(), timeout=None):
        # 実 LO 産の pie fixture から chart XML ごとコピーして「円グラフが刺さった」状態を作る
        import shutil
        shutil.copy2(Path(__file__).parent / "fixtures" / "charts" / "pie.xlsx", out_book)
        return True, None, "ok"
    monkeypatch.setattr(ailine, "basrun_apply", fake_apply)
    rc, out = _run_main(["run", str(book), "商品ごとの構成比を円グラフにして", "--copy"], capsys)
    assert "機械照合できません" not in out, f"category の言及が ③ 誤爆している: {out}"
    assert rc == 0, out
