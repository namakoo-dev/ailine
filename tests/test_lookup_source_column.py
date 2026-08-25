# 塊③(2/2)・中核 op 致命2（2026-08-24 の盲検）── 転記が参照表の「B 列」を決め打ち。
#
# ★ 実測: マスタ = 商品 / 区分 / **単価**（C 列）に対して「単価を転記して」と頼むと
#     変更点: C2: (空)→'果物'   C3: (空)→'果物'   C4: (空)→'高級'
#     ✓ 機械検証済みの内容です
#   **単価の列に「果物/果物/高級」が入った。** 数値であるべき列に文字列が入って ✓。
#
# ★ 根: 書き手（helpers の VLookupFromTable）が `oLook.getCellByPosition(1, j)`＝
#   「参照表 列1=値」を決め打ちし、検算（check_lookup_fill）も同じく列1・列2 決め打ちで
#   期待値を作る。**やる側と見る側が同じ思い込みを共有している**ので必ず一致する ── 恒真。
#   `verify_dsl_args` は「対象列名が依頼文にあるか」しか見ず、
#   **マスタ側の値列が本当にその名前かを一度も照合しない**。
#   ★ 商品コード / 商品名 / 単価 のような 3 列マスタは実務でごく普通で、いつでも当たる。
#
# 契約:
#   ① 参照表の見出しを読み、頼まれた列名が**2 列目**でなければ書く前に断る
#   ② 断る時は「何列目に在るか」を名指しする（人が直せるように）
#   ③ 2 列マスタ（キー・値）は従来どおり通る（誤爆しない）
#   ④ 見出しが読めない参照表は従来どおり（黙って通す・断りの根拠が無い）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402


def _meta(sheets, headers):
    return {"sheets": sheets, "headers": headers,
            "header_rows": {s: 1 for s in sheets}}


def test_three_column_master_is_disclosed(tmp_path):
    """①② 単価が 3 列目なのに「単価を転記して」── 何が書かれるかを名指しする。

    ★ 断らない理由（実測で 1 度誤爆した）:
        事故の形   マスタ=[商品,区分,単価] → 2 列目は「区分」
        正しい依頼 明細  =[商品,数量,単価] → 2 列目は「数量」
      どちらも「2 列目 ≠ 頼まれた列」で、**列の位置だけでは区別できない**。
      断ると正しい依頼まで止める（既存検体 test_verify_dsl_args_lookup_fill_allows_
      non_first_sheet_target で実証された）。判定は変えず、開示して ✓ を降ろす。
    """
    meta = _meta(["明細", "マスタ"],
                 {"明細": ["商品", "単価"], "マスタ": ["商品", "区分", "単価"]})
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "LOOKUP_FILL",
        {"target_sheet": "明細", "target_col": "単価",
         "key_col": "商品", "source_sheet": "マスタ"},
        meta, task="マスタから単価を転記して")
    assert ok, f"開示で足りるのに断った（正しい依頼まで止まる）: {err}"
    warns = resolved.get("_warnings") or []
    assert warns, "単価の列に区分が入るのに黙った"
    text = " ".join(warns)
    assert "区分" in text and "3 列目" in text, text


def test_two_column_master_still_works(tmp_path):
    """③ 誤爆しない: キー・値の 2 列なら従来どおり。"""
    meta = _meta(["明細", "マスタ"],
                 {"明細": ["商品", "単価"], "マスタ": ["商品", "単価"]})
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "LOOKUP_FILL",
        {"target_sheet": "明細", "target_col": "単価",
         "key_col": "商品", "source_sheet": "マスタ"},
        meta, task="マスタから単価を転記して")
    assert ok, f"正しい 2 列マスタを落とした: {err}"


def test_unknown_master_headers_are_not_refused():
    """④ 見出しが読めない参照表は断らない（根拠が無い時に止めない）。"""
    meta = _meta(["明細", "マスタ"], {"明細": ["商品", "単価"]})
    ok, resolved, _inf, err = ailine.verify_dsl_args(
        "LOOKUP_FILL",
        {"target_sheet": "明細", "target_col": "単価",
         "key_col": "商品", "source_sheet": "マスタ"},
        meta, task="マスタから単価を転記して")
    assert ok, f"根拠が無いのに断った: {err}"
