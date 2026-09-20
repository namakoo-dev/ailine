# -*- coding: utf-8 -*-
"""転記: キー列と対象列が同じなら、書く前に断って原因を名指しする（2026-09-11）。

★★ 実測した事故（docs/PENDING-20260910-グラフの項目列と転記のキー列.md ②）:

    依頼   「商品表から商品名を注文シートに転記して」
    解釈   操作:転記 対象列:商品名 参照シート:商品表 キー列:商品名   ★ キー列＝対象列
    生成   VLookupFromTable(oDoc, 0, 2, 2, "商品表")                 ★ 2 と 2
    結果   空の列を空の列で引くので 1 件も一致しない → 文書に変化なし
           → 事後条件「検証対象が 0 件」が × を出す（正しいが、なぜかを教えない）

  battery の名前の台帳に `注文/lookup` として載っていた家系。

★ 判定は**構造だけ**（依頼文も語彙も読まない・detect_new_row_missing_key と同じ線）:
  見るのは「キー列と対象列の名前が同じか」だけ。
★ 数字（× → 断り）が良くなる向きの変更なので、決め手は「親切か」に置く ──
  断りは**原因**と**次に言うこと**を名指しすること（下で縛る）。

★ ここは純ロジックで測る（LLM も LibreOffice も要らない）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import ailine  # noqa: E402

#: battery の「注文」表そのもの（コード/数量/商品名 ＋ 参照表 商品表=コード/商品名）
META = {
    "sheets": ["注文", "商品表"],
    "headers": {"注文": ["コード", "数量", "商品名"], "商品表": ["コード", "商品名"]},
    "header_rows": {"注文": 1, "商品表": 1},
}
TASK = "商品表から商品名を注文シートに転記して"


def _verify(args: dict):
    return ailine.verify_dsl_args("LOOKUP_FILL", dict(args), META, task=TASK,
                                  target_sheet="注文")


def test_key_equal_to_target_is_refused_before_writing():
    """★ 7B が返した通りの引数（キー列＝対象列）は、書く前に断る。"""
    ok, _resolved, _inferred, err = _verify(
        {"target_sheet": "注文", "target_col": "商品名", "source_sheet": "商品表", "key_col": "商品名"})
    assert ok is False, "★ キー列＝対象列なのに通した（空の列を空の列で引く）"
    assert "商品名" in err, f"★ どの列が重なっているかを名指していない: {err}"


def test_the_refusal_names_the_cause_and_the_next_thing_to_say():
    """★★ 断りは「検証対象が 0 件」より親切でなければ入れる意味が無い。

    ★ 原因（キー列と対象列が同じ）と、次に言うこと（参照表の 1 列目＝コード で引く）を
      両方名指しすること。参照表の 1 列目は VLookupFromTable の**契約**であって推測ではない。
    """
    ok, _r, _i, err = _verify(
        {"target_sheet": "注文", "target_col": "商品名", "source_sheet": "商品表", "key_col": "商品名"})
    assert ok is False
    assert "キー列" in err and "対象列" in err, f"★ 原因を言っていない: {err}"
    assert "コード" in err, f"★ 次に言うべき列（参照表の 1 列目）を名指していない: {err}"
    assert "依頼文" in err, f"★ 人が何をすればいいか言っていない: {err}"
    # ★★ 2026-09-20（盲検 5 体目 ⑤）: 例文が**その誤りを再生産しない**こと。
    #   買い手:「その例文どおりに書いたら、また同じエラーになるはずです。
    #   直し方が書いていないのと同じでした」── 導通の盤で一番重い `path_fails`。
    #   ★ 俺は直しで**2 回続けて同じ所を外した**（1 度目は句が隣り合っていない前提、
    #     2 度目は除外する列を間違えた）。だから字面でなく**形**で縛る:
    #     例文の「◯◯で引いて△△を転記して」の ◯◯ と △△ が同じなら赤。
    import re as _re
    for _m in _re.finditer(r"「([^」]+?)で引いて([^」]+?)を転記して」", err):
        assert _m.group(1) != _m.group(2), (
            f"★ 例文が同じ列を両側に置いている（打つと同じ誤りが再生産される）: {_m.group(0)}")


def test_a_proper_key_still_passes():
    """★ 負の被覆 ── 正しいキー列（コード）は今までどおり通る（断りが広がりすぎない）。"""
    ok, resolved, _i, err = _verify(
        {"target_sheet": "注文", "target_col": "商品名", "source_sheet": "商品表", "key_col": "コード"})
    assert ok is True, f"★ 正しい転記まで断った: {err}"
    assert resolved["key_col"] == "コード" and resolved["target_col"] == "商品名"


def test_the_judgment_reads_structure_not_the_request_text():
    """★ 依頼文に何が書いてあっても、判定は列名の一致だけで決まる（語彙を読まない）。"""
    for task in ("商品名を埋めて", "転記して", ""):
        ok, _r, _i, err = ailine.verify_dsl_args(
            "LOOKUP_FILL",
            {"target_sheet": "注文", "target_col": "商品名", "source_sheet": "商品表", "key_col": "商品名"},
            META, task=task, target_sheet="注文")
        assert ok is False, f"★ 依頼文『{task}』で判定が変わった"

#: ★★ 盲検 5 体目の形（2026-09-20）── 参照表の **1 列目が転記先と同じ名前**。
#:   買い手の発注記録がこれで、例文が「商品名で引いて商品名を転記して」になった
#:   ＝ **打つと同じ誤りが再生産される**（導通の盤で一番重い `path_fails`）。
#:   ★ 検体を足した理由: 既存の検体は 1 列目が『コード』で転記先と違うため、
#:     直しを戻す変異が**素通り**した（変異試験が指した）。
BUYER_META = {
    "sheets": ["請求", "商品名・発注金額だけ"],
    "headers": {"請求": ["商品名", "請求金額", "発注金額"],
                "商品名・発注金額だけ": ["商品名", "発注金額"]},
    "header_rows": {"請求": 1, "商品名・発注金額だけ": 1},
}


def test_the_example_never_repeats_the_same_column_in_the_buyers_shape():
    """★★ 参照表の 1 列目が転記先と同名でも、例文が誤りを再生産しないこと。

    ★ 盲検 5 体目の買い手が踏んだ形そのもの。
    """
    ok, _r, _i, err = ailine.verify_dsl_args(
        "LOOKUP_FILL",
        {"target_sheet": "請求", "target_col": "商品名",
         "source_sheet": "商品名・発注金額だけ", "key_col": "商品名"},
        BUYER_META, task="商品名で発注金額を転記して", target_sheet="請求")
    assert ok is False, err
    import re
    pairs = re.findall(r"「([^」]+?)で引いて([^」]+?)を転記して」", err)
    assert pairs, f"★ 例文が無い（通る道を示していない）: {err}"
    for puller, target in pairs:
        assert puller != target, (
            f"★ 例文が同じ列を両側に置いている（打つと同じ誤りが再生産される）: {err}")
    assert "発注金額" in err, f"★ 転記できる列を名指していない: {err}"
