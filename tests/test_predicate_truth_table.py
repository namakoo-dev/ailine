"""EXTRACT 述語の真理値表（E5・architect M2 レビューの作法: ★ 実装より先に手書きで凍結）。

★ なぜ在るか: M2（フォルダ抽出）では選別も検算も Python になり、独立性は
「ailine.py の _extract_predicate」と「ailine_core 側の再実装」という 2 実装から調達する。
2 実装が同じ間違いを共有したら独立の意味が無い ── この表が共有の校正原器になる。
★ 表は仕様（型の保存の哲学）から手で書いた。実装から生成してはならない（測定器を
検体から作らない）。実装と表が食い違ったら: 実装のバグか、仕様の書き直し（理由を
ここに追記）のどちらかを人が決める。

凍結した意味論（2026-08-21・DESIGN-20260821-multifile M2 節）:
- 数値比較 (gte/lte/gt/lt) は 本物の数値（int/float・bool は除く）にだけ効く。
  文字列数値 "40000"・日付・None・bool は不一致（黙って型変換しない）
- eq: 両辺が数値なら許容誤差 1e-6 で比較。値が数値でなければ文字列の完全一致
- contains: 文字列セルのみ。数値セル・None は不一致
"""
import datetime
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402

# (cmp, 条件値, セル値, 期待)
TRUTH = [
    # --- 境界 40000 ---
    ("gte", 40000, 40000, True), ("gte", 40000, 39999.999, False), ("gte", 40000, 40000.0000001, True),
    ("gt",  40000, 40000, False), ("gt", 40000, 40001, True),
    ("lte", 40000, 40000, True),  ("lte", 40000, 40001, False),
    ("lt",  40000, 40000, False), ("lt", 40000, 39999, True),
    # --- 型を黙って変換しない ---
    ("gte", 40000, "40000", False),          # 文字列数値は数値比較に掛けない
    ("gte", 40000, "50000", False),
    ("gte", 40000, True, False),             # bool は数値でない（既定の線）
    ("gte", 1, True, False),
    ("gte", 40000, None, False),
    ("gte", 40000, datetime.date(2026, 8, 1), False),   # 日付は数値比較に掛けない
    # --- eq ---
    ("eq", 40000, 40000, True), ("eq", 40000, 40001, False),
    ("eq", 40000, 40000.0000001, True),
    ("eq", 40000, "40000", False),
    ("eq", "東京", "東京", True), ("eq", "東京", "東京都", False), ("eq", "東京", None, False),
    # --- contains ---
    ("contains", "東京", "東京都港区", True), ("contains", "東京", "京都市", False),
    ("contains", "東京", None, False),
    ("contains", "40", 140000, False),
    ("contains", "40", "140000", True),
    # --- in（どれか・2026-08-27 追加）------------------------------------------------
    # 仕様（手で書いた・実装から作らない）: 条件値は一覧。セルの値が一覧のどれかと
    # **丸ごと**一致するか。部分一致にしない ── 「りんご」が「青りんご」に当たると
    # 頼んでいない行が黙って混じる（contains で実測した事故と同じ形）。
    ("in", ["みかん", "りんご"], "みかん", True),
    ("in", ["みかん", "りんご"], "りんご", True),
    ("in", ["みかん", "りんご"], "ぶどう", False),
    ("in", ["みかん", "りんご"], "青りんご", False),      # ★ 部分一致にしない
    ("in", ["みかん", "りんご"], "りん", False),
    ("in", ["みかん", "りんご"], None, False),
    ("in", ["みかん", "りんご"], "", False),
    ("in", [], "みかん", False),                          # 空の一覧は誰にも当たらない
    # ★ ここだけ「型を黙って変換しない」の線から外れる（意図して外した）:
    #   in の一覧は**依頼文の語**から作る（人が「100 の行」と書く）ので、突き合わせは
    #   表示上の値で行う。gte/eq の型規律と違うのは、一覧の出どころが違うから。
    #   ★ 危険は小さい: 一覧に載るのは「その列に実在する値」だけを機械が確かめて入れる。
    ("in", ["100"], 100, True),
]


def _impls():
    """校正対象の実装一覧。M2 の ailine_core 側実装が生まれたら自動で対象に入る。"""
    impls = [("ailine._extract_predicate", ailine._extract_predicate)]
    try:
        from ailine_core import extract_multi
        impls.append(("extract_multi", extract_multi.predicate))
    except ImportError:
        pass
    return impls


@pytest.mark.parametrize("impl_name,impl", _impls())
@pytest.mark.parametrize("cmp,value,cell,expected", TRUTH)
def test_predicate_agrees_with_frozen_truth_table(impl_name, impl, cmp, value, cell, expected):
    got = bool(impl(cmp, value)(cell)) if callable(impl(cmp, value)) else None
    assert got == expected, f"{impl_name}({cmp}, {value!r})({cell!r}) = {got}, 表は {expected}"
