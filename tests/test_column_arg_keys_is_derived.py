# -*- coding: utf-8 -*-
"""★★ 連鎖の読み替え名簿は、宣言から導いた集合と一致すること（2026-09-16）。

★ 出所: `_COLUMN_ARG_KEYS` は事故のたびに 1 本ずつ足されてきた名簿だった
  （CHART の category_col が「4 本目の配線」と試験に記録されている）。
  配線盤（scripts/wiring_board.py）で数えたら、宣言 4 op に対して同じ扱いを受けるべき op が
  15 在り、実際に落ちていた:

      「単価と数量を掛けた列を作って、金額で部門別にまとめて」
        集計・並べ替え・数値書式 → 通る（前段が作った列を指していると読み替える）
        ピボット・合計追加       → 「列『金額』がありません」で断る

  ★ 道具が自分で作った列を、自分で使えない ── 買い手にはそう見える。

★ だから名簿を**宣言から導く**形にした。導き方はここに 1 本だけ置き、本体は
  その結果をリテラルで持つ（＝配線盤の走査から消えないため・下記）。ずれたらここが赤くなる。
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

import ailine  # noqa: E402

#: ★ 理由つきの例外 1 件。CHART の `category_col` は**列を名指しする引数**だが
#:   **判定される対象ではない**（2026-08-23 の検分で SUBJ_INPUT に決めた ── 対象にすると
#:   「商品ごとの構成比を円グラフにして」で value_col への誤爆が出て ✓ が消える。
#:   tests/test_chart_kind_boundary.py が実機再現で固定している）。
#:   2 つの表は別の問いに答えているので、揃えるのではなく理由を書いて足す。
EXTRA = {"CHART": ("category_col",)}


def derived() -> dict:
    """OP_SUBJECT_SLOTS と OP_WRITE_TARGET から、読み替えてよい引数を導く。

    ・列と宣言された引数のうち、**書き込み先の列**（col_key）は外す
      …… 読み替えを誤ると「エラー」ではなく「別の列への書き込み」になる
    ・writes に remove を含む op は丸ごと外す
      …… 前段が作った列を**消す**方向に化ける
    """
    out = {}
    for op in sorted(ailine.OP_SCHEMA):
        wt = ailine.OP_WRITE_TARGET[op]
        if ailine.WRITE_REMOVE in wt.writes:
            continue
        keys = tuple(k for k, kind in ailine.OP_SUBJECT_SLOTS.get(op, ())
                     if kind == ailine.SUBJ_COLUMN and k != wt.col_key)
        keys = tuple(dict.fromkeys(keys + EXTRA.get(op, ())))
        if keys:
            out[op] = keys
    return out


def test_the_derivation_is_not_empty():
    """★ 空回りの検出 ── 導出が空なら、下の一致も自明に通る。"""
    d = derived()
    assert len(d) >= 12, f"導出が痩せている: {sorted(d)}"


def test_the_roster_equals_the_derivation():
    """★★ 宣言＝導出。op を足した人がここを直し忘れたら赤になる。"""
    got = {k: tuple(v) for k, v in ailine._COLUMN_ARG_KEYS.items()}
    want = derived()
    assert got == want, (
        "連鎖の読み替え名簿が宣言から導いた集合とずれている。"
        f"名簿だけ: {sorted(set(got) - set(want))} / 導出だけ: {sorted(set(want) - set(got))} / "
        f"引数違い: {sorted(k for k in set(got) & set(want) if got[k] != want[k])}")


def test_ops_that_remove_are_never_rescued():
    """★ 反対側の検算 ── 削除する op を読み替えると、前段が作った列を消す。

    ★ 届く範囲を広げる変更は、**広げてはいけない側**まで見て初めて完了する。
    """
    for op in sorted(ailine.OP_SCHEMA):
        if ailine.WRITE_REMOVE in ailine.OP_WRITE_TARGET[op].writes:
            assert op not in ailine._COLUMN_ARG_KEYS, f"削除する op が救済対象に入っている: {op}"


def test_the_write_destination_column_is_never_rewritten():
    """★ 反対側の検算 ── 書き込み先の列を読み替えると、別の列へ書く事故になる。"""
    for op, keys in ailine._COLUMN_ARG_KEYS.items():
        ck = ailine.OP_WRITE_TARGET[op].col_key
        if ck:
            assert ck not in keys, f"{op}: 書き込み先の列 {ck} が読み替え対象に入っている"


def test_the_chart_exception_is_kept_with_its_reason():
    """★ 例外は消えないこと（2026-08-23 の検分の結論を落とさない）。"""
    assert "category_col" in ailine._COLUMN_ARG_KEYS.get("CHART", ()), (
        "CHART の category_col が落ちた ── 依存つき連鎖の 4 本目の配線")
    assert "CHART" in EXTRA and EXTRA["CHART"] == ("category_col",)


def test_the_roster_stays_a_literal_so_the_board_can_see_it():
    """★★ 実装前に確かめたこと ── 計算式で書くと**配線盤から消える**。

    走査（tests/test_op_completeness.py の discover_op_rosters）は dict リテラルだけを見る。
    内包表記にすると列が丸ごと盤外に出て、目録と被覆の番人も赤くなる。
    だから「導出は試験の側に置き、本体はリテラルで持つ」形にした。ここはその形を固定する。
    """
    import ast
    from _product_source import src_files
    # ★ 2026-09-23: 本体 1 冊でなく src の下を全部見る（配線盤の走査と同じ視野）。
    for path in src_files():
        for n in ast.walk(ast.parse(path.read_bytes().decode("utf-8"))):
            if (isinstance(n, ast.Assign)
                    and any(getattr(t, "id", "") == "_COLUMN_ARG_KEYS" for t in n.targets)):
                assert isinstance(n.value, ast.Dict), (
                    "_COLUMN_ARG_KEYS が dict リテラルでなくなった ── 配線盤の走査から消えます")
                return
    raise AssertionError("_COLUMN_ARG_KEYS の代入が見つからない")
