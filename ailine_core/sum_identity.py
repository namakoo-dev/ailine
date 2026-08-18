"""sum_identity — 算術恒等の検算: 「自分より上の全部を足した値」を持つ行を見つける。

★ なぜ（独立レビューの実測）: `check_append_total` は**期待値を「合計式が生成したのと
同じ範囲」から作っていた** ―― 検算が被検算と同じ盲点を使う恒真式。既存の合計 300 を
持つ表に合計を足すと 600 が書かれ、「3 行の合計を検証」と言って `✓ 機械検証済み`・
exit 0 で原本を上書きした。並べ替えの事後条件にも同型がある: 値のみの合計行つきの表を
降順に並べ替えると、合計行が 2 行目に来ても「5 行を検証（降順）」で通り ✓ が出た。
この型は repo が既に禁じている（docs/behavior-corpus/nodes/empty-verification-ban.md）。

★ 設計: **語を一切読まない。** 辞書（「合計」「総計」「小計」）も書式（太字・罫線）も
行の型も使わず、数値の並びだけを見る。合計行が合計行である理由は言語でも装飾でもなく
**算術**だから ―― 英語でも韓国語でも記号でも、同じ恒等式が同じように成り立つ。
この module に docstring 以外の文字列リテラルが1つも無いことは番人テストが機械で守る
（tests/test_sum_identity_unit.py::test_module_reads_no_words）。

★ 返すのは「位置」: 恒等式が成り立つ行が**在ること**自体は異常ではない ―― 表の一番下に
合計があるのは正常な帳票そのものだから。異常なのは、その行が**最下行でない**とき
（上に合計があるのに、その下でもう一度足している＝二重計上）。だから戻り値の各件は
`is_last` を持ち、「存在」ではなく「位置」で判定できる形にしてある。

★ 置き場所: ailine_core/（write_precondition.py と同じ理由）。標準ライブラリだけで閉じ、
openpyxl にも ailine.py にも依存しない ―― 表計算に限らずどんな数値の並びにも使える。
"""
from __future__ import annotations

from dataclasses import dataclass

# 「合計」とみなすために必要な、自分より上にある数値の最小個数。
# ★ 1 だと「2 行目が 1 行目と等しい」だけで当たる（同じ金額が2つ並ぶ表は珍しくない）。
#   2 未満を許すと、この検算は帳票の大半で鳴りっぱなしになる。
MIN_TERMS = 2

# 浮動小数の丸め許容（絶対値 + 相対値）。ailine.py の事後条件が使う 1e-6 に揃える。
TOLERANCE = 1e-6
RELATIVE_TOLERANCE = 1e-9


@dataclass(frozen=True)
class SumIdentityHit:
    """恒等式が成り立った行1件。

    row:       その行番号（呼び出し側が渡した番号をそのまま返す）
    value:     その行の値
    term_rows: 足し合わせた行番号（上から順・数値でない行は入らない）
    total:     term_rows の合計（value と一致したのでこの件が立った）
    is_last:   その行が、渡された並びの中で**最後の数値行**か
    """
    row: int
    value: float
    term_rows: tuple
    total: float
    is_last: bool


def _is_number(v) -> bool:
    """bool は int のサブクラスだが数値としては扱わない（ailine.py の _is_number と同じ線）。"""
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def rows_matching_sum_above(values, *, min_terms: int = MIN_TERMS,
                             tolerance: float = TOLERANCE) -> list:
    """自分より上の全数値の合計と一致する値を持つ行を、上から順に並べて返す。

    values: (行番号, 値) の並び。値が数値でないもの（文字列・None・bool）は
            **項にも候補にもしない** ―― 0 として足すと合計が狂い、候補にすると
            ラベル行が当たる。除外であって 0 扱いではない。
    min_terms: 合計とみなすのに要る、上にある数値の最小個数（MIN_TERMS 参照）。
    tolerance: 浮動小数の丸め許容（絶対 + 相対）。

    ★ 合計が 0 の場合は立てない: 0 == 0 + 0 は恒真で、空欄を 0 で埋めた列が
      丸ごと当たる（0cf9218「空虚な検証合格の禁止」の裏返し ―― 空虚な**不合格**も出さない）。
    """
    numeric = [(row, float(v)) for row, v in values if _is_number(v)]
    last_row = numeric[-1][0] if numeric else None
    hits: list = []
    term_rows: list = []
    running = 0.0
    for row, value in numeric:
        if len(term_rows) >= min_terms and running != 0.0 \
                and abs(value - running) <= tolerance + RELATIVE_TOLERANCE * abs(running):
            hits.append(SumIdentityHit(row=row, value=value, term_rows=tuple(term_rows),
                                        total=running, is_last=row == last_row))
        term_rows.append(row)
        running += value
    return hits
