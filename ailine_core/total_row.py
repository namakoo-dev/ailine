"""単位L — 合計行の識別と、除外を算術で閉じる検査。DESIGN-20260821-multifile.md v2.1。

★ なぜ sum_identity.py（書き側の検算）と別物か:
読み側（縦積み）は取り逃がし=黙って二重計上で、書き側と非対称が反転する ── だから
広く候補を拾い（ラベル語+構造）算術で閉じる検査に回す。ラベル語は断片誤爆を避けるため
『計』だけ完全一致にする（『設計部』『会計課』を部分一致で誤爆しない）。

★ sum_identity.py には触らない（「語を読まない」不変条件・番人つき）。この module は
独立に完結し、sum_identity を呼ぶ必要も無い。
"""
from __future__ import annotations

from dataclasses import dataclass

# ラベル語トリガ。『合計』『小計』『総計』は部分一致で可（『合計金額』も拾う）。
# 『計』だけは完全一致のみ（正規化後のラベル全体が「計」と一致する時だけ）。
_LABEL_SUBSTRINGS = ("合計", "小計", "総計")
_LABEL_EXACT = "計"

# 算術恒等（閉じる検査）の許容誤差。ailine.py の事後条件・sum_identity.TOLERANCE に揃える。
TOLERANCE = 1e-6


@dataclass(frozen=True)
class ExcludedRow:
    """除外した1行。row: 行番号 / label・value: 元のセル値 / reason: 排除トリガの種別。"""
    row: int
    label: object
    value: object
    reason: str


@dataclass(frozen=True)
class Mismatch:
    """閉じる検査が不一致だった1件。両側の数字つき（信用の条件④・感想で終わらせない）。"""
    row: int
    excluded_value: float
    adopted_sum: float


@dataclass(frozen=True)
class TotalRowVerdict:
    """split_total_rows の戻り値。excluded: 除外行 / adopted_rows: 採用した行番号
    （数値の無い空行は含めない）/ mismatches: 算術恒等が閉じなかった件。"""
    excluded: list
    adopted_rows: list
    mismatches: list


def _is_number(v) -> bool:
    """bool は int のサブクラスだが数値としては扱わない（sum_identity._is_number と同じ線）。"""
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_blank_cell(v) -> bool:
    """None、または空白のみの文字列（半角・全角空白）を空とみなす。"""
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip(" 　") == ""
    return False


def _normalize_label(label) -> str:
    """前後の半角・全角空白と末尾コロン（半角/全角）を剥がす。"""
    if _is_blank_cell(label):
        return ""
    s = str(label).strip(" 　")
    if s.endswith(":") or s.endswith("："):
        s = s[:-1].strip(" 　")
    return s


def _label_is_total_word(label) -> bool:
    """ラベル語トリガ: 『計』は完全一致のみ・『合計/小計/総計』は部分一致で可。"""
    norm = _normalize_label(label)
    if not norm:
        return False
    if norm == _LABEL_EXACT:
        return True
    return any(w in norm for w in _LABEL_SUBSTRINGS)


def split_total_rows(rows) -> TotalRowVerdict:
    """rows: (行番号, ラベルセル値, 数値セル値) の列（上から順）。

    排除トリガ（語と構造のみ・算術はトリガにしない）:
      a. ラベル語（合計/小計/総計は包含・計は完全一致）を持ち、数値がある行
      b. ラベルが空（None/空文字/空白のみ）かつ数値がある行
      c. 直上の行が空行（ラベルも値も無い）かつ数値がある行

    算術恒等は閉じる検査専用: 各除外行について、除外行の値が「先頭からの累積和」
    または「直前の除外行より後ろの区間和」のどちらかと一致すれば閉じる（許容誤差
    TOLERANCE）。★ jisaku-review#2: 小計を複数持つ表では総計は累積で・各小計は
    区間で閉じるので、この OR が必要（片方だけだと2個目以降の小計に偽 ⚠ が出る・実測）。
    両方外れた時だけ本当の不一致 ── mismatches に両側の数字つきで積む（除外自体は
    維持する ── 除外の正しさを疑うための検査であって、除外を取り消す検査ではない）。
    """
    excluded_raw = []          # (row, label, value, reason)
    adopted_rows = []
    adopted_numeric = []       # (row, float value) — 閉じる検査の分母
    prev_row_is_blank = False  # 直上行が「ラベルも値も無い」か（先頭行の上は存在しないので False）

    for row_num, label, value in rows:
        label_blank = _is_blank_cell(label)
        value_blank = _is_blank_cell(value)
        has_number = _is_number(value)

        reason = None
        if has_number and _label_is_total_word(label):
            reason = "ラベル語"
        elif has_number and label_blank:
            reason = "ラベル空白"
        elif has_number and prev_row_is_blank:
            reason = "直上空行"

        if reason:
            excluded_raw.append((row_num, label, value, reason))
        elif has_number:
            adopted_rows.append(row_num)
            adopted_numeric.append((row_num, float(value)))
        # else: 数値の無い空行・データにならない行 ── 採用にも除外にも数えない

        prev_row_is_blank = label_blank and value_blank

    # ★ jisaku-review#2 major の直し: 複数の小計グループを持つ表では、除外行ごとに
    # 「先頭からの累積和」だけで閉じるかを見ると、2個目以降の小計に偽 ⚠ が出る
    # （実測: 部署A小計→部署B小計→総計 の表で、部署B小計に部署A分まで足した累積が
    # 比較されて不一致になった）。閉じ方は2通りあってよい ── 小計は「直前の除外行より
    # 後ろの区間和」で閉じ、総計は「先頭からの累積和」で閉じる。どちらか一方が合えば
    # 閉じている（両方外れた時だけ本当の不一致）。
    mismatches = []
    excluded = []
    prev_excluded_row = None
    for row_num, label, value, reason in excluded_raw:
        cumulative_sum = sum(v for r, v in adopted_numeric if r < row_num)
        lower_bound = prev_excluded_row if prev_excluded_row is not None else -1
        segment_sum = sum(v for r, v in adopted_numeric if lower_bound < r < row_num)
        excluded_value = float(value)
        closes_cumulative = abs(excluded_value - cumulative_sum) <= TOLERANCE
        closes_segment = abs(excluded_value - segment_sum) <= TOLERANCE
        if not (closes_cumulative or closes_segment):
            mismatches.append(Mismatch(row=row_num, excluded_value=excluded_value, adopted_sum=cumulative_sum))
        excluded.append(ExcludedRow(row=row_num, label=label, value=value, reason=reason))
        prev_excluded_row = row_num

    return TotalRowVerdict(excluded=excluded, adopted_rows=adopted_rows, mismatches=mismatches)
