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
from ailine_core.primitives import is_number as _is_number

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


def closes_as_total(value: float, adopted: list, last_excluded_row) -> bool:
    """『この値は、それまでの採用行の和として説明がつくか』── 累積和 OR 区間和。

    ★ 2026-08-24（第三波 S4）: トリガ c（直上が空行）の**裏取り**に使う。
    adopted は [(行番号, float), ...]。単一列版と複数列版が**同じ実装**を呼ぶ
    （片配線の禁止 ── 実測で単一列版だけ直して複数列版が素通りしかけた）。
    """
    cum = sum(v for _r, v in adopted)
    lower = last_excluded_row if last_excluded_row is not None else -1
    seg = sum(v for r, v in adopted if r > lower)
    return abs(value - cum) <= TOLERANCE or abs(value - seg) <= TOLERANCE


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
    last_excluded_row = None   # ★ S4: トリガ c の裏取り（区間和）に要る

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
            # ★ 2026-08-24（第三波 S4・実測で本物の行が消えた）: 「直上が空行」は
            #   三つのトリガのうち**最も弱い証拠**で、ラベルが本物の行にも当たる。
            #   実測: [a,100] / [空白だけの行] / [b,200] で b が合計行として除外され、
            #   200 が黙って消えた（しかも Σ は除外後の値を『元』にしていたので
            #   元 100 / 出力 100 で恒真に合った）。
            #   ★ 絞り: 弱い証拠は**裏取りを要求する** ── 算術が閉じる時だけ除外する。
            #   （a/b は構造そのものが「合計」と言っているので従来どおり無条件。
            #   閉じる検査を『除外を取り消す検査には使わない』という設計は、証拠が
            #   強い a/b に対する取り決めであって、c には及ばない。）
            if closes_as_total(float(value), adopted_numeric, last_excluded_row):
                reason = "直上空行"

        if reason:
            last_excluded_row = row_num
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


# ---- 複数数値列版（operator 盲検7度目・$0 の主因の直し・2026-08-21）----
# ★ 単一列版（split_total_rows・上記）はそのまま・触らない（凍結済み10検体）。
# 真因: 全トリガが「指定の数値列（=基準の最初の数値列）に数字がある」を前提にしていた。
# 実務標準形（数量・単価つき請求書）は最初の数値列=数量だが、合計行の数字は金額列にしか
# 無い ── has_number=False で全トリガが沈黙し、Σ が黙って2倍になった（stack の読み側は
# 取り逃がし=黙って二重計上で書き側と非対称が反転する、v2.1 冒頭の設計どおりの実害）。


@dataclass(frozen=True)
class MismatchColumn:
    """閉じる検査が不一致だった1件（複数数値列版）。列名つき ── 単一行が複数の数値列で
       同時に不一致になりうる（各列は独立に検査する）。"""
    row: int
    column: str
    excluded_value: float
    adopted_sum: float


def split_total_rows_multi(rows) -> TotalRowVerdict:
    """rows: (行番号, ラベルセル値, {列名: セル値, ...}) の列（上から順）。

    合計行の候補判定を『対象の数値列集合のどこかに数字がある』に広げた版（単一列版は
    『指定の1本の数値列』だけを見る）。排除トリガの種別（ラベル語/ラベル空白/直上空行）は
    単一列版と同じ語・同じ正規化（_label_is_total_word・_normalize_label を共有）。

    閉じる検査は列ごとに独立: その除外行が数字を持つ**各**数値列について、その列の
    採用行の和（区間 OR 累積・単一列版と同じ OR ロジック）と突き合わせる。不一致は
    列名つきで mismatches（MismatchColumn）に積む（除外自体は維持する）。

    戻り値は単一列版と同じコンテナ（TotalRowVerdict）── excluded は ExcludedRow の
    ままだが、.value は『その行で数字を持つ最初の列（列順）の値』を代表値として使う
    （呼び出し側が特定の列の値を見たければ元セルから引き直せる ── 除外の事実と理由の
    開示に必要な最小限）。"""
    excluded_raw = []          # (row, label, {col: value}, reason)
    adopted_rows = []
    adopted_numeric: dict = {}  # {col: [(row, float value), ...]}
    prev_row_is_blank = False
    last_excluded_row_m = None   # ★ S4: トリガ c の裏取り（区間和）に要る

    col_order: list = []
    for _r, _l, values in rows:
        for c in values:
            if c not in col_order:
                col_order.append(c)
    for c in col_order:
        adopted_numeric[c] = []

    for row_num, label, values in rows:
        label_blank = _is_blank_cell(label)
        numeric_here = [(c, values[c]) for c in col_order
                        if c in values and _is_number(values[c])]
        has_number_any = bool(numeric_here)
        all_values_blank = all(_is_blank_cell(values.get(c)) for c in col_order)
        row_fully_blank = label_blank and all_values_blank   # ★ 直上空行判定: ラベルも数値も無い行

        reason = None
        if has_number_any and _label_is_total_word(label):
            reason = "ラベル語"
        elif has_number_any and label_blank:
            reason = "ラベル空白"
        elif has_number_any and prev_row_is_blank:
            # ★ S4 の絞り（単一列版と同じ線・同じ関数）: 最も弱い証拠なので裏取りを要求。
            #   複数列版では **その行が数字を持つ全ての列**が閉じた時だけ除外する
            #   ── 1 列でも説明がつかなければ「合計」という話は崩れており、
            #   本物のデータ行を黙って消す側の危険のほうが大きい。
            if all(closes_as_total(float(v), adopted_numeric[c], last_excluded_row_m)
                    for c, v in numeric_here):
                reason = "直上空行"

        if reason:
            last_excluded_row_m = row_num
            excluded_raw.append((row_num, label, dict(numeric_here), reason))
        elif has_number_any:
            adopted_rows.append(row_num)
            for c, v in numeric_here:
                adopted_numeric[c].append((row_num, float(v)))
        # else: 数値の無い空行・データにならない行 ── 採用にも除外にも数えない

        prev_row_is_blank = row_fully_blank

    mismatches = []
    excluded = []
    prev_excluded_row = None
    for row_num, label, numeric_vals, reason in excluded_raw:
        for c, v in numeric_vals.items():
            excluded_value = float(v)
            col_adopted = adopted_numeric.get(c, [])
            cumulative_sum = sum(val for r, val in col_adopted if r < row_num)
            lower_bound = prev_excluded_row if prev_excluded_row is not None else -1
            segment_sum = sum(val for r, val in col_adopted if lower_bound < r < row_num)
            closes_cumulative = abs(excluded_value - cumulative_sum) <= TOLERANCE
            closes_segment = abs(excluded_value - segment_sum) <= TOLERANCE
            if not (closes_cumulative or closes_segment):
                mismatches.append(MismatchColumn(row=row_num, column=c,
                                                  excluded_value=excluded_value,
                                                  adopted_sum=cumulative_sum))
        representative_value = next(iter(numeric_vals.values()), None)
        excluded.append(ExcludedRow(row=row_num, label=label, value=representative_value,
                                    reason=reason))
        prev_excluded_row = row_num

    return TotalRowVerdict(excluded=excluded, adopted_rows=adopted_rows, mismatches=mismatches)


# ---- 第二の独立検出器（語のトリップワイヤ・恒真切り・operator 盲検7度目 修正2）----
# ★ 意図的に検出器1（split_total_rows/split_total_rows_multi）と盲点を共有しない設計。
# 列解決を一切使わない（『どれがラベル列か』を知らずに、行の全セル値をただ走査する）──
# 検出器1が何らかの理由で沈黙しても、黙って倍額にはならない、が保証の中身。
# 誤爆（摘要に『7月合計分』等）は ⚠ 1個の確認コストで受ける（疑わしきは鳴らす）。


def row_has_total_word(values) -> str | None:
    """values: 1行のセル値の列（列の意味・並びは問わない）。いずれかの値が合計語
       （_label_is_total_word と同じ規則: 合計/小計/総計は部分一致・計は完全一致・
       断片ガードで『設計部』等は誤爆しない）に一致すれば、正規化後の語を返す。
       無ければ None。"""
    for v in values:
        if _label_is_total_word(v):
            return _normalize_label(v)
    return None


def total_word_trip_findings(rows) -> list:
    """rows: [(識別子（例: ファイル名）, 行番号, [セル値, ...]), ...]。
       合計語を持つ行を名指しで集める（除外はしない・検出のみ）。
       戻り値: [(識別子, 行番号, 見つかった語), ...]（見つかった順・一括検出）。"""
    out = []
    for ident, row_num, values in rows:
        word = row_has_total_word(values)
        if word:
            out.append((ident, row_num, word))
    return out
