"""date_compare — EXTRACT の日付範囲比較（台帳 DATE_RANGE_AGG の正体）。

★ なぜ在るか（2026-08-24 の実測）: 「3/26 から 4/25 までの分だけ現場ごとに集計して」は、
7B が正しく複合計画（EXTRACT×2 → PIVOT）に落としているのに、EXTRACT が
「比較『gte』には数値の値が必要ですが『2026/3/26』は数値に変換できません」で止まっていた。
台帳では `DATE_RANGE_AGG` という**不足 op**として名指ししていたが、実体は**既存 op の穴**
だった ── 新しい op を足さずに、この 1 箇所を埋めれば締め日の期間集計が通る。

★ なぜ「シリアル値に直す」で足りるか: LibreOffice の日付セルは getValue() が
シリアル値（既定 null date = 1899-12-30 起点の日数）を返す。ヘルパ ExtractRows は
既に getValue() で数値比較しているので、**Basic 側も codegen 側も 1 行も変えずに済む**。
変えるのは「依頼文の日付リテラルを、その列に合うシリアル値に直す」ここだけ。

★ 締め日の作法（実務がそう）: 「3/26〜4/25」は**両端を含む**。さらに、日付だけを
書いた人の意図は「その日いっぱい」なので、時刻つきの列でも 4/25 23:00 を落とさない
（lte/gt は日の終わりを閾値にする）。ここを取りこぼすと締め日の売上が静かに消える。
"""
from __future__ import annotations

import datetime as dt
import re

# LibreOffice / Excel の既定 null date。ここがずれると全部 1 日ずれる。
_EPOCH = dt.date(1899, 12, 30)

# 1 日の終わり（23:59:59.9 まで含む）。lte/gt の閾値に足す。
END_OF_DAY = 1 - 1e-6

_DATE_RE = re.compile(
    r"^\s*(?P<y>\d{4})\s*[-/.年]\s*(?P<m>\d{1,2})\s*[-/.月]\s*(?P<d>\d{1,2})\s*日?\s*$"
)

#: 元号と、その元年の前年（元年 = +1 年）。★ **法で決まる閉じた一覧**で、増えるのは
#:  代替わりの時だけ ── 言い回しの列挙（足し続けないと当たらない形）ではない。
_ERAS = {"令和": 2018, "平成": 1988, "昭和": 1925, "大正": 1911, "明治": 1867}

_WAREKI_RE = re.compile(
    r"^\s*(?P<e>" + "|".join(_ERAS) + r")\s*(?P<y>\d{1,2}|元)\s*[-/.年]"
    r"\s*(?P<m>\d{1,2})\s*[-/.月]\s*(?P<d>\d{1,2})\s*日?\s*$"
)


def parse_date_literal(raw) -> dt.date | None:
    """依頼文から来た値を日付として読む。読めなければ None。

    ★ 読まないものを明示する:
      - 年の無い「3/26」── どの年かは機械が決めてよい話ではない
      - 和暦「令和8年3月26日」── ここでは読まない。読みたい呼び出し側が
        `parse_wareki_literal` を**自分で足して**合成する（下の docstring 参照）
      - 数値「100」── 数値は数値のまま（シリアル値として日付に化けさせない）
    """
    if isinstance(raw, dt.datetime):
        return raw.date()
    if isinstance(raw, dt.date):
        return raw
    if not isinstance(raw, str):
        return None
    m = _DATE_RE.match(raw)
    if not m:
        return None
    try:
        return dt.date(int(m.group("y")), int(m.group("m")), int(m.group("d")))
    except ValueError:      # 2026/2/30 のような存在しない日
        return None


def parse_wareki_literal(raw) -> dt.date | None:
    """和暦の日付**だけ**を読む（令和8年8月31日 → 2026-08-31）。読めなければ None。

    ★★ なぜ `parse_date_literal` に混ぜないか（2026-09-12）: あちらは依頼文の値を読み、
      列の型判定（`classify_date_column`）にも使われる。和暦を混ぜると **`ailine run` の
      経路と列の型判定にも黙って波及する** ── 測った範囲（帳票を読む器官）を超える。
      暦の算術（判定）はここに 1 つだけ置き、**どの暦を受けるかの合成は呼び出し側**に置く。

    ★ 読まないもの:
      - 「R8.8.31」「H31/4/1」のアルファベット略記 ── 帳票 178 冊で **0 件**（2026-09-12 実測）で、
        `R8` は品番・部屋番号・セル番地にも見える。★ 発火条件: 実物で出たら、その時に測って足す
      - 年だけ・月日だけ（元号が付いていても日付にならない）

    ★ 元年は 1 年（令和元年 = 2019）。存在しない日（令和8年2月30日）は None。
    """
    if isinstance(raw, (dt.date, dt.datetime)):
        return None                       # ★ 既に日付 ── ここの仕事ではない（呼び分けを曖昧にしない）
    if not isinstance(raw, str):
        return None
    m = _WAREKI_RE.match(raw)
    if not m:
        return None
    y = 1 if m.group("y") == "元" else int(m.group("y"))
    if y < 1:
        return None
    try:
        return dt.date(_ERAS[m.group("e")] + y, int(m.group("m")), int(m.group("d")))
    except ValueError:
        return None


def date_to_serial(d: dt.date) -> int:
    """日付 → LibreOffice/Excel のシリアル値（1899-12-30 起点）。"""
    if isinstance(d, dt.datetime):
        d = d.date()
    return (d - _EPOCH).days


def threshold_for(cmp: str, d: dt.date) -> float:
    """比較の種類に応じた閾値。★ 「その日いっぱい」を含める側を日の終わりにする。

      gte D → D の 0 時（D 当日を含む）
      lte D → D の日の終わり（D 当日を**含む**。時刻つきでも落とさない）
      gt  D → D の日の終わり（D 当日を**含まない**）
      lt  D → D の 0 時（D 当日を含まない）
      eq  D → D の 0 時（日付だけの列を想定。時刻つきの列では当たらない＝呼び出し側が警告）
    """
    base = float(date_to_serial(d))
    if cmp in ("lte", "gt"):
        return base + END_OF_DAY
    return base


def classify_date_column(values) -> tuple[str, bool]:
    """列の中身から、その列が日付として比較できるかを決める。

    戻り値 (kind, has_time):
      "date"      … 日付/日時が入っている（シリアル値で比較できる）
      "text_date" … 日付に**見える文字列**が入っている（★ 辞書順で比べてはいけない
                    ── "2026/3/26" > "2026/12/1" になる。正直に断る側）
      "other"     … 日付ではない
      "empty"     … 中身が無い
    """
    seen_date = seen_text_date = seen_other = 0
    has_time = False
    for v in values:
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        if isinstance(v, dt.datetime):
            seen_date += 1
            if (v.hour, v.minute, v.second, v.microsecond) != (0, 0, 0, 0):
                has_time = True
        elif isinstance(v, dt.date):
            seen_date += 1
        elif isinstance(v, str) and parse_date_literal(v) is not None:
            seen_text_date += 1
        else:
            seen_other += 1
    if seen_date and not seen_other and not seen_text_date:
        return "date", has_time
    if seen_text_date and not seen_other and not seen_date:
        return "text_date", False
    if seen_date or seen_text_date or seen_other:
        return "other", has_time
    return "empty", False
