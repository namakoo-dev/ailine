# -*- coding: utf-8 -*-
"""threshold — 比較の境目の数は**依頼文**から機械が取る（LLM に確定させない・A' 原則）。

★★ 2026-09-14（言い回し 120 件の盲検・判定者 2 人が一致した誤配 16 件のうち 4 件がこの家系）:
    「残業時間が20時間超えてる人だけ教えて」 → 抽出 残業時間 = 0
    「在庫数が10個切ってるの教えて」         → 抽出 在庫数 = 0
    「在庫少ないやつ出して」                 → 抽出 在庫数 ≤ 0
  依頼文に数が**在る**のに LLM の `0` がそのまま通り、**黙って違う答え**が出ていた。
  ★ SET_WHERE は 09-04 に「閾値は依頼文の数字から機械が取る」と直してあったのに、
    兄弟の EXTRACT には届いていなかった（兄弟間の片配線）。だから器官を **1 つ**にして
    両方が同じ物を呼ぶ ── 次に 3 人目の兄弟が来ても、ここを呼ぶだけで同じ規則になる。
★ 数が無い／複数ある回は**断る**（空欄は誤値より安い）。「少ない」の境目を機械が 0 と決めない。
★ ここは純関数（ファイルも LLM も触らない）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

#: 数値の比較（境目の数が要る比較）。
NUMERIC_CMPS = frozenset({"gt", "lt", "gte", "lte"})

_ZENKAKU = str.maketrans("０１２３４５６７８９，．", "0123456789,.")
#: 数（桁区切りつき・小数つき）と、直後の単位（万／千）。★「10万円」を 10 と読まない。
_NUM = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*(万|千)?")
_UNIT = {"万": 10000.0, "千": 1000.0}


def task_numbers(task: str | None) -> list:
    """依頼文に現れる数を出現順に（全角→半角・桁区切りを外す・万／千を畳む）。"""
    out = []
    for m in _NUM.finditer((task or "").translate(_ZENKAKU)):
        raw = m.group(1).replace(",", "")
        if raw.endswith("."):
            raw = raw[:-1]
        try:
            v = float(raw)
        except ValueError:
            continue
        out.append(v * _UNIT.get(m.group(2) or "", 1.0))
    return out


def fmt(v) -> str:
    """人が依頼文に書いた形に近い数（整数なら整数で）。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(f)) if f.is_integer() else str(f)


@dataclass(frozen=True)
class Grounded:
    """境目の数の接地の結果。`refusal` が空でなければ止める（value は使わない）。"""
    value: float | None
    warning: str = ""
    refusal: str = ""


#: 数を言わずに境目を指す語 ── 値は 0（★ 出所: 盲検 120 件 #21「在庫数がマイナスに
#: なってる行を教えて」（判定者 2 人とも 正）・#74「在庫数がマイナスになってるとこ」）。
#: ★ 頭から思いついた語を足さない ── 実際に打たれた語だけ（訓練常識を昇格させない）。
#: ★★ 比較（マイナス＝未満）まで言える語なので、比較語の表（compare_words）を作る日に
#:   そちらへ移す ── いまは「数の側」だけを持つ（cmp は LLM の lt が通る）。
ZERO_BOUNDARY_WORDS = ("マイナス", "赤字")


def zero_boundary(task: str | None) -> bool:
    """依頼文が数を言わずに 0 を境目にしているか（「マイナスになってる行」）。"""
    text = task or ""
    return any(w in text for w in ZERO_BOUNDARY_WORDS)


def ground(task: str | None, llm_value, *, example: str) -> Grounded:
    """境目の数を依頼文から決める。

    - 依頼文の数が **1 つに決まる**時だけ採る。LLM の値と違えば機械が勝つ（開示つき）
    - 数が無い → 断る（「少ない」「多い」の境目を機械が決めない）
    - 数が複数 → 断る（どれが境目か決められない）
    `example` は断り文に添える「通る書き方」（呼び出し側の op に合った例）。
    """
    distinct = sorted(set(task_numbers(task)))
    if not distinct and zero_boundary(task):
        # ★ 「マイナスになってる行」は数を言っていないが境目は 0（盲検 #21 が 正 だった読み）。
        #   ★★ ①の初版はここを断りに変えてしまい、**正だった 1 件を壊した**（実測で気づいた）。
        distinct = [0.0]
    if not distinct:
        return Grounded(None, refusal=(
            "境目の数が依頼文にありません ── 「多い／少ない」だけでは機械が線を引けません"
            f"（例: 「{example}」のように数を書いてください）"))
    if len(distinct) > 1:
        return Grounded(None, refusal=(
            "境目の数が依頼文から一意に決まりません"
            f"（見つかった数: {'、'.join(fmt(v) for v in distinct)}）── "
            f"1 つだけ書いてください（例: 「{example}」）"))
    v = distinct[0]
    warning = ""
    try:
        llm = float(llm_value)
    except (TypeError, ValueError):
        llm = None
    if llm is not None and abs(llm - v) > 1e-9:
        warning = (f"LLM が返した値({fmt(llm)})と依頼文の数({fmt(v)})が食い違うため"
                   f"依頼文の数({fmt(v)})を採用しました")
    return Grounded(v, warning=warning)


# ── 引き算・割り算の向き（2026-09-14・言い回し 120 件の盲検・誤配の家系③）──────────────
#
# ★★ 「出勤と退勤の時刻から実働時間を計算する列を作って」→ `出勤 − 退勤`（符号が逆）。
#   「AとBから」は**向きを言っていない**のに、機械が並び順をそのまま演算の順にしていた。
#   ★ 引き算と割り算は向きで答えが変わる ── 読めないなら**聞き返す**（足し算に落とさない）。
#   ★ 語の列挙で正しい（`_REMOVAL_WORDS` と同じ理屈: 動作の向きは言葉にしか現れない）。
#: 向きを言う型（★ 出所は盲検 120 件の実文）。`{x}` が前・`{y}` が後ろ（結果 = x 演算 y）。
#: ★ 列名は文中に**何度でも**現れる（「売上と原価の間に、売上から原価を引いた利益の列を作って」）
#:   ので、位置ではなく**型**で読む ── 初版は最初の出現だけを見て、この実文を読み落とした。
SUBTRACT_PATTERNS = (
    r"{x}[^。]{{0,8}}から[^。]{{0,12}}{y}[^。]{{0,6}}を[^。]{{0,6}}引",
    r"{y}[^。]{{0,6}}を[^。]{{0,12}}{x}[^。]{{0,8}}から[^。]{{0,6}}引",
    r"{x}\s*(?:マイナス|[-−ー])\s*{y}",
)
DIVIDE_PATTERNS = (
    r"{x}[^。]{{0,6}}を[^。]{{0,12}}{y}[^。]{{0,6}}で[^。]{{0,6}}割",
    r"{x}\s*[÷/]\s*{y}",
)


def direction_of(task: str | None, operands, operator: str = "-") -> list | None:
    """依頼文が言っている被演算子の順（読めなければ None）。

    ★ どちらの向きにも読める文（両方の型が当たる）も None ── 機械が選ばない。
    """
    text = task or ""
    if not (isinstance(operands, (list, tuple)) and len(operands) == 2):
        return None
    a, b = str(operands[0]), str(operands[1])
    if not a or not b or a == b:
        return None
    pats = DIVIDE_PATTERNS if operator == "/" else SUBTRACT_PATTERNS
    hits = []
    for x, y in ((a, b), (b, a)):
        for pat in pats:
            if re.search(pat.format(x=re.escape(x), y=re.escape(y)), text):
                hits.append([x, y])
                break
    return hits[0] if len(hits) == 1 else None
