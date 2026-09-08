# -*- coding: utf-8 -*-
"""依頼文が**式そのものを書いている**とき、実行した計算がそれと同じか。

★★ なぜ在るか（2026-09-08・盲検の検品が唯一の false ✓ として拾った）:

    表     商品 / 売上 / 原価          ← ★ 『利益』という列は**無い**
    依頼   「利益率（**利益÷売上**）の列を追加して」   ← 括弧で式まで書いている
    実行   操作:計算列 演算対象:**売上 と 原価** 演算子:/ 新しい列の名前:利益率
    実物   りんご **1.714**（頼んだ式なら 0.417）
    出力   **✓ 機械検証済み**

  ★ 既存の 2 つの関所はどちらも鳴らない:
      残差（列名）  『利益』は新しい列の名前『利益**率**』に部分一致して消費される
      効果の種類    計算列を頼んで計算列を実行 ── 食い違わない
    宣言（売上÷原価）と実体（=B/C）は完全に一致しているので、事後条件も通る。
    ★ また**依頼**だけが見られていない ── 三項のうち一項が欠けた形。

★★ 何を見るか ── 依頼文の中の**演算記号**を挟む二語と、宣言した演算対象・演算子:

      依頼に `A÷B` の形が在る
      かつ 宣言が `演算対象:X と Y 演算子:Z` を持つ
      かつ (A,B,÷) と (X,Y,Z) が食い違う                → ✓ を出さない

  ★ 割り算と引き算は**順序が意味を変える**ので順序も見る。掛け算と足し算は見ない。
  ★ 宣言が演算対象を持たない回は**黙る** ── 日付（2026/09/30）や社名（A社/B社）の
    区切り記号を式と読み違える危険があり、そこは他の関所の受け持ち。

★ 直さない・止めない ── ⚠ を出して ✓ を降ろすだけ（residue/intent と同じ作法）。
★ ailine を import しない（可搬性の番人が機械で守る層）。
"""
from __future__ import annotations

import re

#: 依頼文に現れる演算記号 → 宣言側の演算子の書き方。
SIGNS = {"÷": "/", "／": "/", "/": "/",
         "×": "*", "＊": "*", "*": "*",
         # ★ 長音符『ー』は入れない ── 実測で「請求明細シート」を『請求明細シ ー ト』と
         #   読んで誤爆した（9,128 件の実走行で唯一の誤爆・2026-09-08）。
         "−": "-", "－": "-", "-": "-",
         "＋": "+", "+": "+"}

#: 順序が意味を変える演算（`A/B` と `B/A` は別物）。
ORDERED = ("/", "-")

#: 語の切れ目になる記号。
BREAKS = set("（）()「」『』〔〕[]{}、。,.，．:：;；=＝!！?？'\"…・　 \t\n")

#: 語の切れ目になる助詞（後ろ側）。★ 長いものから試す。
PARTICLES = ("から", "より", "の", "を", "で", "に", "は", "が", "と", "へ", "や", "も")

#: 語の切れ目になる助詞（前側・1 文字だけ）。★ これが無いと「締め日**を**2026/09」を
#: 丸ごと 1 語として読んでしまう。迷う側には**黙る**方へ倒す。
LEAD_PARTICLES = set("のをでにはがとへやも")

_DECL = re.compile(r"演算対象:(.+?) と (\S+)")
_OPER = re.compile(r"演算子:(\S+)")


def _left_word(text: str) -> str:
    """記号の**手前**から、切れ目に当たるまで遡って 1 語取る。"""
    out = []
    for ch in reversed(text):
        if ch in BREAKS or ch in LEAD_PARTICLES:
            break
        out.append(ch)
    return "".join(reversed(out))


def _right_word(text: str) -> str:
    """記号の**後ろ**から、切れ目か助詞に当たるまで 1 語取る。"""
    out = []
    for i, ch in enumerate(text):
        if ch in BREAKS:
            break
        if out and any(text.startswith(p, i) for p in PARTICLES):
            break
        out.append(ch)
    return "".join(out)


def _is_wordy(s: str) -> bool:
    """語らしいか。★ 数字だけ（日付の `2026/09/30` 等）は式と読まない。"""
    return bool(s) and not s.isdigit() and not re.fullmatch(r"[\d.,%]+", s)


def stated_calculation(task: str) -> tuple | None:
    """依頼文が `A÷B` のように**二語を演算記号で挟んだ形**を書いていれば
       `(左, 右, 演算子, 依頼文にあった記号)` を返す。読み取れなければ None（黙る）。

       ★ 記号が 2 つ以上ある回は None ── 三項以上の式（`（売上-原価）÷売上`）も
         日付（`2026/09/30`）もこの形なので、どれを比べるかを機械が決められない。
       ★ 数を数えるのは**語として読めたか判定する前**にする ── 後で数えると、
         片方が読めなかった三項の式が「記号 1 つ」に化けて通ってしまう。"""
    marks = [i for i, ch in enumerate(task or "") if ch in SIGNS]
    if len(marks) != 1:
        return None
    i = marks[0]
    left, right = _left_word(task[:i]), _right_word(task[i + 1:])
    if not (_is_wordy(left) and _is_wordy(right)):
        return None
    return (left, right, SIGNS[task[i]], task[i])


def declared_calculation(declaration: str) -> tuple | None:
    """宣言（解釈行）が `演算対象:X と Y 演算子:Z` を持てば `(X, Y, Z)` を返す。"""
    m, o = _DECL.search(declaration or ""), _OPER.search(declaration or "")
    if not m or not o:
        return None
    return m.group(1).strip(), m.group(2).strip(), o.group(1).strip()


def _same_name(a: str, b: str) -> bool:
    """語の同一視 ── 切り出しに助詞や修飾が残ることがあるので、包含も同じとみなす。
       ★ 緩い側に倒す（迷ったら鳴らさない）。"""
    return a == b or a in b or b in a


def calculation_mismatch(task: str, declaration: str) -> str | None:
    """依頼文が書いた式と、宣言した計算が食い違うなら**依頼側の式**を返す。

    ★ 返すのは「依頼はこう読める」という文字列（例: `利益÷売上`）。
      呼び出し側はそれを見せるだけ ── 直さない・止めない。"""
    asked = stated_calculation(task)
    done = declared_calculation(declaration)
    if not asked or not done:
        return None
    a_left, a_right, a_op, a_sign = asked
    d_left, d_right, d_op = done
    if a_op != d_op:
        return f"{a_left}{a_sign}{a_right}"
    same_order = _same_name(a_left, d_left) and _same_name(a_right, d_right)
    if same_order:
        return None
    if a_op not in ORDERED and _same_name(a_left, d_right) and _same_name(a_right, d_left):
        return None            # ★ 掛け算・足し算は入れ替えても同じ
    return f"{a_left}{a_sign}{a_right}"
