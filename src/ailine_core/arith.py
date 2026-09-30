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


# ★★ 2026-10-01（依頼の項の台帳で D だった項目）: 2 列の計算の**演算子**は LLM だけが
#   決めていた。「売上から原価を引いた利益の列」に `+` が返ると、事後条件は宣言どおりの
#   式（売上+原価）を確かめて ✓ を出す ── 上の記号の読み手は `A÷B` の形しか読まないので、
#   語で書いた依頼（引いた・掛けた・割った）は誰も見ていなかった。
#   ★ 読み手は 1 本: 記号は上の stated_calculation をそのまま使い、語はこの表で読む。
#   ★ 語は**列挙で増やさない**（足す前に、何を奪うかを測る ── sort_direction と同じ約束）。
#     誤読の元を先に外してある: 不足/満足（足）・取引/値引き/引き継ぐ（引）・差し替え（差）・
#     比較/比べ（比）・減価（減）。
#   ★ 列の名前（『差額』『利益率』）の中の語は数えない ── 呼び出し側が見出しを渡す。
OPERATOR_WORDS = (
    (r"(?<![不満補])足[しすさ]", "+"),
    (r"加え", "+"),
    (r"加算", "+"),
    (r"合計", "+"),
    (r"合算", "+"),
    (r"の和", "+"),
    (r"(?<![取値])引[いく]", "-"),
    (r"引き算", "-"),
    (r"差(?!し[替換込])", "-"),
    (r"減(?:算|ら|じ)", "-"),
    (r"掛[けかる]", "*"),
    (r"かけ[たてる]", "*"),
    (r"倍", "*"),
    (r"の積", "*"),
    (r"乗じ", "*"),
    (r"割[っりるれ]", "/"),
    (r"割合", "/"),
    (r"比(?![較べ])", "/"),
    (r"率", "/"),
)


def read_operator(task: str | None, names=()) -> tuple | None:
    """依頼文から 2 列の計算の演算子を読む。戻りは (演算子, 根拠の語)。演算子は + - * /。

    ★ 読めない時は None（推測しない）: 語も記号も無い／2 種類以上の演算が読める
      （「売上から原価を引いた利益率」は - と / ── どちらか決めない）。
    ★ names（列の名前）の中の語と記号は数えない。
    """
    text = task or ""
    if not text:
        return None
    found: dict = {}
    masked = text
    for n in sorted({str(x) for x in names if x}, key=len, reverse=True):
        masked = masked.replace(n, "・")
    for pat, op in OPERATOR_WORDS:
        m = re.search(pat, masked)
        if m:
            found.setdefault(op, m.group(0))
    # ★ 記号は上の読み手（`A÷B` の形）に任せる ── 記号を持つ列名が依頼に出ている回は読まない。
    if not any(n in text and any(s in n for s in SIGNS) for n in map(str, names or ())):
        stated = stated_calculation(text)
        if stated:
            found.setdefault(stated[2], stated[3])
    if len(found) != 1:
        return None
    op, word = next(iter(found.items()))
    return op, word


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
