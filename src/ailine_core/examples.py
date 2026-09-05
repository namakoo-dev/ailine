"""導線に出す例は、**機械が通ると確かめたものだけ**（2026-09-05）。

★★ 実測した事故: 「整えて」と頼むと道具はこう返していた ──

    「整える」とは具体的に何をしますか（例: けい線を引く／列幅を合わせる／**太字にする**）

  そのまま「太字にして」と打つと **？ 対象『all』は 太字 では未対応です**。
  ★ **道具が自分で示した例を、自分で断っていた。**
  Namakoo:「間違った方向に誘導されると体験を損なう」── 導線が嘘なら、
  導線が無いより悪い。言い直してまた断られる体験になる。

★ 真因は 2 段:
  ① 例文が **few-shot に書かれた作文**で、LLM がそれを写して返していた
  ② 語彙表の `synonyms`（太字・ボールド・強調）も**そのままでは通らない** ──
     実測 10 件中 4 件が断られた。分かれ目は「**対象が要る op かどうか**」で、
     BOLD は「どこを」が要る。`synonyms` は「その op を指す語」であって
     「そのまま打てる依頼文」ではなかった。

★ だから 3 つ目の役目を立てる ── **例文（example_task）**。
    match_phrases … 照合用（断片でよい）
    synonyms      … 表示用の呼び名（3 語まで）
    example_task  … ★ **そのまま打てば通る 1 文**

★★ 規則: **ここに書く文は、実機の番人が毎回「通る」ことを確かめる。**
  （tests/test_examples_actually_work.py）。作文でなく**実測に裏打ちされた例**にする。
  ★ 7473 行の過去の判断「『こう言えば通る』と書くと嘘になるので弱める」を、
    弱めるのではなく**本当にする**側で解いた。

★ 汎用の形にしてある:
  ・例が無い op は**黙る**（無い例を発明しない）
  ・呼び出し側は 1 本の関数を通す（断り・CLARIFY・提案で同じ文面になる）
  ・番人は「例を持つ op すべて」を回すので、足した瞬間から縛られる
"""
from __future__ import annotations

#: op → そのまま打てば通る依頼文。★ 実機の番人が通ることを確かめている。
#: ★ 追加するときは番人を走らせること（通らない文はここに置けない）。
EXAMPLE_TASKS = {
    # --- 見た目（曖昧な依頼から提案されやすい）------------------------------
    "DRAW_BORDERS": "けい線を引いて",
    "AUTOFIT": "列幅を自動調整して",
    "BOLD": "見出しを太字にして",              # ★ 「太字」だけでは通らない（対象が要る）
    "CENTER_ALIGN": "中央揃えにして",
    "FILL_COLOR": "見出しに背景色を付けて",
    "NUMBER_FORMAT": "金額に桁区切りを付けて",
    # --- 表を編集する ------------------------------------------------------
    "SORT": "金額の大きい順に並べ替えて",
    "SET_COLUMN_VALUE": "備考の列を全部「確認済」に書き換えて",
    "SET_WHERE": "金額が1000以上の行の備考に「○」を付けて",
    "DELETE_ROWS": "3行目を削除して",
    "DELETE_COLUMN": "備考の列を削除して",
    "ADD_COLUMN": "区分という列を追加して",
    "APPEND_TOTAL": "金額の合計を一番下に追加して",
    # --- 新しい表を作る ----------------------------------------------------
    "AGGREGATE": "部門ごとに金額をまとめて",
    "EXTRACT": "金額が1000以上の行を抜き出して",
    "DEDUP": "品名が同じ行を重複として除いて",
}


def example_task_for(op: str | None) -> str | None:
    """その op の「そのまま打てば通る」例（無ければ None ── 発明しない）。"""
    return EXAMPLE_TASKS.get(op or "")


def render_example_line(op: str | None, label: str | None = None) -> str | None:
    """導線の 1 行。例が無ければ None（黙る）。

    ★ 文面をここ 1 箇所に置く ── 断り・CLARIFY・もしかして提案で同じ言い方にする
      （呼び出し側に書き写さない）。
    """
    ex = example_task_for(op)
    if not ex:
        return None
    name = f"『{label}』" if label else ""
    return f"  {name}はこう頼めます: 「{ex}」"


def render_examples_for(ops, labels: dict | None = None, limit: int = 3) -> list:
    """複数の op について、例を持つものだけを並べる（曖昧な依頼への提案用）。

    ★ 例が無い op は**列から落とす** ── 「けい線を引く／列幅を合わせる／太字にする」の
      ように、通らない言い方を混ぜない。
    """
    labels = labels or {}
    out = []
    for op in ops:
        ex = example_task_for(op)
        if not ex:
            continue
        out.append(f"「{ex}」")
        if len(out) >= limit:
            break
    return out

#: 曖昧な依頼に対して「まずこれが頼めます」と見せる既定の並び。
#: ★ 見た目の 3 つ（引数が要らない／要っても例が確かめてある）を先頭に置く。
DEFAULT_SUGGESTIONS = ("DRAW_BORDERS", "AUTOFIT", "BOLD")

_EXAMPLE_PAREN = __import__("re").compile(r"（例[:：][^）]*）")


def replace_examples_in_question(question: str, ops=None) -> str:
    """聞き返し文の中の「（例: …）」を、**実測で通る例**に差し替える。

    ★★ なぜ要るか（2026-09-05 実測）: 聞き返しの文は LLM が書いており、その中の例は
      few-shot に書かれた**作文**だった。「整えて」への例に「太字にする」が入っていて、
      そのまま打つと「対象『all』は 太字 では未対応です」で断られた ──
      **道具が自分の示した例を自分で断っていた。**

    ★ 文そのものは書き換えない（聞き返しの主文は LLM の方が場面に合う）。
      **例の括弧だけ**を機械が持つ例に置き換える。例が 1 つも無ければ括弧ごと落とす
      （通らない例を残すより、例が無い方がまし）。
    """
    if not question:
        return question
    examples = render_examples_for(ops or DEFAULT_SUGGESTIONS)
    if not examples:
        return _EXAMPLE_PAREN.sub("", question).strip()
    return _EXAMPLE_PAREN.sub("（例: " + "／".join(examples) + "）", question).strip()

