#!/usr/bin/env python
"""段階的に聞く翻訳（**実験用・製品には入っていない**）── 小さいモデル向け。

★★ 2026-08-30（Namakoo）:「座標で精度が上がることが分かっているなら、そちらで 1B に
  対して聞くべきでは？そっちが真の精度だと思う」── そのとおりだった。
  ★ 座標の分（行・列・値）は**既に機械が実表から決めていて LLM に聞いていない**。
    残っていたのは **op の分類だけ**で、そこが 21 択のまま 7B 向けの形だった。
  ★ 実測: 1B は 21 択だと「分かりません」と降りた依頼を、4 択なら当てる（6 件中 5 件）。
    だから 1B の 66.7% は天井ではなく「**7B 向けの問いを 1B に出した**」数字。

三段に割る:
  ① 軸を 4 択で聞く（行／列／セル／並べ替え）
  ② その軸の中で op を 3〜5 択で聞く
  ③ op を固定して引数だけ埋めさせる（translate_task_fixed_op ── 既存）

★ ①②で決められなければ **OUT_OF_VOCAB を返して降りる**（推測で進まない）。
★ 製品の translate_task は 1 文字も変えない ── ここは差し替え用の別実装。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import ailine  # noqa: E402


# --- ① 軸 ---------------------------------------------------------------------------------

# (返り値, 見せるラベル, 説明) ── ★ ラベルは**そのまま書かせる語**（番号にしない）
_AXES = [
    ("行", "行", ("行",)),
    ("列", "列", ("列",)),
    ("セル", "セル", ("セル",)),
    ("並べ替え", "並べ替え", ("並べ替え", "並び替え", "並べ換え", "ソート")),
]

# --- ② 軸ごとの op ------------------------------------------------------------------------
#
# ★★ ラベルは**モデルが実際に使う語**にする。最初は「行を消す」のような説明的な
#   ラベルにしていたが、「削除して」と頼まれたモデルは「削除」と答えるので当たらなかった
#   （段階 B が全滅した）。★ 受け取る綴りも広げる（aliases）。

_OPS_BY_AXIS = {
    "行": [
        ("ADD_ROW", "追加", ("追加", "足す", "加える", "挿入して値")),
        ("INSERT_ROWS", "空行", ("空行", "空の行")),
        ("DELETE_ROWS", "削除", ("削除", "消す", "除く", "取り除")),
        ("SWAP", "入れ替え", ("入れ替え", "入替", "交換")),
    ],
    "列": [
        ("ADD_COLUMN", "追加", ("追加", "足す", "加える")),
        ("COMPUTE_COLUMN", "計算", ("計算", "掛け", "割り", "引い", "足し算")),
        ("DELETE_COLUMN", "削除", ("削除", "消す", "除く", "取り除")),
        ("SWAP", "入れ替え", ("入れ替え", "入替", "交換")),
        ("EXTRACT_COLUMNS", "抜き出し", ("抜き出", "抽出")),
    ],
    "セル": [
        ("SET_CELL_VALUE", "1つ", ("1つ", "一つ", "１つ", "単一")),
        ("SET_COLUMN_VALUE", "列全部", ("列全部", "列ぜんぶ", "まとめて")),
        ("SET_WHERE", "条件つき", ("条件",)),
        ("NUMBER_FORMAT", "桁区切り", ("桁区切り", "書式")),
        ("BOLD", "太字", ("太字", "ボールド")),
    ],
    "並べ替え": [
        ("SORT", "並べ替える", ("並べ",)),
    ],
}


def _ask_choice(model: str, task: str, question: str, options: list) -> str | None:
    """選択肢から 1 つ**名指し**させる。決められなければ None。

    ★★ 2026-08-30（実測・これが効いた）: 最初は番号で答えさせていた。
        番号で聞く: gemma3:1b 2/6 ／ qwen2.5-coder:1.5b 1/6
        単語で聞く: gemma3:1b 5/6 ／ qwen2.5-coder:1.5b 4/6
      両モデルとも、番号だと**中身に関係なく「3」**を返していた。
      ★ 小さいモデルは「番号に写す」のが苦手で、「名指す」のはできる。
        ── これも問いの立て方（この repo で 3 度目）。
    ★ 照合は**長いラベルから**（「行を足す」と「行を消す」が両方当たらないように）。
    """
    # ★★ 2026-08-30（2 つ目の実測）: 選択肢に**説明を付けると壊れる**。
    #     説明なし: gemma3:1b 5/6 ／ qwen2.5-coder:1.5b 4/6
    #     説明あり: gemma3:1b 2/6 ／ qwen2.5-coder:1.5b 1/6
    #   小さいモデルは説明文を**なぞって返す**（「この依頼は、セルの中身を書き換える操…」）。
    #   ★ 短く聞くほど当たる ── 選択肢は**ラベルだけ**を並べる（説明は消す）。
    labels = "／".join(f"「{label}」" for _key, label, _d in options)
    prompt = (f"{question}{labels}のどれか **1 つだけ** で答えてください。\n"
               f"依頼: {task}")
    try:
        raw = str(ailine.ollama_generate(model, [{"role": "user", "content": prompt}],
                                          temperature=0.0))
    except Exception:
        return None
    head = raw.strip()[:60]
    hits = [(len(w), key) for key, _label, aliases in options
             for w in aliases if w in head]
    return max(hits)[1] if hits else None


def translate_task_staged(model: str, task: str, book_meta: dict,
                           temperature: float = 0.1) -> dict:
    """製品の translate_task と同じ形（{"plan": [...]}）を返す、段階版。"""
    # ★★ 2026-08-30（診断で分かった最大の穴）: 「セル」の軸が壊滅していた
    #   ── cell 期待 24 件に対し SET_CELL_VALUE が **0 件**。
    #   理由: 「ナットの棚を東棟にして」は**列名を含む**ので、モデルは「列」と答える。
    #   軸の問いが原理的に曖昧で、聞き方をどう変えても割れない。
    #   ★★ だが機械は既に答えを持っている ── resolve_cell_target_from_task は
    #     依頼文と実表から 1 つのセルを解ける。**解けたならそれはセル操作**。
    #     モデルに聞かない。列挙も 1 つも要らない（表に訊いているだけ）。
    sheet = (book_meta.get("sheets") or [None])[0]
    try:
        hit = ailine.resolve_cell_target_from_task(task, book_meta, sheet)
    except Exception:
        hit = None
    if hit is not None:
        # ★★ 座標は**機械が解いたもの**をそのまま渡す（LLM の row/col は使わない）。
        #   実測: LLM は row に "3" という文字列を返し、「『3』という行が見つかりません」で
        #   落ちていた ── 機械が既に答えを持っているのに、聞き直して壊していた。
        _row, _col_i, _note = hit
        _hdrs = [str(h) for h in ((book_meta.get("headers") or {}).get(sheet) or [])]
        fixed = ailine.translate_task_fixed_op(model, "SET_CELL_VALUE", task, book_meta,
                                                temperature=temperature)
        _args = dict((fixed or {}).get("args") or {})
        _args["row_number"] = _row
        if 1 <= _col_i <= len(_hdrs):
            _args["col"] = _hdrs[_col_i - 1]
        _args.pop("row", None)
        return {"plan": [{"op": "SET_CELL_VALUE", "args": _args}]}

    # ★★ 2026-08-30（測って分かった 2 つ目の穴）: 「3行目を削除して」「3行目の上に
    #   新品を入れて」が **3 表とも**「軸が決まらない」で落ちていた。
    #   依頼文に出てくるのが**行番号だけ**なので、モデルは行とも列とも言えない。
    #   ★★ だが機械なら分かる ── 依頼文が指す行が**実表で解ける**なら、それは行の操作。
    #     セルの時とまったく同じ手（聞く前に表に訊く）。列挙を 1 つも足さない。
    axis = None
    _at = None
    try:
        _at, _note = ailine.resolve_row_anchor(task, book_meta, sheet)
        if _at is None:
            # ★ 「3行目を削除して」── 向きの言葉が無いので resolve_row_anchor は
            #   決めない（そこは正しい）。だが**実在する行番号を名指ししている**なら、
            #   それが行の操作である証拠になる。
            _n = ailine.task_names_a_row_number(task)
            _hr = int((book_meta.get("header_rows") or {}).get(sheet, 1) or 1)
            if _n and _n > _hr:
                axis = "行"
        else:
            axis = "行"
    except Exception:
        pass
    if axis is None:
        axis = _ask_choice(model, task, "次の依頼は、表に対する何の操作ですか。", _AXES)
    if axis is None:
        return {"plan": [{"op": "OUT_OF_VOCAB", "about": "軸が決まらない", "args": {}}]}
    ops = _OPS_BY_AXIS[axis]
    op = None
    # ★★ 「AとBの間にXを作って」が「行の中で決まらない」で落ちていた。
    #   受け取り語に「作る」を足すのは**言い回しの列挙**（Namakoo の指摘した弱点）。
    #   ★ 代わりに証拠で決める: **位置が解けて、置く値も依頼文から取れる**なら、
    #     それは値つきの行を足す依頼。動詞を 1 語も数え上げない。
    if axis == "行" and _at is not None:
        try:
            if (not ailine._re_row_unit.search(task)
                    and ailine.add_row_values_from_request(task, book_meta, sheet, {})):
                op = "ADD_ROW"
        except Exception:
            pass
    if op is None:
        op = ops[0][0] if len(ops) == 1 else _ask_choice(
            model, task, f"次の依頼は、{axis}の操作のうちどれですか。", ops)
    if op is None:
        return {"plan": [{"op": "OUT_OF_VOCAB", "about": f"{axis}の中で決まらない", "args": {}}]}
    fixed = ailine.translate_task_fixed_op(model, op, task, book_meta,
                                            temperature=temperature)
    args = (fixed or {}).get("args") or {}
    return {"plan": [{"op": op, "args": _machine_fix_args(op, args, task, book_meta)}]}


# --- 引数を機械で縛る（★ ここが Namakoo の言う「機械で誘導する」側）------------------------

_ORDER_WORDS = (("desc", ("降順", "大きい順", "多い順", "高い順", "大きな順")),
                 ("asc", ("昇順", "小さい順", "少ない順", "低い順", "小さな順")))


def _machine_fix_args(op: str, args: dict, task: str, book_meta: dict) -> dict:
    """LLM が埋めた引数を、**依頼文と実表**で上書きする（A' 原則の延長）。

    ★★ 2026-08-30 の実測がこの関数の理由:
      ① 「棚の列を削除して」→ col=備考（**実在するが違う列**）。実在するので今の
         救済（実在しない時だけ働く）は発火しない ── **正しいが違う**は素通りしていた。
         ★ 依頼文がちょうど 1 つの実在列を名指ししているなら、そちらが勝つ。
      ② 「数量で降順に並べ替えて」→ order=「降順」（日本語）。モデルは**正しく答えて
         いる**のに schema が asc/desc なので弾いていた ── 受け口の問題。
    """
    out = dict(args)
    sheet = (book_meta.get("sheets") or [None])[0]
    headers = [str(h) for h in ((book_meta.get("headers") or {}).get(sheet) or [])]

    # ① 列は依頼文が名指しするものを優先（実在する別の列を返された時にも効く）
    named = ailine._task_names_single_real_column(task, headers)
    for key in ("col", "target"):
        if named and out.get(key) and str(out[key]) != named:
            out[key] = named

    # ★★ ② 行の位置は依頼文から取る。実測: 「3行目を削除して」で分類は通ったのに
    #   at が空で「行番号『None』が不正です」と落ちていた ── 依頼文に「3行目」と
    #   書いてあるのに、誰も読んでいなかった。
    if op in ("DELETE_ROWS", "INSERT_ROWS", "ADD_ROW"):
        _n = ailine.task_names_a_row_number(task)
        if _n and not str(out.get("at") or "").isdigit():
            out["at"] = _n

    # ③ 昇順・降順は依頼文から取る（モデルの日本語もそのまま受ける）
    if op == "SORT":
        text = str(task) + " " + str(out.get("order") or "")
        for val, words in _ORDER_WORDS:
            if any(w in text for w in words):
                out["order"] = val
                break
    return out


# --- ③ 表の中身を見せる（★ bench の中だけ・製品のプロンプトは触らない）--------------------
#
# ★★ 2026-08-30（Namakoo「1Bについてもセルマップは渡している？」）── 渡していなかった。
#   製品が第二段に見せているのは **見出しだけ**、しかも生の JSON:
#       対象ブックの構成: {"在庫": ["品名", "棚", "数量", "備考"]}
#   ★ シート名と列名が**同じ並びに見える**。1B が col=在庫 を返したのは幻覚ではなく、
#     そう読めるものを渡していたから ── 表記ゆれでもモデルの限界でもない。
#   ★ ここでは (a) 構造を日本語で言う (b) 実際の値を数行見せる、の 2 つを足して測る。

def _book_text(book_meta: dict, rows: int = 3) -> str:
    """シート・列・実際の値を、**取り違えようのない形**で書く。"""
    path = book_meta.get("path")
    out = []
    for sheet, headers in (book_meta.get("headers") or {}).items():
        hs = "／".join(str(h) for h in headers)
        out.append(f"シート『{sheet}』の列は {hs} です。")
        if not path:
            continue
        try:
            import openpyxl
            ws = openpyxl.load_workbook(path, data_only=True)[sheet]
            hr = int((book_meta.get("header_rows") or {}).get(sheet, 1) or 1)
            for r in range(hr + 1, hr + 1 + rows):
                vals = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
                if all(v in (None, "") for v in vals):
                    break
                pairs = "／".join(f"{h}={v}" for h, v in zip(headers, vals)
                                   if v not in (None, ""))
                out.append(f"  {r}行目: {pairs}")
        except Exception:
            pass
    return "\n".join(out)


def _fixed_op_messages_with_content(op: str, task: str, book_meta: dict) -> list:
    system = ailine.TRANSLATION_FIXED_OP_SYSTEM.format(
        op=op, schema=ailine._op_schema_doc(op))
    return [
        {"role": "system", "content": system},
        {"role": "user", "content": f"{_book_text(book_meta)}\n依頼: 「{task}」"},
    ]


def enable_content_prompt() -> None:
    """第二段のプロンプトを『中身つき』に差し替える（実験用）。"""
    ailine.build_translation_fixed_op_messages = _fixed_op_messages_with_content


# --- 差し替えて ailine を動かす（製品のファイルは触らない）--------------------------------

def main() -> int:
    ailine.translate_task = translate_task_staged      # ★ ここだけが実験の実体
    import os
    if os.environ.get("AILINE_CONTENT_PROMPT"):
        enable_content_prompt()                        # ★ 表の中身も見せる（実験）
    return ailine.main(sys.argv[1:])


if __name__ == "__main__":
    raise SystemExit(main())
