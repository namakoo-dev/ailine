"""見出し行の案内（`--header-row N`）を、機械が実際に見た行から組み立てる（2026-10-01）。

★★ なぜ在るか（盲検 7 体目・「嘘の案内」の家系）:
  ① 試算表（タイトル 2 行・空行・見出しは 4 行目）に「`--header-row 3` のように指定して」と
    言っていた ── 3 は**決め打ちの例**で、その冊の 3 行目は空行だった。
  ② 転記で「`--header-row 4` のように指定して再実行してください」と言い、その通りに打つと
    **同じ文で止まった**。`--header-row` が掛かるのは run が最初に選んだシート
    （依頼文が名指しした参照表）で、案内の列が在るのは別のシートだった（`--sheet` も要った）。
★ 案内は「その通りに打てば、その行がそのシートの見出しになる」ものだけ。要るものは 1 行に全部並べる。
★ ailine を import しない（純関数・画面の文は測れる形で持つ）。
"""
from __future__ import annotations

#: 見出しらしい行が 1 つも見えなかった時の問い（★ 行番号を当て推量で出さない）
NO_CANDIDATE_QUESTION = ("見出しが何行目か分かりません（文字の見出しが 2 つ以上並ぶ行が見当たりません）。"
                         "見出しの行番号を `--header-row <行番号>` で指定して再実行してください")


def header_row_candidates(rows: dict) -> list:
    """StructDump の行の特徴（{行(1起点): {"str": 文字のセル数, "nonempty": 非空セル数}}）から、
       見出しらしい行（非空セルが 2 つ以上で、全部が文字）を、幅の広い順・同じ幅なら上から。
       ★ 見出しの検出（detect_header_row）が「一意に決められない」と言った回に使う ──
         ここは決めない。**人に見せる例**を、実際に在る行から選ぶだけ。"""
    rows = rows or {}
    cands = [r for r, info in rows.items()
             if (info or {}).get("str", 0) >= 2 and info.get("nonempty") == info.get("str")]
    return sorted(cands, key=lambda r: (-rows[r]["str"], r))


def clarify_question(rows: dict) -> str:
    """見出し行が決まらない時の問い。例の行番号は、その冊で見出しらしく見えた行の 1 番手。"""
    cands = header_row_candidates(rows)
    if not cands:
        return NO_CANDIDATE_QUESTION
    shown = "・".join(f"{r} 行目" for r in cands[:3])
    return (f"見出しが何行目か分かりません（見出しらしい行: {shown}）。"
            f"`--header-row {cands[0]}` のように指定して再実行してください")


def _arg(v) -> str:
    """打てる形の引数（空白を含むなら引用符で囲む）。"""
    s = str(v)
    return f'"{s}"' if any(c.isspace() for c in s) else s


def missing_column_hint(raw_col: str, row: int, sheet_name: str, applies_to: str | None) -> str:
    """列が見出し行以外の行 row に見つかった時の案内。

    ★ applies_to は「`--header-row` だけを付けた時に、それが掛かるシート」（run が翻訳の前に選んだ対象）。
      列が在るシートと違うなら、`--sheet` も要る ── **両方を 1 行に**並べる（片方だけ言うと、
      従った人がまた同じ文で止まる）。
    """
    if applies_to is not None and applies_to != sheet_name:
        return (f"列『{raw_col}』は『{sheet_name}』シートの{row}行目に見出しがあるようです。"
                f"`--sheet {_arg(sheet_name)} --header-row {row}` を付けて再実行してください"
                f"（--header-row だけだと『{applies_to}』シートに掛かります）")
    return (f"列『{raw_col}』は{row}行目に見出しがあるようです。"
            f"`--header-row {row}` のように指定して再実行してください")
