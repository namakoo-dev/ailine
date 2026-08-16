"""target_sheet — C10（宣言つき挙動変更 #2）: 対象シートの決定を一箇所に閉じ込める。

★ なぜ: 独立監査が挙げた散在（verify_dsl_args の `first_sheet = sheets[0]`・codegen の
`getByIndex(0)` が9箇所・事後条件の `sheetnames[0]` が12箇所…）は、すべて「対象シートは
常にブックの1枚目」という同じ暗黙の前提の別々の現れだった。決定そのものをここへ寄せ、
ailine.py 側は resolve_target_sheet()/describe_target_sheet()/wrap_basic_for_sheet() を
呼ぶだけにする（verify_dsl_args の resolved["_target_sheet"]・OP_WRITE_TARGET の
sheet_key=None フォールバック・postcondition チェッカーは、みなその結果を読むだけ）。

★ 移植可能性（tests/test_line_budget.py が機械で守る）: ailine を import しない。
wrap_basic_for_sheet() は「対象シートが1枚目のときの既定ラップ」を呼び出し側から
コールバックで受け取る（ailine.py 側の _wrap_basic をそのまま渡す）ことで、
ailine_core → ailine への逆流を避けている。
"""
from __future__ import annotations

import re
from typing import Callable

_SHEET_ORDINAL_RE = re.compile(r"(\d+)\s*枚目")


def resolve_target_sheet(task: str, sheets: list, cli_sheet: str | None = None) -> tuple:
    """★ 挙動変更#2: 対象シートの決定はここ1箇所だけで行う（呼び出し側 [_cmd_run_dispatch]
       が戻り値を a._target_sheet に積み、以降は全部それを読む）。
       戻り値: (対象シート名 or None, source∈{"cli","task","default"}, error_or_None)。
       優先順位: ①--sheet 明示指定（実在しなければエラー） ②依頼文中の実在シート名の
       完全一致（複数一致・部分文字列一致し合う場合は長い方だけ残す。それでも複数なら
       曖昧と判断し、CLARIFY で止めずに③既定へフォールバックする — LOOKUP_FILL のように
       依頼文に転記先/参照元の2シート名が正当に両方登場するケースを誤ってブロックしない
       ため。ユーザーが1シートだけを明示したい場合は --sheet を使う）
       ③「N枚目」の序数表現 ④既定=1枚目（旧挙動と同一）。
       ★ A' 原則: 値は LLM に確定させない。実在するシート名/序数との機械的な照合のみ
       （LOOKUP_FILL が既に source_sheet を名前で受けている仕組みと同じ考え方）。"""
    if not sheets:
        return None, "default", "ブックにシートが無い"
    if cli_sheet:
        if cli_sheet not in sheets:
            return None, "cli", f"シート『{cli_sheet}』がありません。あるシート: {', '.join(sheets)}"
        return cli_sheet, "cli", None
    task = task or ""
    named = [s for s in sheets if s and s in task]
    named = [s for s in named if not any(s != t and s in t for t in named)]   # 部分文字列は長い方だけ残す
    if len(named) == 1:
        return named[0], "task", None
    m = _SHEET_ORDINAL_RE.search(task)
    if m:
        idx = int(m.group(1)) - 1
        if 0 <= idx < len(sheets):
            return sheets[idx], "task", None
    return sheets[0], "default", None


def describe_target_sheet(sheets: list, target_sheet: str | None, source: str) -> str | None:
    """★ 挙動変更#2 最低限: 適用前に対象シートを明示する（査定所見:「これがあれば事故は
       防げた」）。1枚しか無いブックは曖昧さが無いため沈黙する（既存の単一シート帳票の
       出力を一切変えない＝ゴールデン/933テストへの影響ゼロ）。"""
    if not target_sheet or len(sheets) <= 1:
        return None
    idx = sheets.index(target_sheet) + 1 if target_sheet in sheets else None
    ordinal = f"{idx}枚目の" if idx else ""
    if source == "cli":
        return f"操作対象シート: {ordinal}『{target_sheet}』（--sheet 指定）"
    if source == "task":
        return f"このブックは{len(sheets)}シートですが、操作対象は{ordinal}『{target_sheet}』です（依頼文の言及から判断）"
    return f"このブックは{len(sheets)}シートですが、操作対象は{ordinal}『{target_sheet}』です"


def wrap_basic_for_sheet(body: str, wrap_default: Callable[[str], str],
                          sheets: list, target_sheet: str | None) -> str:
    """★ 挙動変更#2: 対象シートの決定を codegen 側でも「一箇所」に寄せる実装点。
       body（と body が Call するヘルパ helpers/AiLineHelpers.bas の12箇所）は
       oDoc.Sheets.getByIndex(0) を「対象シート」として書く前提のまま一切変更しない
       （ailine.py 側 inline Basic の9箇所も同様）。target_sheet がブックの1枚目でない
       場合だけ、Sub Run の中で対象シートを一時的に先頭(index 0)へ移動してから body を
       実行し、実行後に元の位置へ戻す — getByIndex(0) が指す先を実行時に差し替えることで、
       body 自身にもヘルパにも「対象シートはどこか」を教える改修を一切要らなくする。
       対象シートが元から1枚目（従来どおり・既存ゴールデンの大半）なら wrap_default(body)
       と完全に同じ出力（挙動不変・バイト単位で無傷）。
       body には `_scan_last_row_basic` 由来の `Exit Sub`（データ0行なら何もしない）が
       含まれることがあるため、body を別 Sub(__AilineTargetBody)へ切り出して Call する
       ── Exit Sub は呼ばれた側の Sub だけを終了させるので、Run 側の「元の位置へ戻す」行は
       途中終了があっても必ず実行される（元の位置へ戻し損ねてシート順を壊したまま
       保存する事故を防ぐ）。
       wrap_default: 対象シートが1枚目のときに使う、呼び出し側(ailine.py)の既定ラップ関数
       （_wrap_basic をそのまま渡す・ailine_core → ailine の逆流を避けるための注入）。"""
    if not target_sheet or not sheets or target_sheet not in sheets or target_sheet == sheets[0]:
        return wrap_default(body)
    orig_idx = sheets.index(target_sheet)
    esc = target_sheet.replace('"', '""')
    return (
        "Option VBASupport 1\nOption Explicit\n\n"
        "Sub Run(oDoc As Object)\n"
        f'    oDoc.Sheets.moveByName("{esc}", 0)\n'
        "    Call __AilineTargetBody(oDoc)\n"
        f'    oDoc.Sheets.moveByName("{esc}", {orig_idx})\n'
        "End Sub\n\n"
        "Sub __AilineTargetBody(oDoc As Object)\n"
        + body +
        "End Sub\n"
    )
