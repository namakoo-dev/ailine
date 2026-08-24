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

★★ 挙動変更#3（対象シートの取り違え）: 挙動変更#2 の副作用を実測で見つけた ──
`sheets=['売上データ','金額']` のブックへ「金額を降順に並べ替えて」と頼むと、**列を
指したつもりの語**が2枚目『金額』シートの名前と完全一致して、対象シートが2枚目になる
（`sheets=['明細','合計']` + 「合計行を追加して」も同型）。告知は出るので沈黙ではないが、
対象シートにも同名列があれば**間違ったシートを並べ替えて「成功」してしまう** ── エラーで
止まるより後で気づきにくい。そこで:
  - 依頼文で**そのシート名の直後が「シート」「タブ」**（または「N枚目」の序数表現）＝
    明示マーカー付きの言及は、無条件に採用する（衝突チェックを免除）。
  - **裸の言及**は、その語が**どこかのシートの列見出しにも存在する**なら曖昧と判断し、
    既定(1枚目)へ後退する。後退した事実は SheetNameConflict として呼び出し側へ返し、
    呼び出し側が（対話できる場面でだけ）3択で聞く。
★ A' 原則はここでも同じ: 実在シート名・実在列名との**機械照合のみ**で、LLM は使わない。
★ 正直に記録しておく限界: 列見出しは `build_book_meta(source_book)` を引数なしで呼んだ
結果（＝全シート1行目）を使う。タイトル行があって見出しが5行目にあるブックでは衝突を
見逃す（＝挙動変更#2 と同じ挙動になるだけで、退行はしない）。見出し行の確定は対象シートが
決まった後にしかできないため、順番として避けられない（README の「既知の限界」参照）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable

_SHEET_ORDINAL_RE = re.compile(r"(\d+)\s*枚目")
# シート名の直後に来ると「シートを指している」と読める語。閉じ括弧・引用符は挟まってよい
# （「『金額』シート」のような書き方を拾う）。
_SHEET_MARKER_SUFFIX = r"[』」”’\"'）\)]*\s*(?:シート|タブ)"


@dataclass(frozen=True)
class SheetNameConflict:
    """裸の言及が列見出しとも一致したため、既定(1枚目)へ後退したことの記録。

    word:        依頼文中の曖昧な語（＝シート名でもあり、どこかの列名でもある）
    alternative: その語が指しうるもう一方の解釈＝同名のシート（word と同じ文字列だが、
                 「語」と「シート」は別の役割なので別フィールドとして持つ）
    chosen:      実際に採用した既定シート（＝1枚目）
    """
    word: str
    alternative: str
    chosen: str


def conflict_excluded_sheets(conflict: SheetNameConflict | None) -> set:
    """★ 誤爆#3: 衝突で既定へ後退したとき、助言側が「シート言及」から外すべきシート名。

    ★ なぜここ（決めた側）に置くのか: 「その語が曖昧か」は resolve_target_sheet が既に
    決めていて、SheetNameConflict に記録している。助言側（ailine.py の
    mention_overlap_advisory）はそれを読まず、独立に「シート名が依頼文に含まれるか」だけを
    見て「★ 依頼で言及された『金額』は…変更されていません」を出していた ── **同じことを
    2 箇所が別々に決める**形で、後退したのが正しい判断なのに警告だけが誤爆していた。
    助言側で「衝突かどうか」を判定し直すと 3 箇所目が増えるので、判定はここ 1 箇所に置き、
    **決めた側の結果を運ぶ**。
    conflict が None（衝突なし＝そもそも後退していない）なら空集合＝抑制は一切かからない。
    """
    return {conflict.alternative} if conflict else set()


def sheet_names_mentioned_in(task: str, sheets: list) -> list:
    """★ 単位E: 「このシート名は依頼文に含まれるか」という**素材**を1箇所に切り出したもの。

    ★ なぜ: この判定は resolve_target_sheet（決定側）と ailine.py の extract_task_mentions
      （助言側）が**それぞれ独立に同じ文字列照合を書いていた**。規則が片方だけ変わると
      静かにずれる（誤爆#3 で払った授業料と同じ形 ―― 同じことを2箇所が別々に決める）。
      素材はここ1つにして、**決定側は受け取ってから1つに絞る／助言側は全部使う**。
    ★ 戻り値の契約も判定規則も従来と同一（実在シート名が依頼文に部分文字列として現れるか・
      順序はブックのシート順）。ここは切り出しであって挙動変更ではない。"""
    task = task or ""
    return [s for s in (sheets or []) if s and s in task]


def _mentioned_with_marker(task: str, name: str) -> bool:
    """依頼文で name が「〜シート」「〜タブ」の形で言及されているか（明示マーカー）。"""
    return re.search(re.escape(name) + _SHEET_MARKER_SUFFIX, task) is not None


def _is_also_a_column_name(name: str, headers: dict | None) -> bool:
    """name が（どのシートであれ）実在の列見出しと完全一致するか。"""
    for cols in (headers or {}).values():
        for col in cols or []:
            if str(col).strip() == name:
                return True
    return False


def sheets_named_explicitly(task: str, sheets: list) -> list:
    """依頼文が**はっきりシートとして**名指ししているシート名（「〜シート/タブ」or「N枚目」）。

    ★ なぜ「裸の言及」を含めないか（2026-08-24 の実測）: 「売上が60以上の行だけ現場ごとに
      集計して」の『売上』は**列名**であって、同名のシートを指してはいない。裸の言及まで
      「人が指定した」と読むと、複合計画の連鎖が一度も効かなくなる。
      挙動変更#3 の「裸の言及は列名と衝突しうるので曖昧」と同じ線。
    """
    out = []
    for name in sheets or ():
        if name and _mentioned_with_marker(task or "", name):
            out.append(name)
    return out


def drop_names_covered_by_longer(task: str, names: list) -> list:
    """依頼文の中で、**出現位置がすべて別の長い名前の内側にある**名前を落とす。

    ★ なぜ位置で見るのか（2026-08-24 の実測）: 集合だけで「短い方を落とす」と、
      「売上シートと売上60以上シートを見比べて」のように**独立して書かれた短い名前**まで
      消える。逆に畳まないと、「売上60以上シートを集計して」で『売上』を言及と数えて
      「『売上』は変更されていません」という誤った ⚠ が出て、正しくできた仕事が △ に降格する。
      どちらも実測した事故なので、位置で判定する以外に両立しない。
    """
    spans = {}
    for n in names:
        found, start = [], 0
        while True:
            i = task.find(n, start)
            if i < 0:
                break
            found.append((i, i + len(n)))
            start = i + 1
        spans[n] = found
    kept = []
    for n in names:
        mine = spans.get(n) or []
        longer = [o for o in names if o != n and len(o) > len(n)]
        # ★ 落とすのは「出現が1つ以上あり、その**すべて**が、より長い名前の出現の内側に
        #   収まっている」場合だけ。長い候補が無ければ落とさない（all() が空で真になる罠）。
        covered = bool(mine) and bool(longer) and all(
            any(o_s <= a and b <= o_e
                 for o in longer for (o_s, o_e) in spans.get(o, ()))
            for (a, b) in mine
        )
        if not covered:
            kept.append(n)
    return kept or list(names)


def resolve_target_sheet(task: str, sheets: list, cli_sheet: str | None = None,
                          headers: dict | None = None) -> tuple:
    """★ 挙動変更#2: 対象シートの決定はここ1箇所だけで行う（呼び出し側 [_cmd_run_dispatch]
       が戻り値を a._target_sheet に積み、以降は全部それを読む）。
       戻り値: (対象シート名 or None, source∈{"cli","task","default"}, error_or_None,
                conflict∈{SheetNameConflict, None})。
       優先順位: ①--sheet 明示指定（実在しなければエラー・衝突チェックとは無関係に常に最優先）
       ②依頼文中の実在シート名の完全一致（複数一致・部分文字列一致し合う場合は長い方だけ
       残す。それでも複数なら曖昧と判断し、CLARIFY で止めずに④既定へフォールバックする —
       LOOKUP_FILL のように依頼文に転記先/参照元の2シート名が正当に両方登場するケースを
       誤ってブロックしないため。ユーザーが1シートだけを明示したい場合は --sheet を使う）
       ③「N枚目」の序数表現 ④既定=1枚目（旧挙動と同一）。
       ★★ 挙動変更#3: ②で1つに絞れた言及が**裸**（「シート」「タブ」が後ろに無い）で、
       かつその語が headers のどこかの列見出しとも一致する場合は、②を採らずに④へ後退し、
       4つ目の戻り値 SheetNameConflict で「後退した」ことを呼び出し側へ伝える
       （headers が渡されない/列が取れない場合は挙動変更#2 のまま＝退行させない）。
       ★ A' 原則: 値は LLM に確定させない。実在するシート名/序数との機械的な照合のみ
       （LOOKUP_FILL が既に source_sheet を名前で受けている仕組みと同じ考え方）。"""
    if not sheets:
        return None, "default", "ブックにシートが無い", None
    if cli_sheet:
        if cli_sheet not in sheets:
            return None, "cli", f"シート『{cli_sheet}』がありません。あるシート: {', '.join(sheets)}", None
        return cli_sheet, "cli", None, None
    task = task or ""
    named = sheet_names_mentioned_in(task, sheets)   # ★ 単位E: 照合の素材は1箇所（上記）
    named = [s for s in named if not any(s != t and s in t for t in named)]   # 部分文字列は長い方だけ残す
    # ★★ 明示マーカーは裸の言及に勝つ（2026-08-24 の実測）: ブックに『集計』シートが在ると、
    #   依頼文の**動詞**「集計して」がシート名と一致して言及が 2 つになり、はっきり
    #   「売上60以上シートを」と書いた指定まで曖昧扱いで既定へ落ちていた。
    #   裸の言及は一般語との偶然の一致でありうるが、「〜シート/タブ」と書かれた言及は
    #   人が意図して書いたものなので、こちらを優先する（挙動変更#3 の「明示マーカー付きは
    #   無条件に採用」を、単数のときだけでなく**絞り込みの段階**にも効かせる）。
    #   ★ マーカー付きが 2 つ以上あるときは今までどおり曖昧のまま（LOOKUP_FILL の
    #   転記先/参照元のように両方が正当な場面で、勝手に片方へ寄せない）。
    if len(named) > 1:
        marked = [s for s in named if _mentioned_with_marker(task, s)]
        if len(marked) == 1:
            named = marked
    if len(named) == 1:
        name = named[0]
        # ★ 挙動変更#3: 明示マーカー付きの言及は無条件に採用（衝突チェックを免除）。
        if _mentioned_with_marker(task, name) or not _is_also_a_column_name(name, headers):
            return name, "task", None, None
        conflict = SheetNameConflict(word=name, alternative=name, chosen=sheets[0])
        return sheets[0], "default", None, conflict
    m = _SHEET_ORDINAL_RE.search(task)
    if m:
        idx = int(m.group(1)) - 1
        if 0 <= idx < len(sheets):
            return sheets[idx], "task", None, None
    return sheets[0], "default", None, None


def describe_target_sheet(sheets: list, target_sheet: str | None, source: str) -> str | None:
    """★ 挙動変更#2 最低限: 適用前に対象シートを明示する（査定所見:「これがあれば事故は
       防げた」）。1枚しか無いブックは曖昧さが無いため沈黙する（既存の単一シート帳票の
       出力を一切変えない＝ゴールデンへの影響ゼロ）。
       ★★ 挙動変更#3: 3種とも「操作するシート:」で始める ── 旧文言は前置き
       （「このブックは4シートですが、」）から入っていて、**一番大事な情報が文の後ろ**に
       あった。衝突で既定へ後退した場合も文言は既定と同じ（後退したこと自体は、この行では
       なく3択で伝える）。"""
    if not target_sheet or len(sheets) <= 1:
        return None
    idx = sheets.index(target_sheet) + 1 if target_sheet in sheets else None
    ordinal = f"{idx}枚目" if idx else ""
    if source == "cli":
        return f"操作するシート: {ordinal}『{target_sheet}』（--sheet 指定）"
    if source == "task":
        return f"操作するシート: {ordinal}『{target_sheet}』（依頼文から判断・このブックは{len(sheets)}シート）"
    return f"操作するシート: {ordinal}『{target_sheet}』（このブックは{len(sheets)}シート）"


def format_sheet_field(sheets: list, target_sheet: str | None) -> str | None:
    """★ 挙動変更#3: 「解釈:」行の先頭に載せるシート欄
       （例: `シート:『売上データ』(1枚目)`）。

    ★ なぜ: 転記(LOOKUP_FILL)だけが確認行に「対象シート:」を持ち、他の操作は持って
      いなかった。シートを選べるようになった以上、**確かめる行にシートが無いのは片手落ち**。
    ★ 1枚だけのブックでは None（沈黙）── 単一シート帳票の出力はバイト単位で不変。"""
    if not sheets or len(sheets) <= 1 or not target_sheet or target_sheet not in sheets:
        return None
    return f"シート:『{target_sheet}』({sheets.index(target_sheet) + 1}枚目)"


# ★ op → (辞書形, タ形)。3択の文（「…を並べ替える」「…を並べ替えた場合を」）は日本語の
#   活用が要るため、表示ラベル（OP_LABELS の「並べ替え」等）からは機械的に作れない。
#   ここに無い op は `ラベル+する / ラベル+した`（転記する/集計する のように自然に読める形）。
#   ★ 置き場所: ailine.py の OP_LABELS の隣ではなくここ。この表を使うのは3択の文だけで、
#     ailine.py 側は op と表示ラベルを渡すだけにする（新しいロジック本体は ailine_core/ に）。
_OP_VERBS = {
    "SORT": ("並べ替える", "並べ替えた"),
    "COMPUTE_COLUMN": ("計算列を作る", "計算列を作った"),
    "BOLD": ("太字にする", "太字にした"),
    "FILL_COLOR": ("背景色を付ける", "背景色を付けた"),
    "NUMBER_FORMAT": ("数値書式を設定する", "数値書式を設定した"),
    "CHART": ("グラフにする", "グラフにした"),
    "CENTER_ALIGN": ("中央揃えにする", "中央揃えにした"),
    "APPEND_TOTAL": ("合計行を追加する", "合計行を追加した"),
    "INSERT_ROWS": ("行を挿入する", "行を挿入した"),
    "DRAW_BORDERS": ("けい線を引く", "けい線を引いた"),
    "PIVOT": ("ピボットにする", "ピボットにした"),
    "SET_COLUMN_VALUE": ("一括で書き換える", "一括で書き換えた"),
    "AUTOFIT": ("列幅を自動調整する", "列幅を自動調整した"),
    "EXTRACT": ("抽出する", "抽出した"),
    "SPLIT_CELL": ("セルを分割する", "セルを分割した"),
    "DEDUP": ("重複を除去する", "重複を除去した"),
}


def op_verbs(op: str, op_label: str) -> tuple:
    """(辞書形, タ形)。未登録の op はラベルに する/した を付ける（上の表のコメント参照）。"""
    return _OP_VERBS.get(op) or (f"{op_label}する", f"{op_label}した")


def sheet_conflict_choice_lines(conflict: SheetNameConflict, op: str, op_label: str) -> tuple:
    """3択の前置き行と選択肢を組む（純関数・印字はしない）。

    戻り値: (前置きの行リスト, [(key, 文), ...])。
    ★ なぜ翻訳の**後**で聞くのか（設計判断・変えないこと）: 翻訳前は操作が決まっておらず
      「『売上データ』に対して実行します」という中身のない選択肢しか出せない。それでは
      「よく分からないけど1を押す」になる。**操作が確定してから、具体的な日本語で選ばせる。**
      原本にはまだ触れていないので安全。"""
    dict_form, ta_form = op_verbs(op, op_label)
    lines = ["", f"依頼文の「{conflict.word}」は2通りに読めます"
                 f"（『{conflict.alternative}』という名前のシートもあるため）:"]
    choices = [
        ("1", f"『{conflict.chosen}』シートの「{conflict.word}」列を{dict_form}"
              f" ← 上の解釈のとおり実行する"),
        ("2", f"『{conflict.alternative}』シートを{ta_form}場合を見てみる"),
        ("3", "やめる"),
    ]
    return lines, choices


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
