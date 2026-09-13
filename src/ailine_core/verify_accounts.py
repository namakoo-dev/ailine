# -*- coding: utf-8 -*-
"""科目の候補の冊（`ailine accounts` の出力）を**後から・ファイルだけから**独立に検算する
   （需要③・2026-09-13）。設計 docs/DESIGN-20260913-経費の勘定科目を先例から引く.md §6.5

## 何を「独立」と呼ぶか（★ ここを外すと恒真になる）

**鍵こそが規則**なので、検算側で鍵を作り直したら恒真になる（`plan_accounts` を呼ばない ──
`tests/test_verify_accounts.py` が AST で縛っている）。代わりに測るのは:

    ① 出力が書いた**先例の番地**（過去のファイル名 ＋ 行番号）のセルを**読むだけ**で、
       候補の科目が本当にそこに在るか（規則を再現しない・捏造も同時に落ちる）
    ② 分母は**今回の入力**から取る: 借方勘定科目が空で借方金額が在る行 ＝
       候補が出た行 ＋ 検分が理由を書いた行、重複なく 1 回
    ③ 区分と番地の辻褄（裏が取れたと名乗る行は、別々の鍵を 2 本以上引いているか）

★★ 検分が書いた「N 件」は**読まない** ── 書き手の主張を分母にしたら検算にならない。
  この 1 点は「検分の件数を書き換えても判定が変わらない」ことで縛る。
★ 読み手（CSV / xlsx → 行の並び）は書き手と**同じ** `accounts_read` を使う ──
  この器官の独立性は「読み手の独立」ではなく「**規則を再現しないこと**」で担保する
  （`verify_forms` と同じ線）。自分の出力（冊）の側は、書き手（openpyxl）とは別実装の
  `xml_readback` で読む。
★ 「正しい科目を選んだか」は測らない ── それは規則の再現（設計 D5 の★）。
"""
from __future__ import annotations

from ailine_core import accounts_core, accounts_read, field_record, form_read, stack
from ailine_core import xml_readback


def _sheet_rows(out_path, sheet_name: str) -> tuple:
    """出力の 1 枚を ({見出し: 列}, [{列: 値}, ...]) で読む（★ `xml_readback` ── 別実装）。

    ★ 目当てのシートが無いと `read_grid` は 1 枚目へ落ちる ── 落ちたら空で返す
      （黙って別のシートを読んで「破れ 0」を名乗らないため）。
    """
    data = xml_readback.read_grid(out_path, sheet_name)
    if data.get("sheet_fallback") or data.get("sheet_name") != sheet_name:
        return {}, []
    grid = data.get("grid") or {}
    rows = {}
    for (r, c), v in grid.items():
        rows.setdefault(r, {})[c] = v
    if not rows:
        return {}, []
    head = min(rows)
    names = {str(v).strip(): c for c, v in rows[head].items() if str(v or "").strip()}
    body = [cells for r, cells in sorted(rows.items()) if r > head]
    return names, body


def _int_or_none(value):
    text = str(value if value is not None else "").strip()
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and float(value).is_integer():
        return int(value)
    return int(text) if text.isdigit() else None


def _past_books(past_paths) -> tuple:
    """過去の冊を名前で引ける形にする。戻り値 ({名前: 冊}, [読めなかった名前]）。"""
    books, unreadable = {}, []
    for path in past_paths:
        book = accounts_read.read_journal(path)
        if book.refused:
            unreadable.append(f"{book.name}（{book.refused}）")
            continue
        books[book.name] = book
    return books, unreadable


def _cell_at(book, row_num: int, role: str):
    """その冊の (行, 役割の列) のセルと、その行が在ったか。★ 行の不在と空セルを分ける
       （出ないことは信号でない ── 「行が無い」を「空だった」と読むと捏造が通る）。"""
    column = (book.header_map or {}).get(role)
    for r, values in book.rows:
        if r == row_num:
            return accounts_core.cell(values, column), True
    return None, False


def verify_accounts_book(out_path, today_path, past_paths) -> dict:
    """候補の冊を、今回の入力と過去の冊に突き合わせる。

    戻り値: {"breaks": [(名前, 名指し)], "facts": {...}, "mismatch": bool} または
            {"unsupported": 理由}（検算できない形 ── 空虚な合格を名乗らない）。
    """
    creator, _description = xml_readback.read_core_properties(out_path)
    if creator != accounts_core.CREATOR_MARK:
        return {"unsupported": f"この冊に `{accounts_core.CREATOR_MARK}` の印がありません"
                               f"（作成: {creator!r}）── 人のファイルを検算したことにしない。"}
    names, body = _sheet_rows(out_path, accounts_core.SHEET_NAME)
    if not names:
        return {"unsupported": f"候補のシート（{accounts_core.SHEET_NAME}）が読めません: "
                               f"{out_path}"}
    wanted = (stack.PROVENANCE_HEADERS[1],) + accounts_core.OUTPUT_HEADERS
    missing = [n for n in wanted if n not in names]
    if missing:
        return {"unsupported": f"候補のシートに必要な列がありません: {missing}"}
    i_row, i_account, i_grade, _i_reason, i_cite = (names[n] for n in wanted)

    today = accounts_read.read_journal(today_path)
    if today.refused:
        return {"unsupported": f"今回の冊が読めません: {today.refused}"}
    # ★★ 分母は**今回の入力**から（設計 §6.5 の 2）── 出力の行数でも検分の主張でもない。
    denominator = set()
    for row_num, values in today.rows:
        account = accounts_core.cell(values, today.header_map.get(accounts_core.DEBIT_ACCOUNT))
        amount = accounts_core.cell(values, today.header_map.get(accounts_core.DEBIT_AMOUNT))
        if not form_read.norm(account) and form_read.norm(amount):
            denominator.add(row_num)

    breaks = []
    valued, graded, cited = {}, {}, {}
    for cells in body:
        row_num = _int_or_none(cells.get(i_row))
        if row_num is None:
            breaks.append(("★ 元行が数でない", f"{accounts_core.SHEET_NAME}: "
                                               f"{cells.get(i_row)!r}"))
            continue
        account = str(cells.get(i_account) or "").strip()
        graded[row_num] = str(cells.get(i_grade) or "").strip()
        cited[row_num] = accounts_core.parse_citations(cells.get(i_cite))
        if account:
            valued[row_num] = account

    reasoned, report_rows_seen = _blank_reasoned(out_path)

    # ① 分母の突き合わせ（★ 候補が出た行 ＋ 理由を書いた行 ＝ 借方が空で金額が在る行）。
    for row_num in sorted(denominator - (set(valued) | reasoned)):
        breaks.append(("★ 候補も理由も無い行", f"今回の {row_num} 行目"
                                              "（借方が空で金額が在る行なのに、候補も"
                                              "空欄の理由も出ていません）"))
    for row_num in sorted((set(valued) | reasoned) - denominator):
        breaks.append(("★ 宣言外の行に候補または理由",
                       f"今回の {row_num} 行目（借方が空で金額が在る行ではありません）"))
    for row_num in sorted(set(valued) & reasoned):
        breaks.append(("★ 候補と空欄の理由が同じ行に在る", f"今回の {row_num} 行目"))

    # ② 番地のセルを読むだけ（★ 規則を再現しない）。
    books, unreadable = _past_books(past_paths)
    checked = unchecked = 0
    for row_num in sorted(valued):
        citations = cited.get(row_num) or []
        if not citations:
            breaks.append(("★ 科目を出しているのに先例の番地が無い",
                           f"今回の {row_num} 行目 → {valued[row_num]}"))
            continue
        for key, name, cite_row in citations:
            if name is None or cite_row is None:
                breaks.append(("★ 先例の番地が読めない", f"今回の {row_num} 行目 → {key!r}"))
                continue
            book = books.get(name)
            if book is None:
                unchecked += 1
                continue
            found, exists = _cell_at(book, cite_row, accounts_core.DEBIT_ACCOUNT)
            if not exists:
                breaks.append(("★ 先例の番地に行が無い（捏造）",
                               f"{name} の {cite_row} 行目（今回の {row_num} 行目の根拠）"))
                continue
            if form_read.norm(found) != form_read.norm(valued[row_num]):
                breaks.append(("★ 先例の番地にその科目が在りません（捏造）",
                               f"{name} の {cite_row} 行目は {found!r} ── "
                               f"今回の {row_num} 行目は {valued[row_num]!r} を"
                               "根拠にしています"))
                continue
            checked += 1

    # ③ 区分と番地の辻褄（★ 裏が取れたと名乗る行は、別々の鍵を 2 本以上引いている）。
    for row_num, grade in sorted(graded.items()):
        if grade != field_record.CONFIRMED:
            continue
        keys = {key for key, _name, _row in (cited.get(row_num) or []) if key}
        if len(keys) < 2:
            breaks.append(("★ 区分と番地が食い違う",
                           f"今回の {row_num} 行目: {grade} と書いてあるのに、"
                           f"引いている鍵が {sorted(keys)} しかありません"))

    # ④ 割 と 無 の行に科目が出ていないこと（★ `field_record` の凍結した判断の側から測る）。
    for row_num, grade in sorted(graded.items()):
        if grade and grade not in field_record.GRADES_WITH_VALUE and row_num in valued:
            breaks.append(("★ 値を出さない区分なのに科目が在る",
                           f"今回の {row_num} 行目: {grade} / {valued[row_num]}"))

    facts = {"今回の入力（借方が空で金額が在る行）": len(denominator),
             "候補が出た行": len(valued),
             "検分が理由を書いた行": len(reasoned),
             "検分の行（種類だけ読みました・件数は読んでいません）": report_rows_seen,
             "番地を読んで科目が在ることを確かめた件数": checked,
             "過去の冊": len(books)}
    if unchecked:
        # ★ 出ないことを合格の証拠にしない ── 確かめられなかったものは数えて出す。
        facts["確かめられなかった番地"] = (f"{unchecked} 件 ── 名指しされた過去の冊が"
                                          "渡されていません")
    if unreadable:
        facts["こちらでも読めなかった過去の冊"] = f"{len(unreadable)} 件（{', '.join(unreadable[:3])}）"
    return {"breaks": breaks, "facts": facts, "mismatch": bool(breaks)}


def _blank_reasoned(out_path) -> tuple:
    """検分が「空欄の理由」を書いた行番号と、検分の行数。

    ★ **名指し（種類と行番号）だけ**を読む ── 書かれた**件数は読まない**（書き手の主張を
      分母にしたら検算にならない）。この 1 点は「件数を書き換えても判定が変わらない」で縛る。
    """
    names, body = _sheet_rows(out_path, accounts_core.REPORT_SHEET)
    kind_col = names.get(accounts_core.REPORT_HEADERS[0])
    row_col = names.get(accounts_core.REPORT_HEADERS[1])
    if kind_col is None or row_col is None:
        return set(), len(body)
    out = set()
    for cells in body:
        if str(cells.get(kind_col) or "").strip() != accounts_core.BLANK_KIND:
            continue
        row_num = _int_or_none(cells.get(row_col))
        if row_num is not None:
            out.add(row_num)
    return out, len(body)
