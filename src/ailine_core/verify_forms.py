# -*- coding: utf-8 -*-
"""帳票の一覧（`ailine forms` の出力）を**後から・ファイルだけから**独立に検算する（2026-09-12）。

## なぜ在るか

`ailine verify` は一覧に対して `{"unsupported"}` を正直に返すだけだった。
★ **正直に「無い」ことは、在ることにならない。** 一覧の値が本当に元の請求書から来たのか、
買い手が後から確かめる道が無かった。

## 何を「独立」と呼ぶか（★ ここを外すと恒真になる）

**抽出の規則（どのラベルの右を読むか）を再現しない。** `form_read` を呼ばない。
分母は**フォルダの中身**から取り、測るのは存在と辻褄:

    ① 一覧の各行が名指しする出所ファイルが**実在**する
    ② フォルダの帳票が**全部**一覧に在る（★ こちらでも読めなかった冊は咎めない・名前を出す）
    ③ 取り出した値が、その出所ファイルの**中に実際に在る**（含有）── 存在を測る、規則は測らない
    ④ 空欄には**理由**が在る（検分に 1 件ずつ）

★ ③ は「正しい値を選んだか」は測らない（それは規則の再現になる）。測るのは
  **その冊に無い値が載っていないか** ── 捏造と行の混線はこれで捕まる。
★ 見出しと項目名の対応（一覧は `請求額(税込)`・検分は `請求額`）は**出力の形式**なので
  `forms_collect` の登録簿から取る（規則ではない）。

## 測った（2026-09-12・合成検体 87 冊）

    素の一覧            破れ 0 件（含有を確かめた値 260・検分の理由 175）
    値を書き換える       含有の破れ
    理由なしで空欄にする  空欄に理由が無い
    行を 1 つ消す        一覧に無い冊
    出所を無い名前にする  出所のファイルが無い

★ 最初の検算器は 27 件鳴らした。どちらも**物差しの粗さ**だった ──
  請求番号は**数値セル**に入っているのに文字セルだけを探していた（12 件）／
  一覧と検分が同じ項目を別名で呼んでいるのを突き合わせていなかった（13 件）。
  赤が出たら、製品を触る前に 1 冊を手で解く。
"""
from __future__ import annotations

import datetime
import unicodedata

from ailine_core import filetypes, forms_collect, primitives, stack, xml_readback

#: 値を比べるための正規化 ── 空白と記号のゆれだけを畳む（★ 抽出の規則ではない）。
_FOLDED = {"・": "", "‐": "-", "−": "-", "―": "-", "ー": "ー"}


def fold(value) -> str:
    """比べるための形。★ NFKC は全角／半角のゆれだけを消す目的で使う。"""
    text = unicodedata.normalize("NFKC", str(value if value is not None else ""))
    text = "".join(text.split())
    for a, b in _FOLDED.items():
        text = text.replace(a, b)
    return text


def item_of_header(header: str) -> str:
    """一覧の見出し → 検分の「項目」名（位置で 1 対 1・登録簿から取る）。"""
    heads = forms_collect.HEADERS[1:]
    return forms_collect.FIELDS[heads.index(header)] if header in heads else header


def source_values(path) -> list:
    """出所ファイルの中身を「値の並び」で返す。読めなければ None。

    ★★ 読み手の選び方を測って決めた（2026-09-12）: 元の請求書は `openpyxl` の
      **キャッシュ値**で読む。`xml_readback` は数式セルを落とすので、金額（`=I40` のような
      数式）が 1 つも見えず **53 冊で誤って鳴った**（実測）。含有を測るには「人が見ている値」
      が要る。
    ★ したがってこの器官の独立性は**読み手の独立ではなく、規則を再現しないこと**で担保する
      ── どのセルを請求額とみなすかは一切決めない。逆に、自分の出力（一覧）の側は
      書き手（openpyxl）と別実装の `xml_readback` で読む（下の `_list_rows`）。
    ★ PDF はテキスト層をそのまま読む ── `pdf_grid` の語の塊とは別経路。
    """
    suffix = path.suffix.lower()
    try:
        if suffix == filetypes.PDF_SUFFIX:
            import pdfplumber            # ★ 必須依存だが import はここに閉じる（pdf_grid と同じ作法）
            with pdfplumber.open(str(path)) as pdf:
                return [page.extract_text() or "" for page in pdf.pages]
        import openpyxl
        book = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            return [v for sheet in book.worksheets for row in sheet.iter_rows(values_only=True)
                    for v in row if v is not None and str(v).strip()]
        finally:
            book.close()
    except Exception:                    # noqa: BLE001 ── 読めない理由は問わない（咎めないため）
        return None


def date_shapes(value) -> set:
    """日付の**西暦の書き方**をいくつか。★ 和暦は作らない ── 元号の対応表を持った時点で
    「読む規則」を再現し始めるので、確かめられないものは確かめられないと言う側に倒す。"""
    y, m, d = value.year, value.month, value.day
    return {f"{y}年{m}月{d}日", f"{y}/{m}/{d}", f"{y}-{m}-{d}", f"{y}.{m}.{d}",
            f"{y}/{m:02d}/{d:02d}", f"{y}-{m:02d}-{d:02d}", f"{y}{m:02d}{d:02d}"}


def contained(value, values: list):
    """その値が元の中に在るか。★ 数と文字の行き違いを両方見る（片方だけ見ると嘘になる）。

    戻り値は True / False / **None（確かめられなかった）**。
    ★ None は日付だけに出る ── 元が和暦などの書き方だと西暦の形では見つからない。
      そこを False（破れ）にすると嘘の警報になり、True にすると空虚な合格になる。
      **確かめられなかったものは、確かめられなかったと数える。**
    """
    if isinstance(value, (datetime.date, datetime.datetime)):
        want = value.date() if isinstance(value, datetime.datetime) else value
        for v in values:
            if isinstance(v, datetime.datetime) and v.date() == want:
                return True
            if isinstance(v, datetime.date) and not isinstance(v, datetime.datetime) \
                    and v == want:
                return True
        shapes = {fold(s) for s in date_shapes(want)}
        if any(s in fold(v) for s in shapes for v in values):
            return True
        return None
    if primitives.is_number(value) and not isinstance(value, str):
        for v in values:
            if primitives.is_number(v) and not isinstance(v, str) \
                    and abs(float(v) - float(value)) <= 1e-9:
                return True
        shapes = {fold(format(float(value), ".10g"))}
        if float(value).is_integer():
            shapes.add(fold(f"{int(value):,}"))
        return any(s and any(s in fold(v) for v in values) for s in shapes)
    want = fold(value)
    # ★ 数値セルも**文字に直して**見る ── 請求番号が int で入っている冊で 12 件誤って鳴った。
    return bool(want) and any(want in fold(v) for v in values)


def candidates(folder) -> list:
    """フォルダの帳票（★ ailine 自身の出力は入力に数えない ── 分母が汚れる）。"""
    out = []
    for path in sorted(folder.iterdir()):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in (filetypes.OPENPYXL_READABLE_SUFFIX,
                                       filetypes.PDF_SUFFIX):
            continue
        if path.suffix.lower() == filetypes.OPENPYXL_READABLE_SUFFIX:
            creator, _desc = xml_readback.read_core_properties(path)
            if creator in stack.CREATOR_MARKS:
                continue
        out.append(path)
    return out


def _list_rows(list_path) -> tuple:
    """一覧シートの (見出し, [行]) と、検分の {(ファイル, 項目): 理由}。"""
    data = xml_readback.read_grid(list_path, forms_collect.SHEET_NAME)
    grid = data.get("grid") or {}
    rows = {}
    for (r, c), v in grid.items():
        rows.setdefault(r, {})[c] = v
    head = min(rows) if rows else None
    headers = {c: str(v or "").strip() for c, v in (rows.get(head) or {}).items()}
    body = [cells for r, cells in sorted(rows.items()) if head is not None and r > head]

    insp = xml_readback.read_grid(list_path, forms_collect.INSPECT_SHEET)
    igrid, irows = insp.get("grid") or {}, {}
    for (r, c), v in igrid.items():
        irows.setdefault(r, {})[c] = v
    ihead = min(irows) if irows else None
    reasons = {}
    if ihead is not None:
        cols = {str(v or "").strip(): c for c, v in (irows.get(ihead) or {}).items()}
        f, i, w = (cols.get(n) for n in (forms_collect.INSPECT_HEADERS[0],
                                         forms_collect.INSPECT_HEADERS[1],
                                         forms_collect.INSPECT_HEADERS[3]))
        for r, cells in irows.items():
            if r == ihead or None in (f, i, w):
                continue
            reasons[(str(cells.get(f) or ""), str(cells.get(i) or ""))] = str(cells.get(w) or "")
    return headers, body, reasons


def verify_forms_list(list_path, folder) -> dict:
    """一覧をフォルダの帳票と突き合わせる。

    戻り値: {"breaks": [(名前, 名指し)], "facts": {...}, "mismatch": bool} または
            {"unsupported": 理由}。
    """
    headers, body, reasons = _list_rows(list_path)
    if not headers or forms_collect.HEADERS[0] not in headers.values():
        return {"unsupported": f"一覧シート（{forms_collect.SHEET_NAME}）が読めません: {list_path}"}
    name_col = next(c for c, n in headers.items() if n == forms_collect.HEADERS[0])

    breaks, unreadable, unchecked, checked = [], [], [], 0
    listed = {}
    for cells in body:
        name = str(cells.get(name_col) or "").strip()
        if not name:
            breaks.append(("★ 出所の名前が空の行がある", "（元ファイルの列）"))
            continue
        listed[name] = cells
        source = folder / name
        if not source.is_file():
            breaks.append(("★ 出所のファイルが無い", name))
            continue
        values = source_values(source)
        if values is None:
            unreadable.append(name)
            continue
        for col, header in headers.items():
            if col == name_col:
                continue
            value = cells.get(col)
            if value is None or str(value).strip() == "":
                item = item_of_header(header)
                if not reasons.get((name, item), "").strip():
                    breaks.append(("★ 空欄に理由が無い", f"{name}／{item}"))
                continue
            verdict = contained(value, values)
            if verdict is None:
                unchecked.append(f"{name}／{header}")
            elif verdict:
                checked += 1
            else:
                breaks.append(("★ 元に無い値が載っている（含有の破れ）",
                               f"{name}／{header}: {value!r}"))

    for path in candidates(folder):
        if path.name in listed:
            continue
        # ★ こちらでも読めない冊は咎めない（製品が名指しで断った冊を二度叱らない）。
        if source_values(path) is None:
            unreadable.append(path.name)
        else:
            breaks.append(("★ フォルダに在るのに一覧に無い冊", path.name))

    facts = {"一覧の行": len(listed), "フォルダの帳票": len(candidates(folder)),
             "含有を確かめた値": checked, "検分の理由": len(reasons)}
    if unreadable:
        facts["こちらでも読めなかった冊"] = f"{len(unreadable)} 件（{', '.join(unreadable[:5])}）"
    if unchecked:
        # ★ 出ないことを合格の証拠にしない ── 確かめられなかったものは名前を出して数える。
        facts["含有を確かめられなかった値"] = (f"{len(unchecked)} 件"
                                              f"（{', '.join(unchecked[:5])}）"
                                              " ── 元の書き方が西暦でないため")
    return {"breaks": breaks, "facts": facts, "mismatch": bool(breaks)}
