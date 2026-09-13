"""accounts_apply — 候補の冊の「採用」を、元の仕訳に写す（2026-09-13・買い手役 3 回目・会計の 1 位）。

★★ なぜ在るか: 候補の冊は『候補の科目』を右端の列に出し、`借方勘定科目` は空のまま（決めるのは人）。
  会計事務所は毎月 200〜500 行の「U 列 → C 列の転記」と「末尾 6 列の削除」を手でやっていた ──
  「採用の一往復を閉じれば 20,000 円」。人が『採用』列に○を付けた行だけ、候補の科目を
  **元の仕訳ファイルの借方勘定科目に写した**取込用ファイルを書く。列順・文字コード・見出し行・説明行は
  元のまま。○の無い行は空のまま（決めない）。

★ 決めるのは人のまま ── ここは「人が決めたことを、元の形に戻す」だけ。候補を勝手に採らない。
★ 事後条件: 書いた物を読み戻し、**変わったセルが採用の行の借方勘定科目だけ**であることを数える。
"""
from __future__ import annotations

import csv
import io
from pathlib import Path

import openpyxl

from ailine_core import accounts_core, filetypes

#: 採用の列の見出し（人が候補シートの右端に足す）。
ADOPT_HEADER = "採用"
#: ○と読む印（全角・半角・数字の 1・英字）。空欄は「採らない」。
ADOPT_MARKS = frozenset({"○", "〇", "◯", "o", "O", "x", "X", "1", "✓", "はい", "採用", "yes", "y"})


def read_adoptions(candidate_book) -> tuple:
    """候補の冊から {元行: 候補の科目}（○の行だけ）を読む。戻り値: (採用, 断りの文 or None, 見た行数)。"""
    wb = openpyxl.load_workbook(candidate_book, read_only=True, data_only=True)
    try:
        ws = wb[accounts_core.SHEET_NAME] if accounts_core.SHEET_NAME in wb.sheetnames else wb.worksheets[0]
        rows = ws.iter_rows(values_only=True)
        head = [str(v or "").strip() for v in next(rows, ())]
        need = {"元行": None, accounts_core.OUTPUT_HEADERS[0]: None, ADOPT_HEADER: None}
        for i, h in enumerate(head):
            if h in need and need[h] is None:
                need[h] = i
        if need[ADOPT_HEADER] is None:
            return {}, (f"候補の冊に『{ADOPT_HEADER}』の列がありません ── 『{accounts_core.SHEET_NAME}』シートの"
                        f"右端に『{ADOPT_HEADER}』という見出しの列を足し、採用する行に ○ を入れてください"
                        "（○ の無い行は借方勘定科目を空のまま残します）"), 0
        if need["元行"] is None or need[accounts_core.OUTPUT_HEADERS[0]] is None:
            return {}, "候補の冊の見出しに『元行』か『候補の科目』がありません（ailine accounts の出力ですか）", 0
        adopted, seen = {}, 0
        for r in rows:
            seen += 1
            mark = str(r[need[ADOPT_HEADER]] or "").strip() if need[ADOPT_HEADER] < len(r) else ""
            if mark not in ADOPT_MARKS:
                continue
            src_row = r[need["元行"]] if need["元行"] < len(r) else None
            account = r[need[accounts_core.OUTPUT_HEADERS[0]]] if need[accounts_core.OUTPUT_HEADERS[0]] < len(r) else None
            if src_row is None or not str(account or "").strip():
                continue
            adopted[int(src_row)] = str(account).strip()
        return adopted, None, seen
    finally:
        wb.close()


def _newline_of(raw: bytes) -> str:
    return "\r\n" if b"\r\n" in raw else "\n"


def apply_to_csv(journal_path, adopted: dict, encoding: str, account_col: int, out_path) -> dict:
    """CSV の該当セルだけを書き換えて `out_path` に書く。戻り値: {"changed": n, "rows": n}。

    ★ 説明行・見出し行・列順・文字コード・改行は元のまま。引用符は csv の最小引用（元と違いうる ──
      値は 1 文字も変えない）。
    """
    raw = Path(journal_path).read_bytes()
    text = raw.decode(encoding)
    bom = text.startswith("\ufeff")
    if bom:
        text = text[1:]
    nl = _newline_of(raw)
    rows = list(csv.reader(io.StringIO(text)))
    changed = 0
    for num, account in adopted.items():
        i = num - 1
        if 0 <= i < len(rows):
            while len(rows[i]) < account_col:
                rows[i].append("")
            if rows[i][account_col - 1] != account:
                rows[i][account_col - 1] = account
                changed += 1
    buf = io.StringIO()
    csv.writer(buf, lineterminator=nl).writerows(rows)
    out = buf.getvalue()
    Path(out_path).write_bytes(("\ufeff" + out if bom else out).encode(encoding))
    return {"changed": changed, "rows": len(rows)}


def apply_to_book(journal_path, adopted: dict, account_col: int, out_path) -> dict:
    wb = openpyxl.load_workbook(journal_path)
    ws = wb.worksheets[0]
    changed = 0
    for num, account in adopted.items():
        cell = ws.cell(row=num, column=account_col)
        if cell.value != account:
            cell.value = account
            changed += 1
    wb.save(out_path)
    wb.close()
    return {"changed": changed, "rows": ws.max_row}


def diff_cells(a_path, b_path, encoding: str | None) -> list:
    """2 つの仕訳ファイルで値が違うセル [(行, 列, 前, 後)] ── 事後条件の材料（別実装で読み戻す）。"""
    def cells(p):
        p = Path(p)
        if p.suffix.lower() == filetypes.CSV_SUFFIX:
            text = p.read_bytes().decode(encoding or "utf-8-sig").lstrip("\ufeff")
            return [list(r) for r in csv.reader(io.StringIO(text))]
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        try:
            return [[("" if v is None else str(v)) for v in r] for r in wb.worksheets[0].iter_rows(values_only=True)]
        finally:
            wb.close()
    a, b = cells(a_path), cells(b_path)
    out = []
    for i in range(max(len(a), len(b))):
        ra, rb = (a[i] if i < len(a) else []), (b[i] if i < len(b) else [])
        for j in range(max(len(ra), len(rb))):
            va, vb = (ra[j] if j < len(ra) else ""), (rb[j] if j < len(rb) else "")
            if str(va) != str(vb):
                out.append((i + 1, j + 1, va, vb))
    return out
