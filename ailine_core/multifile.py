"""multifile — M1読み: `ailine scan <folder>` の棚卸しロジック（書き込みゼロ）。
   DESIGN-20260821-multifile.md v2 §1(M1読み)・§2(骨)・§3(③④分母/原本無変更)。

   ★ LibreOffice は一切起動しない（openpyxl のみで完結・読むだけ）。
   ★ ailine を import しない（tests/test_line_budget.py の移植可能性番人）── 見出し行の
   推定（detect_header_row/_row_char_stats）は ailine.py 側（cmd_scan）が既存のものを1回だけ
   呼び、その結果（header_row・base_headers）をこのモジュールの関数へ値として渡す。
"""
from __future__ import annotations

from pathlib import Path

import openpyxl

_TEMP_PREFIX = "~$"                      # Excel の一時ファイル（開いている間だけ現れる隣接ファイル）
_EXCEL_EXTS = {".xlsx", ".xls"}
_MAX_HEADER_COLS = 200                   # 見出し行を読む安全上限（ailine.py の MAX_COLS とは独立）


def classify_folder_contents(folder: Path):
    """folder 直下（サブフォルダの中は見ない）を分類する。
       戻り値: (candidates: 名前順の Path リスト, excluded: {"temp": n, "subdirs": n})。
       ★ 分母そのものが検証対象（V7）── ~$ 一時ファイルとサブフォルダは対象外として数える
       （1件以上あれば呼び出し側が1行ずつ開示する）。その他の拡張子は黙って無視してよい。"""
    candidates = []
    excluded = {"temp": 0, "subdirs": 0}
    for item in sorted(folder.iterdir(), key=lambda p: p.name):
        if item.is_dir():
            excluded["subdirs"] += 1
            continue
        if not item.is_file():
            continue
        if item.name.startswith(_TEMP_PREFIX):
            excluded["temp"] += 1
            continue
        if item.suffix.lower() in _EXCEL_EXTS:
            candidates.append(item)
    return candidates, excluded


def open_base_workbook(candidates):
    """基準ファイル方式: パス辞書順（呼び出し側で名前順に並べ済み）で最初に読めた .xlsx を
       基準にする。戻り値: (path, workbook) または、読める .xlsx が1つも無ければ (None, None)。
       ★ .xls は openpyxl で開けないため基準になれない（読めたものだけが資格を持つ）。"""
    for path in candidates:
        if path.suffix.lower() != ".xlsx":
            continue
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception:
            continue
        return path, wb
    return None, None


def read_row_headers(ws, header_row: int) -> list:
    """header_row（1起点）の先頭列から連続する非空セルを見出し名として読む
       （第一波は単純に ── 列の間に空白を挟む見出しは扱わない・DESIGN §2骨）。"""
    headers = []
    for c in range(1, _MAX_HEADER_COLS + 1):
        v = ws.cell(row=header_row, column=c).value
        if v in (None, ""):
            break
        headers.append(str(v))
    return headers


def find_matching_sheet(wb, base_sheet_name: str | None):
    """他ファイルは基準と同名のシートを探し、無ければ最初のシートで照合する（DESIGN §2骨）。"""
    if base_sheet_name and base_sheet_name in wb.sheetnames:
        return wb[base_sheet_name]
    return wb[wb.sheetnames[0]]


def classify_headers(base_headers: list, other_headers: list):
    """3判定（列名の完全一致のみが根拠・ゆるい寄せはしない）:
       並びまで一致 → ("取れた", None)
       多重集合が一致・順序だけ違う → ("取れた", "並べ替え")
       それ以外 → ("取れなかった", "欠け/余りの名指し")"""
    if other_headers == base_headers:
        return "取れた", None
    if sorted(other_headers) == sorted(base_headers):
        return "取れた", "並べ替え"
    missing = [h for h in base_headers if h not in other_headers]
    extra = [h for h in other_headers if h not in base_headers]
    parts = []
    if missing:
        parts.append(f"欠け: {', '.join(missing)}")
    if extra:
        parts.append(f"余り: {', '.join(extra)}")
    return "取れなかった", "; ".join(parts) if parts else "列名が一致しません"


def evaluate_file(path: Path, base_headers: list, base_sheet_name: str | None, header_row: int) -> dict:
    """1ファイルを基準と照合する。戻り値: {"name", "status", "reason"(取れなかった時),
       "reordered"(並べ替えで取れた時)}。★ どんな失敗でも例外を上げず名指し+理由にして返す
       （$0 条件「黙って失敗する」の裏返し ── 報告が成果物）。"""
    if path.suffix.lower() != ".xlsx":
        return {"name": path.name, "status": "取れなかった", "reason": "旧形式(.xls)"}
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return {"name": path.name, "status": "取れなかった", "reason": f"読み込み失敗: {e}"}
    try:
        ws = find_matching_sheet(wb, base_sheet_name)
        other_headers = read_row_headers(ws, header_row)
    finally:
        wb.close()
    status, detail = classify_headers(base_headers, other_headers)
    entry = {"name": path.name, "status": status}
    if status == "取れなかった":
        entry["reason"] = detail
    elif detail:   # "並べ替え"
        entry["reordered"] = True
    return entry
