"""stack — M1書き `ailine stack <folder> --out <path>`（縦積み・UNION ALL）の本体。
   DESIGN-20260821-multifile.md v2 §1(M1書き)・v2.1。

   ★ ailine を import しない（tests/test_line_budget.py の移植可能性番人）── 見出し行の
   推定（detect_header_row/_row_char_stats）は ailine.py 側（cmd_stack）が既存のものを1回だけ
   呼び、その結果を値としてこの module へ渡す（cmd_scan と同じ配線）。
   ★ 既存部品の再利用: 分母・基準ファイル方式・3判定は multifile.py、合計行の識別と
   閉じる検査は total_row.py（どちらも既存・単位L で完成済み）。この module は
   「積む行の決定」と「出所列つきで積む」という stack 固有の配線だけを持つ。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import openpyxl

from ailine_core import multifile, total_row

PROVENANCE_HEADERS = ("元ファイル", "元行")
# ★ 赤2 の直し（2026-08-21 実機敵対検分）: 署名判定はサフィックス形（元ファイル_2 等）も
#   「自分」と認める。素の名前そのもの、または末尾に _数字 が付いた形のどちらにも当たる。
_PROVENANCE_SIGNATURE_RE = tuple(re.compile(rf"^{re.escape(name)}(_\d+)?$")
                                 for name in PROVENANCE_HEADERS)


def _is_blank(v) -> bool:
    return total_row._is_blank_cell(v)


def _row_has_any_value(ws, row: int, num_cols: int) -> bool:
    """row の 1..num_cols 列のどこかに値があるか（★ adopted_rows とは別の判定 ──
       金額が空欄でも他の列に値があれば『データ行』とみなす。凍結検体
       test_data_row_with_empty_numeric_cell_is_still_stacked の配線）。"""
    return any(not _is_blank(ws.cell(row=row, column=c).value) for c in range(1, num_cols + 1))


def own_output_headers(headers: list) -> list:
    """利用者の列名と衝突したら機械的サフィックス（元ファイル_2 等）で逃がす。
       戻り値: 実際に使う出所列名2本（衝突が無ければ PROVENANCE_HEADERS そのまま）。"""
    out = []
    for name in PROVENANCE_HEADERS:
        candidate = name
        n = 2
        while candidate in headers or candidate in out:
            candidate = f"{name}_{n}"
            n += 1
        out.append(candidate)
    return out


def is_own_signature(headers: list) -> bool:
    """headers の末尾2列が出所列の見出し（素の名前 または 衝突時のサフィックス形
       元ファイル_N / 元行_N）と一致するか（＝『これは ailine stack が前回書いた出力』の
       署名判定）。★ 赤2: サフィックス形も『自分』と認めないと、基準ファイルが既に
       『元ファイル』列を持つ場合の前回出力（サフィックスつき）を人の出力と誤認して
       関所が exit 7 で閉まる（自己参照除外 V6 も同じ判定関数を使うので同じ根で破れる）。"""
    if len(headers) < 2:
        return False
    return all(pat.match(str(h)) for pat, h in zip(_PROVENANCE_SIGNATURE_RE, headers[-2:]))


@dataclass(frozen=True)
class FileStackResult:
    """1ファイルを積んだ（または積めなかった）結果。
       rows: [(base_headers 順の値リスト, 元行番号), ...]（積めた時のみ）。
       excluded/mismatches: total_row.split_total_rows の戻り値そのまま。
       col_a_mismatch: (col_a_count, used_range_count) 食い違い時のみ（③・可視化専用）。"""
    name: str
    status: str                      # "積んだ" / "積めなかった"
    reason: str | None = None
    reordered: bool = False
    rows: list = field(default_factory=list)
    excluded: list = field(default_factory=list)
    mismatches: list = field(default_factory=list)
    col_a_mismatch: tuple | None = None


def evaluate_and_stack(path, base_headers: list, base_sheet_name, header_row: int,
                        value_col_name: str | None) -> FileStackResult:
    """1ファイルを基準と照合し、取れていれば『積む行』を確定して値まで読む。
       ★ どんな失敗でも例外を上げず名指し+理由で返す（multifile.evaluate_file と同じ線）。"""
    if path.suffix.lower() != ".xlsx":
        return FileStackResult(name=path.name, status="積めなかった", reason="旧形式(.xls)")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return FileStackResult(name=path.name, status="積めなかった", reason=f"読み込み失敗: {e}")
    try:
        ws = multifile.find_matching_sheet(wb, base_sheet_name)
        other_headers = multifile.read_row_headers(ws, header_row)
        status, detail = multifile.classify_headers(base_headers, other_headers)
        if status == "取れなかった":
            return FileStackResult(name=path.name, status="積めなかった", reason=detail)
        reordered = bool(detail)

        # 各 base 列 → このファイル自身の列位置（並べ替えファイルで位置がずれる対策）。
        col_for_base = {bh: multifile._column_index(other_headers, bh) for bh in base_headers}
        max_row = ws.max_row or header_row
        all_rows = list(range(header_row + 1, max_row + 1))

        label_col = col_for_base.get(base_headers[0]) if base_headers else None
        value_col = col_for_base.get(value_col_name) if value_col_name else None
        if label_col and value_col:
            triples = [(r, ws.cell(row=r, column=label_col).value,
                        ws.cell(row=r, column=value_col).value) for r in all_rows]
            verdict = total_row.split_total_rows(triples)
        else:
            verdict = total_row.TotalRowVerdict(excluded=[], adopted_rows=[], mismatches=[])

        excluded_rows = {e.row for e in verdict.excluded}
        num_cols = len(other_headers)
        data_rows = [r for r in all_rows if _row_has_any_value(ws, r, num_cols)]
        stack_rows = [r for r in data_rows if r not in excluded_rows]

        rows = []
        for r in stack_rows:
            values = [ws.cell(row=r, column=col_for_base[bh]).value for bh in base_headers]
            rows.append((values, r))

        col_a_count = sum(1 for r in all_rows if not _is_blank(ws.cell(row=r, column=1).value))
        used_range_count = len(all_rows)
        col_a_mismatch = None
        if col_a_count != used_range_count:
            col_a_mismatch = (col_a_count, used_range_count)

        return FileStackResult(name=path.name, status="積んだ", reordered=reordered, rows=rows,
                                excluded=verdict.excluded, mismatches=verdict.mismatches,
                                col_a_mismatch=col_a_mismatch)
    finally:
        wb.close()


def fmt_num(v) -> str:
    """数値表示: 整数値は小数点なしで（650.0 でなく 650）。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(f)) if f.is_integer() else str(f)
