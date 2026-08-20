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
import zipfile
from dataclasses import dataclass, field
from xml.etree import ElementTree as ET

import openpyxl

from ailine_core import multifile, total_row

PROVENANCE_HEADERS = ("元ファイル", "元行")
# ★ 赤2 の直し（2026-08-21 実機敵対検分）: 署名判定はサフィックス形（元ファイル_2 等）も
#   「自分」と認める。素の名前そのもの、または末尾に _数字 が付いた形のどちらにも当たる。
_PROVENANCE_SIGNATURE_RE = tuple(re.compile(rf"^{re.escape(name)}(_\d+)?$")
                                 for name in PROVENANCE_HEADERS)

# ★ jisaku-review#1 critical の直し: 署名を「列名だけ」から「列名 AND 書き手の印」に。
# stack が出力を書く時に docProps/core.xml の dc:creator へこの印を残す
# （cmd_stack が wb.properties.creator = CREATOR_MARK を設定）。
CREATOR_MARK = "ailine stack"
_CORE_NS = {"cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
           "dc": "http://purl.org/dc/elements/1.1/"}


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
       元ファイル_N / 元行_N）と一致するか。★ 列名だけの一致であり、これ単独では
       『自分の出力』の証明にならない（jisaku-review#1 実測: たまたま同じ列名の
       人のファイルを誤認する）── 呼び出し側は `is_own_output` を使うこと。"""
    if len(headers) < 2:
        return False
    return all(pat.match(str(h)) for pat, h in zip(_PROVENANCE_SIGNATURE_RE, headers[-2:]))


def _read_creator(path) -> str | None:
    """docProps/core.xml の dc:creator を直読み（zip 直読み・軽い専用の読み）。
       読めなければ None（壊れている/該当なし = 印なし = 他人のファイル扱い ── fail closed）。"""
    try:
        with zipfile.ZipFile(path) as z:
            if "docProps/core.xml" not in z.namelist():
                return None
            root = ET.fromstring(z.read("docProps/core.xml"))
    except Exception:
        return None
    el = root.find("dc:creator", _CORE_NS)
    return el.text if el is not None else None


def is_own_output(path, headers: list) -> bool:
    """署名 = 列名の一致 **AND** 書き手の印（docProps/core.xml の creator）。
       ★★ jisaku-review#1 critical の直し（実機再現済み）: `is_own_signature`（列名だけ）は
       末尾2列がたまたま『元ファイル』『元行』という名前の人のファイルを前回出力と誤認し、
       --overwrite 無しで無警告上書きしてデータを消した。名前が合っていても印が無ければ
       他人のファイル ── fail closed（自己参照除外 V6・書き込み関所のどちらもこの関数を使う）。"""
    if not is_own_signature(headers):
        return False
    return _read_creator(path) == CREATOR_MARK


def numeric_column_names(ws, header_row: int, headers: list) -> list:
    """headers（基準ファイルの列名）のうち、データ行のどこかで数値を持つ列名の一覧。
       ★ jisaku-review#3/#6 の直し: Σ 照合・報告を『最初の数値列』1本だけでなく
       全数値列に広げるための土台（合計行検出の keyed 列＝ multifile.numeric_value_column
       の1本はここでは変えない・呼び出し側で従来どおり別に決める）。"""
    max_row = ws.max_row or header_row
    out = []
    for i, name in enumerate(headers, start=1):
        for row in range(header_row + 1, max_row + 1):
            v = ws.cell(row=row, column=i).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append(name)
                break
    return out


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
