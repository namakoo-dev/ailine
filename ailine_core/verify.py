"""verify — `ailine verify <out.xlsx> <srcfolder>`（検算の単独再実行）の本体。
   DESIGN-20260821-multifile.md v2 §3⑧。

   ★ 信用の条件⑥「信じる対象が道具から検算に移る」: stack の出力ブックと元フォルダだけを
   引数に、検算（行数照合・数値列ごとの Σ 照合）を独立に再実行する。
   ★ 出力側・元側とも xml_readback（zipfile+ElementTree）で読む ── openpyxl は
   この module では import しない（本体の書き込み経路と同じ道具を検算に混ぜない。
   stack.py が openpyxl を使うのとは別の口 ── 信じる対象を道具から検算へ移す）。
   ★ ailine を import しない（tests/test_line_budget.py の移植可能性番人）。
   ★ 読むだけ ── この module にファイルへの書き込みは一切無い。
"""
from __future__ import annotations

from ailine_core import total_row, xml_readback

TOLERANCE = total_row.TOLERANCE


def fmt_num(v) -> str:
    """数値表示: 整数値は小数点なしで（600.0 でなく 600）。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(f)) if f.is_integer() else str(f)


def _column_index(headers: list, name: str):
    try:
        return headers.index(name) + 1
    except ValueError:
        return None


def _find_header_row(data: dict, base_headers: list, max_scan: int = 30):
    """base_headers と同じ多重集合を持つ最初の行を見出し行とみなす（並べ替えも許す ──
       multifile.classify_headers の『取れた』と同じ線）。無ければ None。"""
    for r in range(1, max_scan + 1):
        names = xml_readback.header_names(data, header_row=r)
        if names and sorted(names) == sorted(base_headers):
            return r
    return None


def _numeric_columns(grid: dict, base_headers: list, rows: list) -> list:
    """base_headers のうち、出力側で数値を持ったことがある列名（=『数値列』）。
       ★ 出力の列順は base_headers そのままなので、位置(1起点)で引ける。"""
    out = []
    for i, name in enumerate(base_headers, start=1):
        if any(isinstance(grid.get((r, i)), (int, float)) and not isinstance(grid.get((r, i)), bool)
               for r in rows):
            out.append(name)
    return out


def _expected_rows_for_source(path, base_headers: list, label_col_name, value_col_name):
    """1元ファイルを独立に読み直し、『積まれるはずだった行』の行番号集合と、
       数値列ごとの値 {列名: {行番号: 値}} を返す。見出し行はこのファイル自身から探す
       （★ header_row=1 固定にしない ── multifile と同じ『名前の一致』基準）。
       見出し行が見つからなければ (set(), {})（この元ファイルは無視する）。"""
    data = xml_readback.read_grid(path)
    header_row = _find_header_row(data, base_headers)
    if header_row is None:
        return set(), {}
    src_headers = xml_readback.header_names(data, header_row=header_row)
    col_for_base = {bh: _column_index(src_headers, bh) for bh in base_headers}
    max_row = data["max_row"]
    all_rows = list(range(header_row + 1, max_row + 1))
    grid = data["grid"]

    label_col = col_for_base.get(label_col_name) if label_col_name else None
    value_col = col_for_base.get(value_col_name) if value_col_name else None
    if label_col and value_col:
        triples = [(r, grid.get((r, label_col)), grid.get((r, value_col))) for r in all_rows]
        verdict = total_row.split_total_rows(triples)
    else:
        verdict = total_row.TotalRowVerdict(excluded=[], adopted_rows=[], mismatches=[])
    excluded_rows = {e.row for e in verdict.excluded}
    num_cols = len(src_headers)
    data_rows = [r for r in all_rows if xml_readback.row_has_any_value(data, r, num_cols)]
    expected_rows = {r for r in data_rows if r not in excluded_rows}

    values: dict = {}
    for bh in base_headers:
        col = col_for_base.get(bh)
        if col is None:
            continue
        values[bh] = {r: grid[(r, col)] for r in expected_rows if (r, col) in grid}
    return expected_rows, values


def verify_output(out_path, src_folder) -> dict:
    """検算だけを独立に再実行する。戻り値:
       {"row_count": {"source": int, "output": int},
        "sums": {列名: {"source": float, "output": float}},
        "mismatch": None または {"kind": "row_count"|"sum", "column": str|None,
                                  "source": float, "output": float}}
       ★ mismatch が立ったら、そこで比較を止める（呼び出し側が exit 5 にする）。
       行数の不一致を先に見る ── 行が消えていれば Σ もどうせ狂うが、名指しは行数から。
    """
    out_data = xml_readback.read_grid(out_path)
    out_headers = xml_readback.header_names(out_data, header_row=1)
    base_headers = out_headers[:-2]     # ★ 出所2列は名前でなく位置（末尾2列）で判定
    file_col, row_col = len(out_headers) - 1, len(out_headers)

    out_rows = xml_readback.data_row_numbers(out_data, header_row=1)
    grid = out_data["grid"]

    refs: dict = {}
    for r in out_rows:
        fname, src_row = grid.get((r, file_col)), grid.get((r, row_col))
        if fname is None or src_row is None:
            continue
        refs.setdefault(str(fname), []).append(int(src_row))

    numeric_cols = _numeric_columns(grid, base_headers, out_rows)
    label_col_name = base_headers[0] if base_headers else None
    value_col_name = numeric_cols[0] if numeric_cols else None

    expected_total = 0
    sums_source = {name: 0.0 for name in numeric_cols}
    for fname in sorted(refs):
        path = src_folder / fname
        if not path.exists():
            continue
        expected_rows, values = _expected_rows_for_source(path, base_headers,
                                                            label_col_name, value_col_name)
        expected_total += len(expected_rows)
        for name in numeric_cols:
            for v in values.get(name, {}).values():
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    sums_source[name] += v

    actual_total = len(out_rows)
    row_count = {"source": expected_total, "output": actual_total}
    if expected_total != actual_total:
        return {"row_count": row_count, "sums": {}, "mismatch": {
            "kind": "row_count", "column": None, "source": expected_total, "output": actual_total}}

    sums_output = {name: 0.0 for name in numeric_cols}
    for i, name in enumerate(base_headers, start=1):
        if name not in numeric_cols:
            continue
        for r in out_rows:
            v = grid.get((r, i))
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                sums_output[name] += v

    sums = {name: {"source": sums_source[name], "output": sums_output[name]} for name in numeric_cols}
    for name in numeric_cols:
        if abs(sums_source[name] - sums_output[name]) > TOLERANCE:
            return {"row_count": row_count, "sums": sums, "mismatch": {
                "kind": "sum", "column": name,
                "source": sums_source[name], "output": sums_output[name]}}

    return {"row_count": row_count, "sums": sums, "mismatch": None}
