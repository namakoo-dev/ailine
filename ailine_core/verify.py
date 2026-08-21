"""verify — `ailine verify <out.xlsx> <srcfolder>`（検算の単独再実行）の本体。
   DESIGN-20260821-multifile.md v2 §3⑧。

   ★ 信用の条件⑥「信じる対象が道具から検算に移る」: stack の出力ブックと元フォルダだけを
   引数に、検算（行数照合・数値列ごとの Σ 照合）を独立に再実行する。
   ★ 出力側・元側とも xml_readback（zipfile+ElementTree）で読む ── openpyxl は
   この module では import しない（本体の書き込み経路と同じ道具を検算に混ぜない。
   stack.py が openpyxl を使うのとは別の口 ── 信じる対象を道具から検算へ移す）。
   ★ ailine を import しない（tests/test_line_budget.py の移植可能性番人）。
   ★ 読むだけ ── この module にファイルへの書き込みは一切無い。
   ★ M2 の正直な但し書き: 条件の判定だけは extract_multi.predicate（純関数）を使う ──
   ここで述語を三度目に書き直すと、真理値表（tests/test_predicate_truth_table.py）が
   校正していない実装が増えて独立の意味が薄れる。extract_multi は openpyxl を import
   するため間接的には読み込まれるが、**この module の読みの経路は依然 xml_readback だけ**
   （openpyxl のオブジェクトはこの module に一切現れない）。
"""
from __future__ import annotations

import json

from ailine_core import extract_multi, total_row, xml_readback

TOLERANCE = total_row.TOLERANCE

# ★ M2（architect 致命3）: 検算できる印の集合。ailine_core/stack.py にも同じ集合があるが、
#   書き側（stack.py）と読み側（この module）は「別の口」で在ることが設計意図なので
#   import で結ばず、小さな集合リテラルの重複を選ぶ（module の独立性の宣言どおり）。
_CREATOR_MARKS = {"ailine stack", "ailine extract"}


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


def _expected_rows_for_source(path, base_headers: list, label_col_name, value_col_name,
                               sheet_name: str | None = None):
    """1元ファイルを独立に読み直し、『積まれるはずだった行』の行番号集合と、
       数値列ごとの値 {列名: {行番号: 値}} を返す。見出し行はこのファイル自身から探す
       （★ header_row=1 固定にしない ── multifile と同じ『名前の一致』基準）。
       見出し行が見つからなければ (set(), {})（この元ファイルは無視する）。
       ★ P2（architect 致命5・出荷済みの食い違い直し）: sheet_name は出力ブックのシート名
       （= stack が基準のシート名を付けている）。stack は基準名のシートを find_matching_sheet
       で優先するのに、ここが常に先頭シートを読むと基準名シートが2枚目以降にあるソースで
       別のシートを照合してしまう ── sheet_name で同じシートを狙う（無ければ read_grid が
       1枚目へ落ちる・従来どおり）。"""
    data = xml_readback.read_grid(path, sheet_name=sheet_name)
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
    """★ M2（architect 致命3）: 検算の入口。出力ブックの印（creator）と焼いた条件
       （description）から**種類**を先に決め、種類ごとの検算へ振り分ける。
       ailine の印が無いブックには {"unmarked": True} を返す ── 0 件照合で合格を
       名乗らない（空虚な合格の禁止・呼び出し側が exit 4 にする）。"""
    creator, description = xml_readback.read_core_properties(out_path)
    if creator == "ailine stack" and not description:
        return _verify_stack(out_path, src_folder)
    if creator in _CREATOR_MARKS and description:
        try:
            cond = json.loads(description)
        except (TypeError, ValueError):
            cond = None
        if isinstance(cond, dict) and cond.get("tool") == "ailine" and cond.get("kind") == "extract":
            return verify_extract(out_path, src_folder, cond.get("column"), cond.get("cmp"),
                                   cond.get("value"), sheet_name=cond.get("sheet"))
    return {"unmarked": True}


def _expected_rows_for_extract_source(path, base_headers: list, label_col_name, cond_col_name,
                                       match_fn, sheet_name: str | None = None):
    """M2: 1元ファイルを独立に読み直し、『抽出されるはずだった行』の行番号集合と、
       列ごとの値 {列名: {行番号: 値}} を返す（_expected_rows_for_source の抽出版）。
       違いは2点だけ:
       ① 合計行の除外に使う数値列は『最初の数値列』ではなく**条件列そのもの**
          （extract_multi.evaluate_and_extract と同じ選択 ── 書いた側と同じ列で同じ
          除外を再現しないと、検算が別の行を数えて偽 ⚠ を出す）
       ② 除外を引いた候補行を、さらに条件（match_fn）で絞る
       ★ 読めない元ファイル（.xls・壊れ）は (set(), {})＝この元ファイルは無視する
       （run 側が『取れなかった』として名指しで開示済み）。"""
    try:
        data = xml_readback.read_grid(path, sheet_name=sheet_name)
    except Exception:
        return set(), {}
    header_row = _find_header_row(data, base_headers)
    if header_row is None:
        return set(), {}
    src_headers = xml_readback.header_names(data, header_row=header_row)
    col_for_base = {bh: _column_index(src_headers, bh) for bh in base_headers}
    max_row = data["max_row"]
    all_rows = list(range(header_row + 1, max_row + 1))
    grid = data["grid"]

    label_col = col_for_base.get(label_col_name) if label_col_name else None
    cond_col = col_for_base.get(cond_col_name) if cond_col_name else None
    if label_col and cond_col:
        triples = [(r, grid.get((r, label_col)), grid.get((r, cond_col))) for r in all_rows]
        verdict = total_row.split_total_rows(triples)
    else:
        verdict = total_row.TotalRowVerdict(excluded=[], adopted_rows=[], mismatches=[])
    excluded_rows = {e.row for e in verdict.excluded}
    num_cols = len(src_headers)
    data_rows = [r for r in all_rows if xml_readback.row_has_any_value(data, r, num_cols)]
    candidate_rows = [r for r in data_rows if r not in excluded_rows]
    expected_rows = {r for r in candidate_rows
                     if match_fn(grid.get((r, cond_col)) if cond_col else None)}

    values: dict = {}
    for bh in base_headers:
        col = col_for_base.get(bh)
        if col is None:
            continue
        values[bh] = {r: grid[(r, col)] for r in expected_rows if (r, col) in grid}
    return expected_rows, values


def verify_extract(out_path, src_folder, col, cmp, value, sheet_name: str | None = None,
                    sources=None) -> dict:
    """M2 の出力（抽出集約）の検算。戻り値の形は _verify_stack と同一
       （row_count / sums / mismatch）── 報告（cli_render.render_verify_report）は
       種類を知らないままでよい。
       ★ sources: 照合する元ファイルの明示リスト（省略時は出所列に出てくるファイルだけ）。
       run の事後条件は候補ファイル全部を渡す ── 一致 0 行のファイルは出所列に現れないため、
       出所列だけを頼りにすると『1 冊まるごと落ちた』が検算をすり抜ける。
       ★ 読むだけ（openpyxl は経由しない・xml_readback のみ）。"""
    out_data = xml_readback.read_grid(out_path)
    out_headers = xml_readback.header_names(out_data, header_row=1)
    base_headers = out_headers[:-2]     # ★ 出所2列は名前でなく位置（末尾2列）で判定
    out_rows = xml_readback.data_row_numbers(out_data, header_row=1)
    grid = out_data["grid"]

    if sources is None:
        file_col, row_col = len(out_headers) - 1, len(out_headers)
        names = {str(grid[(r, file_col)]) for r in out_rows if (r, file_col) in grid}
        paths = [src_folder / n for n in sorted(names)]
    else:
        paths = list(sources)

    numeric_cols = _numeric_columns(grid, base_headers, out_rows)
    label_col_name = base_headers[0] if base_headers else None
    match_fn = extract_multi.predicate(cmp, value)

    expected_total = 0
    sums_source = {name: 0.0 for name in numeric_cols}
    for path in paths:
        if not path.exists():
            continue
        expected_rows, values = _expected_rows_for_extract_source(
            path, base_headers, label_col_name, col, match_fn, sheet_name=sheet_name)
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


def _verify_stack(out_path, src_folder) -> dict:
    """検算だけを独立に再実行する（M1書き = 縦積みの出力）。戻り値:
       {"row_count": {"source": int, "output": int},
        "sums": {列名: {"source": float, "output": float}},
        "mismatch": None または {"kind": "row_count"|"sum", "column": str|None,
                                  "source": float, "output": float}}
       ★ mismatch が立ったら、そこで比較を止める（呼び出し側が exit 5 にする）。
       行数の不一致を先に見る ── 行が消えていれば Σ もどうせ狂うが、名指しは行数から。
    """
    out_data = xml_readback.read_grid(out_path)
    base_sheet_name = out_data.get("sheet_name")   # ★ P2: 各ソースをこの名前で引き当てる
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
                                                            label_col_name, value_col_name,
                                                            sheet_name=base_sheet_name)
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
