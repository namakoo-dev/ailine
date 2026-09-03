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
   ★ M3 も同じ線: match の正規化規則（normalize_key）とキー再集計（independent_key_stats）
   は ailine_core/match.py の純関数を import して使う（match.py は openpyxl を import
   するが、この module がその関数を呼んでも openpyxl オブジェクトはここに一切現れない ──
   ★ 片配線の自己点検: cmd_run_match の書き込み時事後条件（ailine.py）と _verify_match
   （この module）は independent_key_stats を共有する。別実装を2つ作らない）。
"""
from __future__ import annotations

from pathlib import Path

from ailine_core.filetypes import OPENPYXL_READABLE_SUFFIX

import json

from ailine_core import extract_multi, match, total_row, xml_readback
from ailine_core.primitives import column_index as _column_index
from ailine_core.primitives import fmt_num

TOLERANCE = total_row.TOLERANCE

# ★ M2（architect 致命3）: 検算できる印の集合。ailine_core/stack.py にも同じ集合があるが、
#   書き側（stack.py）と読み側（この module）は「別の口」で在ることが設計意図なので
#   import で結ばず、小さな集合リテラルの重複を選ぶ（module の独立性の宣言どおり）。
#   ★ M3 P 先行 commit: stack.CREATOR_MARKS に "ailine match" を足したのに合わせてここも
#   足す（片配線を作らない）。match 用の検算（_verify_match）はまだ無く、この集合に
#   入れても description ガード（下の verify_output）が無ければ {"unmarked": True} のまま
#   ── 配線は次波（M3 本体）。
#   ★ CSV 検疫接続（2026-08-22）: 同じ理由で "ailine csv" も足す（stack.CREATOR_MARKS と
#   同期・tests/test_stack_e2e.py の番人）。独立検算（csv kind 専用の verify）はまだ無い
#   ── 下の verify_output は creator=="ailine csv" を {"unsupported": ...} で正直に返すだけ
#   （{"unmarked": True} に混ぜて「他人のファイル」と誤判定しない、が今回配線する範囲）。
_CREATOR_MARKS = {"ailine stack", "ailine extract", "ailine match", "ailine csv"}


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


# ★ 検算が「書き手と同じ規則で」落とした行の控え（2026-08-26・致命①⑨）。
#   戻り値のタプルを増やすと試験の直呼びが壊れるので、パス → 行番号 の形で残す。
#   ここに入った行は「検算が裏取りしていない除外」── 呼び出し側が必ず人へ見せる。
_LAST_DROPPED: dict = {}
# ★ 読めなかった冊（致命⑥）── 黙って飛ばさず名指しする
_UNREADABLE: dict = {}


def _expected_rows_for_source(path, base_headers: list, label_col_name, numeric_col_names: list,
                               sheet_name: str | None = None):
    """1元ファイルを独立に読み直し、『積まれるはずだった行』の行番号集合と、
       数値列ごとの値 {列名: {行番号: 値}} を返す。見出し行はこのファイル自身から探す
       （★ header_row=1 固定にしない ── multifile と同じ『名前の一致』基準）。
       見出し行が見つからなければ (set(), {})（この元ファイルは無視する）。
       ★ P2（architect 致命5・出荷済みの食い違い直し）: sheet_name は出力ブックのシート名
       （= stack が基準のシート名を付けている）。stack は基準名のシートを find_matching_sheet
       で優先するのに、ここが常に先頭シートを読むと基準名シートが2枚目以降にあるソースで
       別のシートを照合してしまう ── sheet_name で同じシートを狙う（無ければ read_grid が
       1枚目へ落ちる・従来どおり）。
       ★ operator 盲検7度目の直し（2026-08-21）: 合計行の除外は『指定の1本の数値列』
       （旧 value_col_name）でなく『基準の数値列集合すべて』（numeric_col_names）を見る ──
       stack.evaluate_and_stack（書いた側）と同じ検出でないと、片方だけが取り逃がして
       恒真ペア（間違った数字同士が一致）になる。"""
    # ★★ 2026-08-26（複数ファイルの盲検・致命⑥）: ここは try を持たず、フォルダに
    #   壊れた .xlsx が 1 本混ざるだけで **生の traceback** で落ちていた（exit 1）。
    #   同じ verify.py の中でも `_attribution_mismatch` は try/except を持っており、
    #   **同一ファイル内で非対称**だった。scan と run <フォルダ> は名指しして完走する
    #   ── 4 経路のうち 2 経路だけ塞がっている片配線。
    #   ★ 黙って飛ばさない: 読めなかった冊は呼び出し側が名指しできるよう記録する。
    try:
        data = xml_readback.read_grid(path, sheet_name=sheet_name)
    except Exception as e:
        _UNREADABLE[str(path)] = f"{type(e).__name__}: {e}"
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
    value_cols = {name: col_for_base[name] for name in numeric_col_names
                  if col_for_base.get(name)}
    if label_col and value_cols:
        rows_in = []
        for r in all_rows:
            label_val = grid.get((r, label_col))
            vals = {name: grid.get((r, idx)) for name, idx in value_cols.items()}
            rows_in.append((r, label_val, vals))
        verdict = total_row.split_total_rows_multi(rows_in)
    else:
        verdict = total_row.TotalRowVerdict(excluded=[], adopted_rows=[], mismatches=[])
    excluded_rows = {e.row for e in verdict.excluded}
    num_cols = len(src_headers)
    data_rows = [r for r in all_rows if xml_readback.row_has_any_value(data, r, num_cols)]
    expected_rows = {r for r in data_rows if r not in excluded_rows}
    # ★★ 2026-08-26（複数ファイルの盲検・致命①⑨）: ここは書き手（stack）と**同じ関数**で
    #   除外を決めている。だから書き手が本物のデータ行を「合計行」と誤判定して落とすと、
    #   検算も同じ行を落とし、**両方が同じ間違いをして一致する＝恒真**になる。
    #   実測: 区切りの空行がある表で、3 列すべて埋まった売上 1,000 円が消え、
    #   stack も verify も exit 0（--json の mismatches も空）。
    #   ★ 上の docstring が書くとおり、共有は**別の恒真**（片方だけ取り逃がす）を避ける
    #     ための意図的な判断だった。だが同じ罠の裏返しでしかない。
    #   ★ 三項で解く: 書き手は除外する／検算は除外を**判断しない**／差は人に見せる。
    #     判定（expected_rows）はここでは変えない ── 変えると正しい合計行を持つ表が
    #     全部不一致になる。代わりに**落とした行を記録して、呼び出し側が開示する**。
    _LAST_DROPPED[str(path)] = sorted(r for r in data_rows if r in excluded_rows)

    values: dict = {}
    for bh in base_headers:
        col = col_for_base.get(bh)
        if col is None:
            continue
        values[bh] = {r: grid[(r, col)] for r in expected_rows if (r, col) in grid}
    return expected_rows, values


def _values_agree(a, b) -> bool:
    """帰属検算の1セル比較。両方が数値なら許容誤差 TOLERANCE、それ以外は完全一致
       （_extract_predicate/extract_multi.predicate の型の保存の哲学と同じ線）。"""
    a_num = isinstance(a, (int, float)) and not isinstance(a, bool)
    b_num = isinstance(b, (int, float)) and not isinstance(b, bool)
    if a_num and b_num:
        return abs(float(a) - float(b)) <= TOLERANCE
    return a == b


def _attribution_mismatch(out_data: dict, base_headers: list, out_rows: list, src_folder,
                           sheet_name: str | None = None):
    """★ review3#3 major の直し: 集計（行数・Σ）だけでは、Σ 不変の値入れ替え（帰属の嘘 ──
       どの行がどのファイルの何行目かの主張と実物が食い違う）を見逃す（実機再現済み）。
       出力の各行を、出所列（末尾2列＝元ファイル/元行）が指す原本の行と列名解決で
       突き合わせる。最初に見つかった不一致を
       {"file","src_row","column","output_value","source_value"} で返す（無ければ None）。
       ★ ファイルごとに1回だけ読み直してキャッシュする（xml_readback のみ・読むだけ）。
       ★ 読めない/見出しが見つからない元ファイルはここでは無視する（行数/Σ 側の検査が
       既に名指し済みのはず ── ここは『読めた行の値』だけを見る）。"""
    grid = out_data["grid"]
    out_headers = xml_readback.header_names(out_data, header_row=1)
    file_col, row_col = len(out_headers) - 1, len(out_headers)
    cache: dict = {}   # fname -> (col_for_base, src_grid) または None（無視）

    for r in out_rows:
        fname_v, src_row_v = grid.get((r, file_col)), grid.get((r, row_col))
        if fname_v is None or src_row_v is None:
            continue
        fname, src_row = str(fname_v), int(src_row_v)
        if fname not in cache:
            path = src_folder / fname
            entry = None
            if path.exists():
                try:
                    data = xml_readback.read_grid(path, sheet_name=sheet_name)
                except Exception:
                    data = None
                if data is not None:
                    header_row = _find_header_row(data, base_headers)
                    if header_row is not None:
                        src_headers = xml_readback.header_names(data, header_row=header_row)
                        col_for_base = {bh: _column_index(src_headers, bh) for bh in base_headers}
                        entry = (col_for_base, data["grid"])
            cache[fname] = entry
        entry = cache[fname]
        if entry is None:
            continue
        col_for_base, src_grid = entry
        for i, bh in enumerate(base_headers, start=1):
            src_col = col_for_base.get(bh)
            source_value = src_grid.get((src_row, src_col)) if src_col else None
            output_value = grid.get((r, i))
            if not _values_agree(source_value, output_value):
                return {"file": fname, "src_row": src_row, "column": bh,
                        "output_value": output_value, "source_value": source_value}
    return None


def verify_output(out_path, src_folder) -> dict:
    """★ M2（architect 致命3）: 検算の入口。出力ブックの印（creator）と焼いた条件
       （description）から**種類**を先に決め、種類ごとの検算へ振り分ける。
       ailine の印が無いブックには {"unmarked": True} を返す ── 0 件照合で合格を
       名乗らない（空虚な合格の禁止・呼び出し側が exit 4 にする）。
       ★ M3（design v2「verify」節）: この経路（元 1 フォルダ形）は match 出力を検算しない
       ── 照合は原本 2 冊のうちどちらが A/B かをフォルダから機械的に選べない（本実装
       verify_match_output は `ailine verify <出力> <元A> <元B>` の3引数形専用）。
       ここで黙って {"unmarked": True} に混ぜる（空虚な合格の匂わせ）よりは、
       {"unsupported": ...} を正直に返して呼び出し側が exit 4 にする。"""
    creator, description = xml_readback.read_core_properties(out_path)
    if creator == "ailine stack" and not description:
        return _verify_stack(out_path, src_folder)
    if creator == "ailine match":
        return {"unsupported": "照合出力の検算には原本2冊の指定が必要です"
                                "（`ailine verify <出力> <元A> <元B>` の形で実行してください）。"}
    if creator == "ailine csv":
        # ★ CSV 検疫接続（2026-08-22・次便）: 独立の検算はまだ無い。{"unmarked": True} に
        #   混ぜると「ailine の印が無い人のファイル」と誤って言うことになる（own 印は
        #   ある）ので、{"unsupported": ...} で正直に区別する。
        return {"unsupported": "CSV 検疫の出力は `ailine csv <元のcsv>` を再実行して"
                                "確認してください（独立の検算はまだ実装していません）。"}
    if creator in _CREATOR_MARKS and description:
        try:
            cond = json.loads(description)
        except (TypeError, ValueError):
            cond = None
        if isinstance(cond, dict) and cond.get("tool") == "ailine" and cond.get("kind") == "extract":
            return verify_extract(out_path, src_folder, cond.get("column"), cond.get("cmp"),
                                   cond.get("value"), sheet_name=cond.get("sheet"))
    return {"unmarked": True}


def _expected_rows_for_extract_source(path, base_headers: list, label_col_name,
                                       numeric_col_names: list, cond_col_name,
                                       match_fn, sheet_name: str | None = None):
    """M2: 1元ファイルを独立に読み直し、『抽出されるはずだった行』の行番号集合と、
       列ごとの値 {列名: {行番号: 値}} を返す（_expected_rows_for_source の抽出版）。
       違いは2点だけ:
       ① 除外を引いた候補行を、さらに条件（match_fn・条件列 cond_col_name）で絞る
       ② それ以外は _expected_rows_for_source と同じ ── 合計行の除外は基準の数値列
          すべて（numeric_col_names）を見る。★ operator 盲検7度目の直し（2026-08-21）:
          旧仕様は『条件列そのもの』だけを見ていた（extract_multi.evaluate_and_extract と
          揃えるためだったが、これは stack と非対称な単一列版の穴を extract 側にも
          持ち込んでいただけ ── 書いた側（evaluate_and_extract）も numeric_col_names を
          見るよう直したので、検算側もここで揃える。片方だけ直すと片配線になる）。
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
    cond_col = col_for_base.get(cond_col_name) if cond_col_name else None   # ★ 述語の対象列
    value_cols = {name: col_for_base[name] for name in numeric_col_names
                  if col_for_base.get(name)}
    if label_col and value_cols:
        rows_in = []
        for r in all_rows:
            label_val = grid.get((r, label_col))
            vals = {name: grid.get((r, idx)) for name, idx in value_cols.items()}
            rows_in.append((r, label_val, vals))
        verdict = total_row.split_total_rows_multi(rows_in)
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


def _total_word_mismatches(out_data: dict, out_headers: list, out_rows: list) -> list:
    """★ 第二の独立検出器（operator 盲検7度目 修正2）: 検出器1（split_total_rows/
       split_total_rows_multi）と意図的に盲点を共有しない ── 列解決を一切使わず、出力の
       各行の全セル値を走査して合計語を持つ行を名指しする（total_row.total_word_trip_findings
       と同じ規則）。出所列（末尾2列）があればファイル名/元行も添える。除外はしない
       （検出のみ・mismatches に kind="total_word" として積む）。
       ★ 再演検分の直し（2026-08-21 19:1x）: 走査対象は**データ列のみ**（out_headers の
       末尾2列＝出所列を除く）── 出所列の値（うちが付けたファイル名・元行番号）は対象外。
       『月次_合計表.xlsx』のように合計語を含むファイル名は実務に普通にあり、出所列まで
       走査すると正当な出力が誤発火して verify が exit 5 になっていた（stack/extract の
       own_output_headers は常に末尾2列＝出所列という契約に乗る・サフィックス形でも
       列の“位置”は変わらないので、末尾2列を機械的に除くだけで十分）。"""
    grid = out_data["grid"]
    has_provenance = len(out_headers) >= 2
    file_col = len(out_headers) - 1 if has_provenance else None
    row_col = len(out_headers) if has_provenance else None
    data_col_count = len(out_headers) - 2 if has_provenance else len(out_headers)
    trip_rows = [(r, r, [grid.get((r, c)) for c in range(1, data_col_count + 1)])
                 for r in out_rows]
    out = []
    for _ident, row_num, word in total_row.total_word_trip_findings(trip_rows):
        fname = grid.get((row_num, file_col)) if file_col else None
        src_row = grid.get((row_num, row_col)) if row_col else None
        out.append({"kind": "total_word", "file": str(fname) if fname is not None else None,
                    "row": src_row if src_row is not None else row_num, "word": word})
    return out


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
        file_col = len(out_headers) - 1   # ★ row_col は使っていないので持たない
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
            path, base_headers, label_col_name, numeric_cols, col, match_fn, sheet_name=sheet_name)
        expected_total += len(expected_rows)
        for name in numeric_cols:
            for v in values.get(name, {}).values():
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    sums_source[name] += v

    actual_total = len(out_rows)
    row_count = {"source": expected_total, "output": actual_total}
    if expected_total != actual_total:
        rc_mismatch = {"kind": "row_count", "column": None,
                       "source": expected_total, "output": actual_total}
        return {"row_count": row_count, "sums": {}, "mismatch": rc_mismatch,
                "mismatches": [rc_mismatch]}

    sums_output = {name: 0.0 for name in numeric_cols}
    for i, name in enumerate(base_headers, start=1):
        if name not in numeric_cols:
            continue
        for r in out_rows:
            v = grid.get((r, i))
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                sums_output[name] += v

    sums = {name: {"source": sums_source[name], "output": sums_output[name]} for name in numeric_cols}
    # ★ デモ撮影のリハで発覚（2026-08-21 13:2x）: 憲法1（誘導）+ 一括検出 ── 最初の不一致
    #   （Σ）で打ち切ると、帰属検算が名指しできるはずの「どの行がいくつ→いくつ」を
    #   一度も計算しないまま終わっていた。Σ と帰属は独立に検査できる（帰属は出力の各行を
    #   自分が申告する元セルと突き合わせるだけで、Σ の合否に依存しない）── 両方走らせて
    #   全所見を集める。★ 書き込み時事後条件（run/cmd_run_folder）の挙動は変えない ──
    #   `mismatch`（単数）は従来どおり「最初に見つかった1件」で、優先順位（Σ列の並び順→
    #   帰属）も従来と同じにする。追加する `mismatches`（複数）だけが新しい。
    mismatches = []
    for name in numeric_cols:
        if abs(sums_source[name] - sums_output[name]) > TOLERANCE:
            mismatches.append({"kind": "sum", "column": name,
                               "source": sums_source[name], "output": sums_output[name]})

    attribution = _attribution_mismatch(out_data, base_headers, out_rows, src_folder,
                                         sheet_name=sheet_name)
    if attribution:
        mismatches.append({"kind": "attribution", "column": attribution["column"],
                           "source": attribution["source_value"], "output": attribution["output_value"],
                           "file": attribution["file"], "src_row": attribution["src_row"]})

    # ★ operator 盲検7度目 修正2: 語のトリップワイヤ（第二の独立検出器）。
    #   `mismatch`（単数）の末尾に積む ── row_count/sum/attribution が1件も無い時だけ
    #   `mismatch` に選ばれる（呼び出し側 cmd_run_folder の書き込み時事後条件は kind で
    #   除外し、これだけでは書き込みを止めない ── 除外はしない設計と一致させる）。
    mismatches.extend(_total_word_mismatches(out_data, out_headers, out_rows))

    return {"row_count": row_count, "sums": sums,
            "mismatch": mismatches[0] if mismatches else None, "mismatches": mismatches}


def _verify_stack(out_path, src_folder) -> dict:
    """検算だけを独立に再実行する（M1書き = 縦積みの出力）。戻り値:
       {"row_count": {"source": int, "output": int},
        "sums": {列名: {"source": float, "output": float}},
        "mismatch": None または見つかった最初の1件（呼び出し側=書き込み時事後条件が
                     使う・従来と同じ優先順位: row_count → Σ(列順) → 帰属）,
        "mismatches": 見つかった不一致**全部**のリスト（`ailine verify` の人間向け報告が
                      使う・憲法1: 修正箇所への誘導は全所見を集めてから言う）}
       ★ デモ撮影のリハ（2026-08-21）の直し: 行数が一致すれば、Σ と帰属は独立に検査できる
       （帰属は出力の各行を自分の申告する元セルと突き合わせるだけ・Σ の合否に依存しない）
       ── 最初の不一致で打ち切らず両方走らせる。行数だけは別（行が消えていれば Σ も帰属も
       意味を成さないため、従来どおり単独で早期に返す）。

       ★★ 2026-08-24 の根治（盲検の契約レビュー）: 旧版は元側のファイル一覧を
       **出力自身の出所列**から作っていた。だからフォルダに在るのに積まれなかった冊は
       元側にも現れず、**出力を出力自身と比べていた** ── 「道具を信じる代わりに使う
       独立チェック」が、一番肝心な「冊が丸ごと落ちた」を原理的に見られなかった。
       実測: 3 冊のうち 1 冊が見出しの綴り違いで積まれず、それでも
       「行数 元3/出力3・Σ 元600/出力600・exit 0」。
       → **分母はフォルダの実ファイルから作る**。出所列に現れない冊は
       `kind: "missing_source"` として名指しする。
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
    # ★ operator 盲検7度目の直し: 合計行検出は数値列すべて（旧: 最初の1本だけ）。

    # ★★ 分母は**フォルダの実ファイル**から作る（出力の出所列からではない）。
    #   出所列に一度も現れない冊＝丸ごと落ちた冊を、ここで初めて見られるようになる。
    missing_sources = []
    if src_folder is not None and Path(src_folder).is_dir():
        from ailine_core import multifile as _mf, stack as _st
        folder_files, _exc = _mf.classify_folder_contents(Path(src_folder))
        for q in folder_files:
            if q.suffix.lower() != OPENPYXL_READABLE_SUFFIX:
                continue
            if _st.is_own_output(q):
                continue          # 自分の出力は入力ではない（cmd_stack と同じ判定）
            if q.name not in refs:
                refs.setdefault(q.name, [])   # 分母に入れる（行は 0 件）
                missing_sources.append(q.name)

    expected_total = 0
    sums_source = {name: 0.0 for name in numeric_cols}
    unbacked = []          # ★ 検算が裏取りしていない除外（複数ファイルの盲検・致命①⑨）
    _LAST_DROPPED.clear()
    _UNREADABLE.clear()
    for fname in sorted(refs):
        path = src_folder / fname
        if not path.exists():
            continue
        expected_rows, values = _expected_rows_for_source(path, base_headers,
                                                            label_col_name, numeric_cols,
                                                            sheet_name=base_sheet_name)
        for _r in _LAST_DROPPED.get(str(path), []):
            unbacked.append({"kind": "unbacked_exclusion", "file": fname, "row": _r})
        expected_total += len(expected_rows)
        for name in numeric_cols:
            for v in values.get(name, {}).values():
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    sums_source[name] += v

    actual_total = len(out_rows)
    row_count = {"source": expected_total, "output": actual_total}
    # ★ 冊が丸ごと落ちているなら、**行数の差より先にその事実を名指しする**。
    #   「元 5 / 出力 3」だけでは、どの冊が落ちたのか人には分からない
    #   （盲検の使い勝手レビューでも「どのファイルが原因か一言も言わない」が致命だった）。
    ms_mismatches = [{"kind": "missing_source", "column": None, "name": n,
                       "source": n, "output": None} for n in sorted(missing_sources)]
    # ★ 致命⑥: 読めなかった冊は「無かったこと」にしない ── 不一致として名指しする
    #   （分母から黙って消えると、冊が丸ごと落ちたフォルダに ✓ が出る）。
    for _p, _why in sorted(_UNREADABLE.items()):
        ms_mismatches.append({"kind": "unreadable_source", "column": None,
                               "name": Path(_p).name, "source": _why, "output": None})
    if expected_total != actual_total:
        rc_mismatch = {"kind": "row_count", "column": None,
                       "source": expected_total, "output": actual_total}
        first = ms_mismatches[0] if ms_mismatches else rc_mismatch
        return {"row_count": row_count, "sums": {}, "mismatch": first,
                "mismatches": ms_mismatches + [rc_mismatch]}
    if ms_mismatches:
        # 行数が偶然一致していても、積まれていない冊が在るなら通さない
        return {"row_count": row_count, "sums": {}, "mismatch": ms_mismatches[0],
                "mismatches": ms_mismatches}

    sums_output = {name: 0.0 for name in numeric_cols}
    for i, name in enumerate(base_headers, start=1):
        if name not in numeric_cols:
            continue
        for r in out_rows:
            v = grid.get((r, i))
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                sums_output[name] += v

    sums = {name: {"source": sums_source[name], "output": sums_output[name]} for name in numeric_cols}
    # ★ デモ撮影のリハで発覚（2026-08-21 13:2x）: 憲法1（誘導）+ 一括検出 ── 最初の不一致
    #   （Σ）で打ち切ると、帰属検算が名指しできるはずの「どの行がいくつ→いくつ」を
    #   一度も計算しないまま終わっていた。Σ と帰属は独立に検査できる ── 両方走らせて
    #   全所見を集める。★ 書き込み時事後条件（stack/cmd_stack）の挙動は変えない ──
    #   `mismatch`（単数）は従来どおり「最初に見つかった1件」（優先順位も従来と同じ）。
    #   追加する `mismatches`（複数）だけが新しい。
    mismatches = []
    for name in numeric_cols:
        if abs(sums_source[name] - sums_output[name]) > TOLERANCE:
            mismatches.append({"kind": "sum", "column": name,
                               "source": sums_source[name], "output": sums_output[name]})

    attribution = _attribution_mismatch(out_data, base_headers, out_rows, src_folder,
                                         sheet_name=base_sheet_name)
    if attribution:
        mismatches.append({"kind": "attribution", "column": attribution["column"],
                           "source": attribution["source_value"], "output": attribution["output_value"],
                           "file": attribution["file"], "src_row": attribution["src_row"]})

    # ★ operator 盲検7度目 修正2: 語のトリップワイヤ（第二の独立検出器）。
    #   cmd_stack の書き込み時事後条件は kind=="attribution" だけを見るので、これが
    #   `mismatch`（単数）に選ばれても書き込みは止まらない（除外しない設計と一致）。
    #   `ailine verify` 単独実行（cmd_verify）は kind を問わず exit 5 にする ── 意図どおり。
    mismatches.extend(_total_word_mismatches(out_data, out_headers, out_rows))

    # ★ 2026-08-26（致命①⑨）: 書き手と同じ規則で落ちた行は、検算が裏取りしたものではない。
    #   数字が合っていても黙らない ── どのファイルの何行目を落としたかを人へ見せる。
    #   ★ 判定（exit）は変えない: 正しい合計行を持つ表を全部不合格にしないため。
    #     ただし「一致しました」と言い切らせない（開示が出た回は ✓ を名乗らない）。
    return {"row_count": row_count, "sums": sums,
            "mismatch": mismatches[0] if mismatches else None, "mismatches": mismatches,
            "unbacked_exclusions": unbacked}


# ---- M3: 照合出力の単独検算（`ailine verify <出力> <元A> <元B>`） ----


def _display_key_for_report(nk) -> str:
    """normalize_key の戻り値 → 報告用の表示文字列。None（キー不明）は match.py の
       第5区分ラベルをそのまま使う。"""
    return match.UNKNOWN_KEY_LABEL if nk is None else nk[1]


def verify_match_output(out_path, book_a, book_b) -> dict:
    """M3（照合）出力の単独検算の入口。出力ブックの印（creator）・焼いた条件
       （dc:description の JSON: キー列/金額列/両冊のヘッダー）を読み、原本2冊を
       xml_readback で独立に読み直してキー勘定（件数・Σ・差額）を再集計、照合シートの
       全行（キー不明含む）と突き合わせる。3冊とも読むだけ。
       ★ 印が違う・条件が焼かれていない/壊れている場合は {"unmarked"/"unsupported": ...}
       を返す（空虚な合格の禁止・呼び出し側が exit 4 にする）。"""
    creator, description = xml_readback.read_core_properties(out_path)
    if creator != match.CREATOR_MARK:
        return {"unmarked": True}
    try:
        cond = json.loads(description) if description else None
    except (TypeError, ValueError):
        cond = None
    if not (isinstance(cond, dict) and cond.get("tool") == "ailine" and cond.get("kind") == "match"):
        return {"unsupported": "照合出力に条件（dc:description）が焼かれていません"
                                "（壊れているか、ailine match 以外が作った可能性があります）。"
                                "検算できません。"}
    key_a, key_b = cond.get("key_a"), cond.get("key_b")
    amount_a, amount_b = cond.get("amount_a"), cond.get("amount_b")
    headers_a, headers_b = cond.get("headers_a"), cond.get("headers_b")
    if not (key_a and key_b and amount_a and amount_b
            and isinstance(headers_a, list) and isinstance(headers_b, list)):
        return {"unsupported": "照合出力の条件が不完全です（キー列/金額列/見出しの一部が"
                                "欠けています）。検算できません。"}
    return _verify_match(out_path, book_a, book_b, key_a, key_b, amount_a, amount_b,
                          headers_a, headers_b)


def _verify_match(out_path, book_a, book_b, key_a, key_b, amount_a, amount_b,
                   headers_a, headers_b) -> dict:
    """検算の本体。★ 片配線の自己点検: cmd_run_match（ailine.py）の書き込み時事後条件と
       同じ independent_key_stats（ailine_core/match.py）を使う ── キー正規化（前後空白
       除去のみ・型が違えば別キー）がここでズレると恒真検査になり偽陽性を生む。"""
    data_a = xml_readback.read_grid(book_a)
    header_row_a = _find_header_row(data_a, headers_a)
    data_b = xml_readback.read_grid(book_b)
    header_row_b = _find_header_row(data_b, headers_b)
    if header_row_a is None or header_row_b is None:
        which = "A" if header_row_a is None else "B"
        return {"unsupported": f"元{which}に条件の見出しが見つかりません"
                                "（別の元ファイルが渡された可能性があります）。検算できません。"}

    a_headers_x = xml_readback.header_names(data_a, header_row=header_row_a)
    a_rows_x = [r for r in xml_readback.data_row_numbers(data_a, header_row_a)
                if xml_readback.row_has_any_value(data_a, r, len(a_headers_x))]
    b_headers_x = xml_readback.header_names(data_b, header_row=header_row_b)
    b_rows_x = [r for r in xml_readback.data_row_numbers(data_b, header_row_b)
                if xml_readback.row_has_any_value(data_b, r, len(b_headers_x))]

    a_stats = match.independent_key_stats(data_a["grid"], a_rows_x, a_headers_x, key_a, amount_a)
    b_stats = match.independent_key_stats(data_b["grid"], b_rows_x, b_headers_x, key_b, amount_b)

    out_data = xml_readback.read_grid(out_path, sheet_name=match.MATCH_SHEET_NAME)
    out_rows = xml_readback.data_row_numbers(out_data, header_row=1)
    grid = out_data["grid"]

    mismatches = []
    seen_keys = set()
    for r in out_rows:
        key_val = grid.get((r, 1))
        a_count_v = grid.get((r, 2), 0) or 0
        a_sum_v = grid.get((r, 3), 0) or 0
        b_count_v = grid.get((r, 4), 0) or 0
        b_sum_v = grid.get((r, 5), 0) or 0
        diff_v = grid.get((r, 6), 0) or 0
        nk = None if key_val == match.UNKNOWN_KEY_LABEL else match.normalize_key(key_val)
        seen_keys.add(nk)
        a_entry = a_stats.get(nk, {"count": 0, "sum": 0.0})
        b_entry = b_stats.get(nk, {"count": 0, "sum": 0.0})
        if a_count_v != a_entry["count"]:
            mismatches.append({"kind": "count", "side": "A", "key": key_val,
                               "expected": a_entry["count"], "written": a_count_v})
        if b_count_v != b_entry["count"]:
            mismatches.append({"kind": "count", "side": "B", "key": key_val,
                               "expected": b_entry["count"], "written": b_count_v})
        if abs(a_sum_v - a_entry["sum"]) > match.TOLERANCE:
            mismatches.append({"kind": "sum", "side": "A", "key": key_val,
                               "expected": a_entry["sum"], "written": a_sum_v})
        if abs(b_sum_v - b_entry["sum"]) > match.TOLERANCE:
            mismatches.append({"kind": "sum", "side": "B", "key": key_val,
                               "expected": b_entry["sum"], "written": b_sum_v})
        if abs(diff_v - (a_sum_v - b_sum_v)) > match.TOLERANCE:
            mismatches.append({"kind": "diff", "key": key_val,
                               "expected": a_sum_v - b_sum_v, "written": diff_v})

    # ★ 行の過不足（キーの欠落/捏造）: 独立再集計に居るのに出力に無い／出力に居るのに
    #   独立再集計のどちらにも居ない、を両方見る（一括検出・最初の1件で止めない）。
    all_source_keys = set(a_stats) | set(b_stats)
    for nk in sorted(all_source_keys - seen_keys, key=lambda k: _display_key_for_report(k)):
        mismatches.append({"kind": "missing_key", "key": _display_key_for_report(nk)})
    for nk in sorted(seen_keys - all_source_keys, key=lambda k: _display_key_for_report(k)):
        mismatches.append({"kind": "extra_key", "key": _display_key_for_report(nk)})

    a_total_sum = sum(v["sum"] for v in a_stats.values())
    b_total_sum = sum(v["sum"] for v in b_stats.values())
    return {"ok": not mismatches, "mismatches": mismatches,
            "sums": {"A": a_total_sum, "B": b_total_sum}, "keys": len(out_rows)}
