"""xml_readback — 検算専用の独立読み実装。DESIGN-20260821-multifile.md v2 §3⑦⑧。

★ ailine stack の事後条件②（数値列ごとの Σ 照合）と ailine verify サブコマンドは、
  どちらもこの module を通して数字を作る。zipfile + xml.etree.ElementTree で
  xl/worksheets/*.xml・xl/sharedStrings.xml・xl/workbook.xml を直読みし、
  **openpyxl を一切 import しない**。本体（書き込み・読み込み）は openpyxl で、
  検算は別実装 ── 同じ勘定を2箇所が違う実装で書く（_extract_predicate の作法の適用）。
  片方にしかないバグ（例: openpyxl の書き込み経路が値を丸める・列がずれる）を、
  検算側が独立に再現しなければ機械で検出できる。

★ ailine を import しない（tests/test_line_budget.py の移植可能性番人）。
★ 読むだけ ── この module にファイルへの書き込みは一切無い。
"""
from __future__ import annotations

import datetime
import re
import zipfile
from xml.etree import ElementTree as ET

_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
_R_ID_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_CELL_REF_RE = re.compile(r"([A-Za-z]+)(\d+)")

# ★ 赤1 の直し（2026-08-21 実機敵対検分）: 組み込みの日付/時刻 numFmtId（ECMA-376）。
#   これらは openpyxl 側が datetime/date として読む列 ── xml_readback も同じ列を
#   「数値」候補から外さないと、数値列の引き当てが stack(openpyxl) と食い違う。
_BUILTIN_DATE_FMT_IDS = set(range(14, 23)) | set(range(45, 48))
# 引用符内（"件" 等のリテラル接尾辞）を除いた書式コードに残る日付トークン。
_QUOTED_RE = re.compile(r'"[^"]*"')
_DATE_TOKEN_RE = re.compile(r"[ymdhYMDH]")
_EXCEL_EPOCH = datetime.datetime(1899, 12, 30)   # openpyxl と同じ基準日（1900年閏年バグ込み）


def _col_to_index(letters: str) -> int:
    """列記号（"A"・"AB" 等）→ 1起点の列番号。"""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n


def _parse_ref(ref: str) -> tuple:
    """セル参照（例 "B3"）→ (行, 列) の1起点タプル。"""
    m = _CELL_REF_RE.match(ref)
    if not m:
        raise ValueError(f"読めないセル参照: {ref!r}")
    return int(m.group(2)), _col_to_index(m.group(1))


def _load_shared_strings(z: zipfile.ZipFile) -> list:
    """sharedStrings.xml → 文字列のリスト（<si> の出現順。リッチテキストの<r><t>は連結）。"""
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall("main:si", _NS):
        # ★ 2026-08-24: `.//main:t` は **rPh（ふりがな）の <t> まで拾っていた**。
        #   日本語版 Excel は IME 入力した文字列に読み仮名を自動で埋めるので、
        #   「山田太郎」が「山田太郎ヤマダタロウ」になっていた（実測）。
        #   しかも export-csv の照合は declared 側も同じ read_grid 由来なので**恒真**で、
        #   誤った値を書いて ✓ を出していた。本文は <si> 直下の <t> と
        #   リッチテキストの <r>/<t> だけ ── rPh は本文ではない。
        texts = [t.text or "" for t in si.findall("main:t", _NS)]
        texts += [t.text or "" for t in si.findall("main:r/main:t", _NS)]
        out.append("".join(texts))
    return out


def _load_numfmts(z: zipfile.ZipFile) -> dict:
    """styles.xml の <numFmts> ── カスタム numFmtId → formatCode。"""
    if "xl/styles.xml" not in z.namelist():
        return {}
    root = ET.fromstring(z.read("xl/styles.xml"))
    out = {}
    for nf in root.findall(".//main:numFmts/main:numFmt", _NS):
        try:
            out[int(nf.get("numFmtId"))] = nf.get("formatCode") or ""
        except (TypeError, ValueError):
            continue
    return out


def _load_cell_xfs(z: zipfile.ZipFile) -> list:
    """styles.xml の <cellXfs> ── xf のインデックス（セルの s 属性）→ numFmtId。"""
    if "xl/styles.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/styles.xml"))
    xfs = root.find("main:cellXfs", _NS)
    if xfs is None:
        return []
    out = []
    for xf in xfs.findall("main:xf", _NS):
        try:
            out.append(int(xf.get("numFmtId", "0")))
        except (TypeError, ValueError):
            out.append(0)
    return out


def _is_date_numfmt(numfmt_id: int, custom_numfmts: dict) -> bool:
    """numFmtId が日付/時刻の書式か。組み込みID、または引用符を除いた書式文字列に
       日付トークン（y/m/d/h）を含むカスタム書式なら真（★ 赤1 の判定・DESIGN 追補）。"""
    if numfmt_id in _BUILTIN_DATE_FMT_IDS:
        return True
    code = custom_numfmts.get(numfmt_id)
    if not code:
        return False
    return bool(_DATE_TOKEN_RE.search(_QUOTED_RE.sub("", code)))


def _serial_to_date(serial: float):
    """Excel のシリアル値 → datetime.date（時刻成分が無ければ）/ datetime.datetime。
       ★ openpyxl と同じ基準日（1899-12-30）を使う ── 1900年を閏年扱いする Excel の
       バグごと踏襲することで、両実装の変換結果が一致する。"""
    dt = _EXCEL_EPOCH + datetime.timedelta(days=serial)
    return dt.date() if dt.time() == datetime.time(0, 0) else dt


def _sheet_target(z: zipfile.ZipFile, sheet_name: str | None) -> tuple:
    """workbook.xml + _rels からシート名 → 実体の worksheet xml パスを引く。
       sheet_name=None なら最初のシート。戻り値: (xml パス, 実際に使ったシート名, fallback)。
       ★ P2（architect 致命5・出荷済みの食い違い）: sheet_name を指定したのに同名シートが
       無ければ、無警告のまま1枚目へ落ちる（従来どおり例外は上げない）が、fallback=True で
       その旨を呼び出し側（verify.py・read_grid の戻り値経由）へ返せるようにする。"""
    wb_root = ET.fromstring(z.read("xl/workbook.xml"))
    sheets = wb_root.findall(".//main:sheets/main:sheet", _NS)
    if not sheets:
        raise ValueError("workbook.xml にシート定義が無い")
    chosen = sheets[0]
    fallback = False
    if sheet_name:
        found = False
        for s in sheets:
            if s.get("name") == sheet_name:
                chosen = s
                found = True
                break
        fallback = not found
    rid = chosen.get(_R_ID_ATTR)
    rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_root.findall("rel:Relationship", _NS):
        if rel.get("Id") == rid:
            target = (rel.get("Target") or "").lstrip("/")
            # ★ Target は「xl/ 相対」（例 "worksheets/sheet1.xml"）と「パッケージ絶対」
            #   （例 "/xl/worksheets/sheet1.xml"、openpyxl が書く形）の両方があり得る。
            path = target if target.startswith("xl/") else f"xl/{target}"
            return path, chosen.get("name"), fallback
    raise ValueError(f"シート {sheet_name!r} の関係(rels)が見つからない")


def _cell_value(c, shared: list, is_date: bool = False):
    """<c> 要素1個の値を型つきで取り出す（数値は int/float・文字列は str・空は None）。
       ★ 数式セル(<f>)はキャッシュ済みの<v>があればそれを使う。無ければ None
       （式キャッシュ欠けを「読めない」として扱う ── 他モジュールと同じ線）。
       ★ 赤1 の直し: is_date=True（セルの書式が日付/時刻）なら、数値の<v>を
       date/datetime へ変換して返す ── openpyxl(data_only) が同じセルを
       datetime として読むのに合わせる。isinstance(v,(int,float)) が False になり、
       『数値列』の候補にも Σ 対象にも入らなくなる（stack 側と一致）。"""
    t = c.get("t")
    v_el = c.find("main:v", _NS)
    if t == "s":
        if v_el is None or v_el.text is None:
            return None
        idx = int(v_el.text)
        return shared[idx] if 0 <= idx < len(shared) else None
    if t == "inlineStr":
        is_el = c.find("main:is", _NS)
        if is_el is None:
            return None
        return "".join(t2.text or "" for t2 in is_el.findall(".//main:t", _NS))
    if t == "str":
        return v_el.text if v_el is not None else None
    if t == "b":
        if v_el is None or v_el.text is None:
            return None
        return bool(int(v_el.text))
    # 既定（t属性無し、または "n"）: 数値（日付書式なら date/datetime へ）。
    if v_el is None or v_el.text is None:
        return None
    txt = v_el.text
    try:
        f = float(txt)
    except ValueError:
        return txt
    if is_date:
        try:
            return _serial_to_date(f)
        except (OverflowError, ValueError):
            pass   # 変換できない異常値は数値のまま返す（読めないものを黙って落とさない）
    return int(f) if f.is_integer() else f


def read_grid(path, sheet_name: str | None = None) -> dict:
    """xlsx を zipfile + ElementTree だけで読み、疎な grid を返す。
       戻り値: {"grid": {(行,列): 値}, "max_row": int, "max_col": int,
                "sheet_name": str（実際に読んだシート名）,
                "sheet_fallback": bool（sheet_name を指定したのに同名シートが無く
                1枚目へ落ちた時だけ True）}。
       ★ P2: sheet_name を渡せば基準名のシートを狙って読める（省略時は従来どおり先頭）。"""
    with zipfile.ZipFile(path) as z:
        shared = _load_shared_strings(z)
        custom_numfmts = _load_numfmts(z)
        cell_xfs = _load_cell_xfs(z)   # xf index(=セルの s属性) → numFmtId
        sheet_path, used_sheet_name, sheet_fallback = _sheet_target(z, sheet_name)
        root = ET.fromstring(z.read(sheet_path))
        grid: dict = {}
        max_row = max_col = 0
        for row_el in root.findall(".//main:sheetData/main:row", _NS):
            for c in row_el.findall("main:c", _NS):
                ref = c.get("r")
                if not ref:
                    continue
                r, col = _parse_ref(ref)
                try:
                    s_idx = int(c.get("s", "0"))
                except ValueError:
                    s_idx = 0
                numfmt_id = cell_xfs[s_idx] if 0 <= s_idx < len(cell_xfs) else 0
                is_date = _is_date_numfmt(numfmt_id, custom_numfmts)
                value = _cell_value(c, shared, is_date=is_date)
                if value is None:
                    continue
                grid[(r, col)] = value
                max_row = max(max_row, r)
                max_col = max(max_col, col)
        return {"grid": grid, "max_row": max_row, "max_col": max_col,
                "sheet_name": used_sheet_name, "sheet_fallback": sheet_fallback}


def header_names(data: dict, header_row: int = 1, max_scan_cols: int = 200) -> list:
    """header_row の先頭列から連続する非空セルを見出し名として読む
       （multifile.read_row_headers と同じ規則 ── 列の間に空白を挟む見出しは扱わない）。"""
    grid = data["grid"]
    names = []
    for c in range(1, max_scan_cols + 1):
        v = grid.get((header_row, c))
        if v in (None, ""):
            break
        names.append(str(v))
    return names


def data_row_numbers(data: dict, header_row: int) -> list:
    """header_row より下で、値を1つでも持つ行番号（昇順）。"""
    return sorted({r for (r, _c) in data["grid"] if r > header_row})


def row_has_any_value(data: dict, row: int, num_cols: int) -> bool:
    """row の 1..num_cols 列のどこかに値があるか。"""
    grid = data["grid"]
    return any((row, c) in grid for c in range(1, num_cols + 1))


def numeric_cells_became_strings(path, cell_refs, sheet_name: str | None = None) -> list:
    """cell_refs（例 ["C2","C3"]）のうち、実際に非数値の文字列になっているものだけを返す。

    ★ operator 盲検10度目 ⑤: `=B2*C2`（キャッシュ値は数値）の列を、openpyxl の素の
    .value（data_only=False）で見ると「文字列 '=B2*C2' に変わった」と誤検出される。
    数式セルはキャッシュ値（<v>）で判定すれば数値のまま ── ここは read_grid をそのまま
    再利用するだけ（数式か否かを問わず<v>の型で読む・二重実装しない）。
    キャッシュが無い（None）セルは「本当に文字列になった」と断定できないので含めない
    （保守的＝誤検知回避）。
    戻り値: 渡した順のうち、本当に非数値文字列だったものだけのサブセット。"""
    data = read_grid(path, sheet_name=sheet_name)
    grid = data["grid"]
    out = []
    for ref in cell_refs:
        row, col = _parse_ref(ref)
        v = grid.get((row, col))
        if isinstance(v, str):
            out.append(ref)
    return out


def read_core_properties(path) -> tuple:
    """docProps/core.xml の (dc:creator, dc:description) を直読みする。
       ★ M2（verify の種類判定・architect 致命3）: 検算の入口は『このブックは誰が何として
       書いたか』を先に読む ── creator が印（ailine stack / ailine extract）で、
       description に条件（EXTRACT の col/cmp/value）が機械可読で焼いてある。
       読めなければ (None, None)（壊れている/該当なし＝印なし＝他人のファイル扱い・
       fail closed）。★ ailine 非依存の汎用読み（この module の他の関数と同じく読むだけ）。"""
    try:
        with zipfile.ZipFile(path) as z:
            if "docProps/core.xml" not in z.namelist():
                return None, None
            root = ET.fromstring(z.read("docProps/core.xml"))
    except Exception:
        return None, None
    ns = {"dc": "http://purl.org/dc/elements/1.1/"}
    out = []
    for tag in ("dc:creator", "dc:description"):
        el = root.find(tag, ns)
        out.append(el.text if el is not None else None)
    return out[0], out[1]
