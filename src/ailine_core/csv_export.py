"""csv_export ── CSV_EXPORT: `ailine export-csv`。csv_quarantine の逆方向。
   DESIGN-20260824-format-map.md「CSV_EXPORT の憲法」の実装。

★ 検疫の逆方向: csv_quarantine（CSV→xlsx）は「1セルも変えずに読んだ」を、全セルを
  文字列として読み型を機械決定することで保証した。ここ（xlsx→CSV）は「1セルも変えずに
  書いた」を、書いた CSV を読み戻して元シートと突き合わせる（欠落/不一致/余剰の3計数）
  ことで保証する ── 主張の形は同じ、向きだけが逆。

★ ailine を import しない（tests/test_line_budget.py と同じ移植可能性の作法）。依存は
  stdlib + ailine_core.xml_readback のみ（openpyxl は使わない ── 読みは独立読み実装
  そのものに乗る。csv_quarantine 側が openpyxl で「書く」のと非対称なのは、こちらは
  「読んで CSV へ書く」だけで xlsx を新たに作らないため openpyxl が要らないから）。

## 0 落ちを作らない（★ 設計の芯）
  xml_readback.read_grid は t 属性（s/inlineStr/str = 文字列・既定 = 数値）をそのまま
  Python の型（str/int/float/date）へ落として返す ── つまり「文字列として保持されている
  セル」（先頭ゼロ・16桁品番等）は最初から Python の str として届く。ここでは str 型の
  値を CSV 上で常に引用符で囲む（csv.QUOTE_NONNUMERIC）ことで、読み戻した側が誤って
  数値へ変換するのを防ぐ。int/float はそのまま引用せずに書く。
  ★ 数式セルは xml_readback が既にキャッシュ値（<v>）を返す ── 「表示されている値」を
  書くという設計の要求は、read_grid を読むだけで自動的に満たされる。

## 文字コード
  既定は utf-8（BOM 付き・Excel が誤ってANSI/cp932と誤認しないため）。cp932 も選べる
  （会計ソフト向け）。cp932 選択時は BOM を付けない（Shift_JIS に標準の BOM 慣行が無い）。
"""
from __future__ import annotations

import csv
import datetime
import io
from dataclasses import dataclass

from ailine_core import xml_readback

# ★ Excel 想定で改行は CRLF 既定（会計ソフト向け・設計文書の指示どおり）。
LINE_TERMINATOR = "\r\n"

# ★ 引用の規則（開示文の1行と同期。変えたらここと _render 側の両方を直す）。
QUOTING_DISCLOSURE = ("引用: 数値以外の値（文字列・日付等）は常に引用符で囲みます"
                      "（区切り・改行・引用符を含む値も安全に扱えます）")


class EncodingWriteError(ValueError):
    """選んだ文字コードで書けない文字があった（cp932 で書けない漢字等）。"""


@dataclass(frozen=True)
class EncodingChoice:
    """label: 開示に使う表示名。codec: Python の codec 名。bom: BOM を付けるか。"""
    label: str
    codec: str
    bom: bool


# ★ 第一波: utf-8(既定・BOM付き)/cp932(会計ソフト向け・BOM無し)の2択。他の符号化は
#   資料も凍結検体も無いので対象外（憲法「様式は人が作る」と同じ理由 ── 機械が対象外の
#   符号化を黙って類推しない）。
_ENCODING_ALIASES = {
    "utf-8": EncodingChoice(label="utf-8", codec="utf-8", bom=True),
    "utf8": EncodingChoice(label="utf-8", codec="utf-8", bom=True),
    "cp932": EncodingChoice(label="cp932", codec="cp932", bom=False),
    "shift_jis": EncodingChoice(label="cp932", codec="cp932", bom=False),
    "shift-jis": EncodingChoice(label="cp932", codec="cp932", bom=False),
    "sjis": EncodingChoice(label="cp932", codec="cp932", bom=False),
}


def resolve_encoding(requested: str | None) -> EncodingChoice | None:
    """--encoding の値 → EncodingChoice。既定(None/省略)は utf-8。未知の指定は None
       （呼び出し側が「対応していない」とエラーにする）。"""
    key = (requested or "utf-8").strip().lower()
    return _ENCODING_ALIASES.get(key)


@dataclass(frozen=True)
class SourceGrid:
    """read_source が返す、独立読み実装(xml_readback)で読んだシートの中身。
       grid: {(行,列)1起点: 値(str/int/float/date/bool)}。空セルは含まない。"""
    grid: dict
    max_row: int
    max_col: int
    sheet_fallback: bool
    # ★ 2026-08-26（致命3）: 数式だがキャッシュ値が無いセル。空セルと同じ扱いにすると
    #   分母（declared）から消え、空欄で書き出しても「欠落 0」が成立してしまう。
    uncached_formulas: tuple = ()


def read_source(path, sheet_name: str) -> SourceGrid:
    """book の sheet_name を xml_readback.read_grid で読む（openpyxl を使わない独立読み）。
       ★ 数式セルはキャッシュ値（<v>）がそのまま入る＝「表示されている値」を書くという
       要求を、read_grid を読むだけで満たす。"""
    data = xml_readback.read_grid(path, sheet_name=sheet_name)
    return SourceGrid(grid=data["grid"], max_row=data["max_row"], max_col=data["max_col"],
                       sheet_fallback=data["sheet_fallback"],
                       uncached_formulas=tuple(data.get("uncached_formulas") or ()))


def _cell_for_csv(value):
    """xml_readback 型の値 → csv.writer に渡す python 値。int/float はそのまま
       （QUOTE_NONNUMERIC が非引用で書く）。str/date/bool は str へ寄せる（常に引用される）。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return value
    # ★ 2026-08-26: 時刻は時刻として書く（datetime.time は date の下位型ではないので
    #   下の分岐に吸われないが、順序を明示しておく ── 読み手の誤解を残さない）。
    if isinstance(value, datetime.time):
        return value.isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    return str(value)


@dataclass(frozen=True)
class WriteResult:
    """declared: {(行,列): 元の型つき値}（source 由来。空セルは含まない ── xml_readback の
       grid が最初から空セルを持たないのでそのまま流用できる）。raw_bytes: 実際に書いた
       バイト列（呼び出し側がファイルへ書く／エンコード失敗時は例外で届く）。"""
    declared: dict
    rows_written: int
    cols_written: int
    raw_bytes: bytes


def build_csv(grid: SourceGrid, enc: EncodingChoice) -> WriteResult:
    """grid を CSV 本文（bytes）へ組み立てる。★ 数値化しない: 元が str 型のセルは
       QUOTE_NONNUMERIC で必ず引用され、CSV を素朴に読む側（会計ソフト等）が誤って
       数値化しても先頭ゼロが落ちないよう文字として残る。
       選んだ符号化で書けない文字があれば EncodingWriteError を投げる（黙って落とさない・
       黙って別文字に置換しない）。"""
    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_NONNUMERIC, lineterminator=LINE_TERMINATOR)
    declared: dict = {}
    for r in range(1, grid.max_row + 1):
        row_out = []
        for c in range(1, grid.max_col + 1):
            v = grid.grid.get((r, c))
            row_out.append(_cell_for_csv(v))
            if v is not None:
                declared[(r, c)] = v
        writer.writerow(row_out)
    text = buf.getvalue()
    try:
        raw = text.encode(enc.codec)
    except UnicodeEncodeError as e:
        raise EncodingWriteError(
            f"指定した文字コード『{enc.label}』で書けない文字があります: "
            f"{e.object[e.start:e.end]!r} (U+{ord(e.object[e.start]):04X})"
        ) from e
    if enc.bom:
        raw = b"\xef\xbb\xbf" + raw
    return WriteResult(declared=declared, rows_written=grid.max_row, cols_written=grid.max_col,
                        raw_bytes=raw)


@dataclass(frozen=True)
class RoundtripResult:
    """missing: declared にあるが読み戻しに無い (row,col)。
       mismatched: 両方にあるが値が違う (row,col,declared値,actual読み戻しテキスト)。
       surplus: 読み戻しにあるが declared に無い (row,col,actual読み戻しテキスト)。"""
    missing: list
    mismatched: list
    surplus: list

    @property
    def ok(self) -> bool:
        return not (self.missing or self.mismatched or self.surplus)


def _agree(declared, actual_text: str) -> bool:
    """declared（xml_readback 型）と、書いた CSV を読み戻した生テキストが一致するか。
       CSV にはそもそも型が無い ── declared 側の型に応じてテキストを解釈し直して比べる
       （TOLERANCE 不使用・csv_quarantine._values_agree と同じ規律）。"""
    if isinstance(declared, bool):
        return actual_text in ("TRUE", "FALSE") and (actual_text == "TRUE") == declared
    if isinstance(declared, (int, float)):
        try:
            v = float(actual_text)
        except ValueError:
            return False
        return v == float(declared)
    if isinstance(declared, datetime.date):
        return actual_text == declared.isoformat()
    if isinstance(declared, str):
        return actual_text == declared
    return False


def verify_roundtrip(declared: dict, raw_bytes: bytes, enc: EncodingChoice) -> RoundtripResult:
    """書いた CSV バイト列を同じ符号化で読み戻し、declared と3計数で突き合わせる。
       ★ newline='' で csv.reader に渡す（python csv モジュールの推奨作法・CRLF を
       二重に解釈させない／引用内改行を壊さない）。"""
    decode_codec = enc.codec + "-sig" if enc.codec == "utf-8" else enc.codec
    text = raw_bytes.decode(decode_codec)
    rows = list(csv.reader(io.StringIO(text, newline="")))

    actual: dict = {}
    for r, row in enumerate(rows, start=1):
        for c, cell_text in enumerate(row, start=1):
            if cell_text != "":
                actual[(r, c)] = cell_text

    missing = []
    mismatched = []
    for key, dval in declared.items():
        if key not in actual:
            missing.append(key)
            continue
        if not _agree(dval, actual[key]):
            mismatched.append((key[0], key[1], dval, actual[key]))
    surplus = [(r, c, v) for (r, c), v in actual.items() if (r, c) not in declared]

    return RoundtripResult(missing=missing, mismatched=mismatched, surplus=surplus)
