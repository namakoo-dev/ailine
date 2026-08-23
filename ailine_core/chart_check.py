"""chart_check — グラフ段の事後条件②本体（zip+ElementTree で chart XML を直読み）。

★ spike 実測（2026-08-23）: chart XML は標準 OOXML ── 種別は c:plotArea 直下の子タグ名
  （barChart/lineChart/pieChart）、参照は c:ser/c:val/c:numRef/c:f（値列）・
  c:ser/c:cat/c:strRef/c:f（項目列）。この XPath は3種共通（kind 非依存）。
  読むのは**参照のみ**（タイトル/dLbls 等の見た目要素はスタイリング有無で変わるので見ない）。

★ 恒真殺し: 「グラフ数が +1」だけでは、意図した列を描いていない/種別が違うグラフでも
  ✓ が出てしまう（ailine.check_chart はそこまでしか見ない・旧実装のまま残す）。
  この module は「その1個のグラフが、頼んだ種別で、頼んだ値列を指しているか」まで見る。

★ ailine を import しない（xml_readback.py と同じ移植可能性の作法）。読むだけ ── 書き込み無し。
"""
from __future__ import annotations

import re
import zipfile
from xml.etree import ElementTree as ET

_NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
_KIND_TAGS = {"bar": "barChart", "line": "lineChart", "pie": "pieChart"}
_CHART_XML_RE = re.compile(r"xl/charts/chart\d+\.xml", re.IGNORECASE)
_COL_REF_RE = re.compile(r"\$([A-Za-z]+)\$\d+")


def _chart_xml_paths(z: zipfile.ZipFile) -> list:
    """xl/charts/chartN.xml のパス一覧（style/colors 等の付随ファイルは除く）。"""
    return sorted(n for n in z.namelist() if _CHART_XML_RE.fullmatch(n))


def _ref_column(f_text: str) -> str | None:
    """"Sheet!$B$2:$B$4" のような参照文字列から列記号(大文字)を取り出す。読めなければ None。"""
    m = _COL_REF_RE.search(f_text or "")
    return m.group(1).upper() if m else None


def check_chart_series(path, kind: str, value_col_letter: str,
                        category_col_letter: str | None = None) -> tuple:
    """(status, reason)。status ∈ {"pass", "fail"}。

    ① 種別（c:plotArea 直下の子タグ）が kind と一致するか
    ② 値列の参照（c:ser/c:val/c:numRef/c:f）が value_col_letter と一致するか
    ③ category_col_letter を渡した場合のみ、項目列の参照（c:ser/c:cat/c:strRef/c:f）も見る

    読めない/該当が無ければ fail（読めない時に pass へ倒れない ── check_extract 等と同じ線）。
    """
    want_tag = _KIND_TAGS.get(kind)
    if want_tag is None:
        return "fail", f"未知のグラフ種類『{kind}』（bar/line/pie のいずれかのはず）"
    try:
        with zipfile.ZipFile(path) as z:
            chart_paths = _chart_xml_paths(z)
            if not chart_paths:
                return "fail", "グラフが見つからない（chart XML が無い）"
            for chart_path in chart_paths:
                root = ET.fromstring(z.read(chart_path))
                plot_area = root.find("c:chart/c:plotArea", _NS)
                if plot_area is None:
                    continue
                found_tag = None
                found_el = None
                for tag in _KIND_TAGS.values():
                    el = plot_area.find(f"c:{tag}", _NS)
                    if el is not None:
                        found_tag, found_el = tag, el
                        break
                if found_el is None:
                    continue   # plotArea はあるが既知の3種のどれでもない → 次の chart を見る
                if found_tag != want_tag:
                    return "fail", f"種別が『{kind}』でなく『{found_tag}』になっている"
                ser = found_el.find("c:ser", _NS)
                if ser is None:
                    return "fail", "系列(c:ser)が見つからない"
                val_f = ser.find("c:val/c:numRef/c:f", _NS)
                if val_f is None or not val_f.text:
                    return "fail", "値列の参照(c:val)が見つからない"
                actual_val_col = _ref_column(val_f.text)
                if actual_val_col != str(value_col_letter).upper():
                    return "fail", (
                        f"値列の参照が『{value_col_letter}』列でなく『{actual_val_col}』列に"
                        f"なっている（{val_f.text}）")
                if category_col_letter is not None:
                    cat_f = ser.find("c:cat/c:strRef/c:f", _NS)
                    if cat_f is None or not cat_f.text:
                        return "fail", "項目列の参照(c:cat)が見つからない"
                    actual_cat_col = _ref_column(cat_f.text)
                    if actual_cat_col != str(category_col_letter).upper():
                        return "fail", (
                            f"項目列の参照が『{category_col_letter}』列でなく"
                            f"『{actual_cat_col}』列になっている（{cat_f.text}）")
                return "pass", f"種別『{kind}』・値列『{value_col_letter}』の参照を確認"
            return "fail", "グラフの種別タグ（bar/line/pieChart）が読めない"
    except Exception as e:
        return "fail", f"グラフ検証に失敗: {type(e).__name__}: {e}"
