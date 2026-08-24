# 第三波 #9 ── openpyxl の往復で消える図形（描かれた角印・社判）。
#
# ★ この repo の中で実測した事実（推測ではない）:
#     画像（xdr:pic / xl/media）                    → load→save で **残る**
#     画像でない図形（sp・テキストボックス・オートシェイプ）→ load→save で **消える**
#   `ailine export-pdf` は指定シートの抽出とページ設定のため原本を openpyxl で
#   一度書き直すので、描かれた角印が PDF から消える。出来上がりは完成品に見える。
#
# 契約:
#   ① 図形が在れば、変換の**前に**名指しする（消えたものは差分に出ない）
#   ② 画像だけの冊では 1 文字も警告しない（誤爆防止）
#   ③ 検出は zip の drawing XML を直接読む ── openpyxl に聞くと、openpyxl が
#      読めない図形は最初から見えないので恒真になる（別実装で測る）

import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
from ailine_core import pdf_export  # noqa: E402

NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_SHAPE = ('<oneCellAnchor><from><col>5</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from>'
          '<ext cx="500000" cy="500000"/>'
          '<sp macro="" textlink=""><nvSpPr><cNvPr id="99" name="KAKUIN"/><cNvSpPr/></nvSpPr>'
          f'<spPr><a:prstGeom xmlns:a="{NS}" prst="ellipse"/></spPr></sp><clientData/></oneCellAnchor>')


def _png(w=8, h=8):
    import struct, zlib
    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xffffffff)
    raw = b"".join(b"\x00" + bytes((200, 0, 0)) * w for _ in range(h))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _book_with_image(tmp_path):
    from openpyxl.drawing.image import Image as XLImage
    seal = tmp_path / "seal.png"
    seal.write_bytes(_png())
    p = tmp_path / "img.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "請求書"
    wb.active.add_image(XLImage(str(seal)), "D2")
    wb.save(p)
    return p


def _book_with_shape(tmp_path):
    """画像の冊に、画像でない図形を 1 つ手で足す（LibreOffice/Excel が作る形の代用）。"""
    base = _book_with_image(tmp_path)
    p = tmp_path / "shape.xlsx"
    shutil.copy2(base, p)
    tmp = str(p) + ".t"
    with zipfile.ZipFile(p) as zi, zipfile.ZipFile(tmp, "w") as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == "xl/drawings/drawing1.xml":
                data = data.decode("utf-8").replace("</wsDr>", _SHAPE + "</wsDr>").encode("utf-8")
            zo.writestr(item, data)
    import os
    os.replace(tmp, p)
    # ★ 治具が本当に仕込めたかを先に確かめる（初版は仕込めておらず、自分の壊れた
    #   治具を測って「消えない」と誤結論しかけた）。
    with zipfile.ZipFile(p) as z:
        assert "KAKUIN" in z.read("xl/drawings/drawing1.xml").decode("utf-8"), "治具が仕込めていない"
    return p


# --- 前提そのものの測定（この試験が守る事実）------------------------------------------

def test_openpyxl_roundtrip_keeps_images_but_drops_shapes(tmp_path):
    """★ 前提の凍結: openpyxl の版が変わってこの事実が変われば、ここが最初に鳴る。"""
    shaped = _book_with_shape(tmp_path)
    out = tmp_path / "after.xlsx"
    openpyxl.load_workbook(shaped).save(out)
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        assert any("media" in n for n in names), "画像まで落ちた（前提が変わった）"
        drawing = z.read("xl/drawings/drawing1.xml").decode("utf-8") if \
            "xl/drawings/drawing1.xml" in names else ""
    assert "KAKUIN" not in drawing, \
        "図形が残った ── 前提が変わったので、警告そのものを見直すこと"


# --- ①② 検出器 ----------------------------------------------------------------------

def test_detects_the_vanishing_shape(tmp_path):
    assert pdf_export.vanishing_shapes(_book_with_shape(tmp_path)) == ["KAKUIN"]


def test_image_only_book_is_silent(tmp_path):
    assert pdf_export.vanishing_shapes(_book_with_image(tmp_path)) == [], \
        "画像しか無い冊に誤爆した（写真の印は消えないのに消えると言った）"


def test_unreadable_file_is_silent(tmp_path):
    p = tmp_path / "broken.xlsx"
    p.write_bytes(b"not a zip")
    assert pdf_export.vanishing_shapes(p) == []


def test_warning_names_the_shape_and_offers_a_way_out(tmp_path):
    lines = pdf_export.vanishing_shapes_warning(["KAKUIN"])
    text = "\n".join(lines)
    assert "KAKUIN" in text and "消えます" in text
    assert "LibreOffice" in text or "Excel" in text, f"逃げ道が無い: {text}"
    assert pdf_export.vanishing_shapes_warning([]) == []


# --- ③ 別実装であること -----------------------------------------------------------------

def test_detection_does_not_ask_openpyxl():
    src = (REPO / "src" / "ailine_core" / "pdf_export.py").read_text(encoding="utf-8")
    body = src[src.index("def vanishing_shapes("):]
    # 物差しの訂正（封印者ナギ）: 初版は "openpyxl" の文字列を見ていて docstring の
    # 説明文まで掴んだ（測定器が粗い）。意図は「openpyxl に聞いていない」なので
    # 呼び出しの形で測る。assert の意図は変えていない。
    assert "load_workbook" not in body.split("def vanishing_shapes_warning")[0], \
        "openpyxl に聞いている（読めない図形は見えないので恒真になる）"
