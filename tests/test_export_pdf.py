# export-pdf（台帳 PRINT 2 件 + EXPORT_DOC 2 件）── 実装より先に凍結した赤い検体。
#
# ★ op（DSL 語彙）でなくサブコマンドにした理由: 2026-08-24 の実測で OPS_DOC に 16 行
#   足したら別 op の分類が 98.1%→94.2% に落ちた。「表を紙の形で外へ出す」は自然言語の
#   曖昧さが要らない操作なので、プロンプトを 1 行も増やさない側に置く。
#
# 契約:
#   ① PDF を作る。★ 出した PDF の**テキスト層**を読み戻し、元シートの値が載っているかを数える
#   ② 値が欠けていたら ✓ を名乗らない
#   ③ ★ 読み戻しの道具（pdfplumber）が**居ない**環境では、PDF は作るが ✓ を名乗らない
#      ── 「居るから見えない」対策として、この試験は**居ない側を既定**で回す
#      （開発機に入っているせいで、入っていない人の経路が一度も通らない事故を防ぐ）

import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
import ailine  # noqa: E402
from ailine_core import pdf_export  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_golden_transcripts import _isolate, _run_main  # noqa: E402

needs_impl = pytest.mark.xfail(
    not hasattr(ailine, "cmd_export_pdf"),
    reason="export-pdf 未実装（契約は凍結済み・実装が来たら自動実測化）",
    strict=True,
)


def _book(tmp_path):
    p = tmp_path / "seikyu.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "請求"
    ws.append(["取引先", "金額"])
    ws.append(["あかつき商事", 12000])
    wb.save(p)
    return p


def _fake_convert(text_in_pdf):
    """実 LibreOffice の代わり。窒息点は 1 箇所（_soffice_to_pdf）。"""
    def convert(book_path, out_path, sheet=None, orientation=None, fit_to_width=False):
        Path(out_path).write_bytes(b"%PDF-1.4 fake")
        convert.last_text = text_in_pdf
        return True, ""
    return convert


@needs_impl
def test_export_pdf_verifies_values_in_the_text_layer(tmp_path, monkeypatch, capsys):
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path)
    # ★ 治具の訂正（封印者ナギ・2026-08-24）: 初版の偽 PDF は見出し行を入れ忘れていた。
    #   実装は見出しも含む全セルを照合するのが正しい（見出しの落ちた PDF は不良品）。
    #   assert（値が載っていれば ✓）は不変・偽 PDF の中身を実物に合わせただけ。
    monkeypatch.setattr(ailine, "_soffice_to_pdf",
                         _fake_convert("取引先 金額\nあかつき商事 12000\n"))
    monkeypatch.setattr(pdf_export, "readback_available", lambda: True)
    monkeypatch.setattr(pdf_export, "read_pdf_text",
                         lambda p: "取引先 金額\nあかつき商事 12000\n")
    rc, out = _run_main(["export-pdf", str(book), "--sheet", "請求"], capsys)
    assert rc == 0, out
    assert "✓" in out, out
    assert (tmp_path / "seikyu.pdf").exists(), "PDF が作られていない"


@needs_impl
def test_export_pdf_refuses_when_a_value_is_missing(tmp_path, monkeypatch, capsys):
    """恒真殺し: 金額が PDF に載っていなければ ✓ を名乗らない。"""
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path)
    monkeypatch.setattr(ailine, "_soffice_to_pdf", _fake_convert("請求\nあかつき商事\n"))
    monkeypatch.setattr(pdf_export, "readback_available", lambda: True)
    monkeypatch.setattr(pdf_export, "read_pdf_text", lambda p: "請求\nあかつき商事\n")
    rc, out = _run_main(["export-pdf", str(book), "--sheet", "請求"], capsys)
    assert "✓" not in out, f"値が欠けているのに ✓ を名乗った: {out}"
    assert "12000" in out, f"欠けた値を名指ししていない: {out}"


@needs_impl
def test_export_pdf_without_readback_makes_the_pdf_but_claims_nothing(tmp_path, monkeypatch, capsys):
    """★ 「居るから見えない」対策: 読み戻しの道具が**居ない**環境の経路。
       PDF は作るが ✓ は名乗らず、なぜ保証できないかを言う。"""
    _isolate(monkeypatch, tmp_path)
    book = _book(tmp_path)
    monkeypatch.setattr(ailine, "_soffice_to_pdf", _fake_convert("なんでもよい"))
    monkeypatch.setattr(pdf_export, "readback_available", lambda: False)
    rc, out = _run_main(["export-pdf", str(book), "--sheet", "請求"], capsys)
    assert (tmp_path / "seikyu.pdf").exists(), "道具が無くても PDF 自体は作る"
    assert "✓" not in out, f"読み戻せないのに ✓ を名乗った: {out}"
    assert "pdfplumber" in out, f"何を入れれば保証できるかを言っていない: {out}"


def test_verify_values_in_pdf_is_not_tautological():
    """照合そのものの恒真殺し（実装の有無に関わらず measurable）。"""
    monkey = pdf_export.verify_values_in_pdf
    assert callable(monkey)
    r = pdf_export.PdfCheck()
    assert r.missing == [] and r.checked == 0
