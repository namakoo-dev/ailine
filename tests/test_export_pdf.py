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

# ── ★ 2026-09-16: 案内は**そのまま打てる**こと（盲検の買い手役⑥）────────────────────
#
# ★★ 事故（買い手の引用）: 言われたとおりに打つと必ず止まっていた。
#     × …14 個が PDF の中に見つかりません → `--fit-to-width` を付けて出し直してください
#     $ ailine export-pdf 9月売上.xlsx --fit-to-width
#     × 出力先 …9月売上.pdf が既にあります。別名にするなら --out、上書きしてよければ --overwrite を
#   ★ PDF は**検算より先に**書かれるので、失敗した 1 回目が出力先を塞ぐ。
#     `--overwrite` も要ることは 1 回目の案内に書かれていなかった。
# ★★ もう 1 つ: `--fit-to-width` を付けた 2 回目も × だったとき、**次の一手が 1 行も出なかった**。
#   買い手「PDF はフォルダに出来ているが、渡していいのか判断できない」。

from ailine_core.pdf_export import next_step_for_missing   # noqa: E402


def test_the_advice_can_be_typed_as_printed_when_the_output_is_taken():
    """★ 事故そのもの ── 出力先が塞がっているなら、案内に --overwrite が入ること。"""
    lines = next_step_for_missing(fit_to_width=False, out_exists=True)
    joined = "".join(lines)
    assert "--fit-to-width" in joined and "--overwrite" in joined, lines


def test_the_advice_does_not_add_overwrite_when_the_output_is_free():
    """★ 陰性対照 ── 塞がっていない時に余計なフラグを勧めない（言われたとおりが最短であること）。"""
    lines = next_step_for_missing(fit_to_width=False, out_exists=False)
    joined = "".join(lines)
    assert "--fit-to-width" in joined and "--overwrite" not in joined, lines


def test_it_does_not_repeat_a_flag_that_is_already_on():
    """★ `--fit-to-width` を既に付けている回に、同じ物をもう一度勧めない。"""
    lines = next_step_for_missing(fit_to_width=True, out_exists=False)
    assert not any("--fit-to-width" in l for l in lines), lines


def test_it_never_goes_silent_when_the_flags_run_out():
    """★★ 打つ手が尽きた回に**黙らない** ── 「PDF は出来ている・目で確かめてよい」まで言う。
       ★ 沈黙は「まだ手がある」と読まれる。無いなら無いと言う。"""
    lines = next_step_for_missing(fit_to_width=True, out_exists=False, orientation="landscape")
    assert lines, lines
    joined = "".join(lines)
    assert "ここまで" in joined, joined
    assert "PDF 自体は出来ています" in joined, joined


def test_it_offers_landscape_before_giving_up():
    """★ 手が残っているうちは、その手を出す（諦めを先に言わない）。"""
    lines = next_step_for_missing(fit_to_width=True, out_exists=False)
    joined = "".join(lines)
    assert "landscape" in joined and "ここまで" not in joined, joined
