"""pdf_export ── PRINT / EXPORT_DOC: `ailine export-pdf`（台帳 4 件）。

出所: 3664021「請求書と領収書が印刷できる」/ 4260741「印刷ボタンの修正」/
5581892「PDF 化」。台帳では PRINT と EXPORT_DOC の 2 つの不足 op として数えていたが、
実体は同じ ── **表を、紙の形で外へ出す**。だから op（DSL の語彙）ではなく
サブコマンドにした。★ この選択でプロンプト（OPS_DOC）は 1 行も増えない
（2026-08-24 の実測: OPS_DOC に 16 行足したら別 op の分類が 98.1%→94.2% に落ちた）。

★ 検証の形（CSV_EXPORT と同じ主張・別の読み戻し）:
  出した PDF の**テキスト層を読み戻して**、元シートのセル値と突き合わせる。
  2026-08-24 の実測でテキスト層の抽出は活字 12/12 厳密・0.01 秒（OCR ではない・
  確率でない）。だから「PDF にした」ではなく「PDF に元の値が入っている」を言える。

★ 読み戻しの道具（pdfplumber）は**任意の依存**にしてある。
  居なければ PDF は作るが **✓ を名乗らない**（⚠ で「機械保証はありません」と言う）。
  ── 「居るから見えない」を避けるため、試験は**居ない側を既定**で回す。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


def readback_available() -> bool:
    """テキスト層の読み戻しができるか（任意依存 pdfplumber の在否）。"""
    try:
        import pdfplumber  # noqa: F401
        return True
    except Exception:
        return False


def read_pdf_text(path) -> str:
    """PDF のテキスト層を全ページ連結して返す。読めなければ空文字。"""
    try:
        import pdfplumber
    except Exception:
        return ""
    try:
        with pdfplumber.open(str(path)) as pdf:
            return chr(10).join(page.extract_text() or "" for page in pdf.pages)
    except Exception:
        return ""


def _renderable(value) -> str:
    """セルの値を、PDF の上に出るはずの文字列にする。

    ★ 正直な限界: 表示は**セルの書式**で決まる（1000 → 「1,000」、日付 → 「2026/07/31」）。
      ここでは書式を再現しない。だから照合は「見つかった/見つからない」の弱い主張に留め、
      見つからなかった値は**嘘をつかずに数える**（欠落として名指しする）。
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


@dataclass
class PdfCheck:
    checked: int = 0
    missing: list = field(default_factory=list)   # PDF の中に見つからなかった値
    # ★ 2026-08-24: 「見なかったセル」を分母から**黙って消さない**ための枠。
    #   数式セル（キャッシュ値なし）は data_only=True で None になり、照合の対象から
    #   静かに落ちていた ── 金額列が全部数式の見積書で「✓ 3 個の値が載っている（欠落 0）」
    #   が出る（金額を 1 個も見ていない）。分母は開示しなければ主張にならない。
    uncheckable: int = 0
    text_chars: int = 0
    available: bool = True


def verify_values_in_pdf(pdf_path, values) -> PdfCheck:
    """元シートの値が、出した PDF のテキスト層に載っているかを数える。

    ★ 空白の扱い: PDF の抽出は語間に空白を挟むことがあるので、比較は**空白を除いた**
      文字列同士で行う（LibreOffice の描画の都合で ✓ が落ちるのを避けるため）。
    """
    r = PdfCheck()
    if not readback_available():
        r.available = False
        return r
    text = read_pdf_text(pdf_path)
    r.text_chars = len(text)
    flat = text.replace(" ", "").replace(chr(160), "")
    for v in values:
        rendered = _renderable(v)
        if rendered == "":
            continue
        r.checked += 1
        if rendered.replace(" ", "") not in flat:
            r.missing.append(rendered)
    return r
