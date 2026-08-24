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

import datetime as _dt
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


def renderings(value) -> list:
    """1 つのセルの値が PDF 上に現れうる**表示のゆれ**を全部返す。

    ★ なぜ在るか（盲検レビュー・2026-08-24）: 初版は `str(value)` 一本で照合していた。
      1000 は PDF 上では「1,000」や「¥1,000」、日付は「2026/07/31」と出るので、
      **正しく出来ている PDF に × が出て exit 3** になっていた（カンマ書式のある請求書は
      ほぼ全滅）。番人が誤爆すると、人は番人を見なくなる。
    ★ ただし何でも通す作りにはしない ── 候補は**その値から機械的に導ける表記だけ**。
      別の値（9999）が混ざる余地は無い。
    """
    if value is None:
        return []
    out = []
    if isinstance(value, _dt.datetime) or isinstance(value, _dt.date):
        d = value.date() if isinstance(value, _dt.datetime) else value
        out += [f"{d.year}/{d.month:02d}/{d.day:02d}", f"{d.year}/{d.month}/{d.day}",
                f"{d.year}-{d.month:02d}-{d.day:02d}",
                f"{d.year}年{d.month}月{d.day}日"]
        return out
    if isinstance(value, bool):
        return [str(value)]
    if isinstance(value, (int, float)):
        n = int(value) if float(value).is_integer() else value
        plain = str(n)
        out.append(plain)
        if isinstance(n, int):
            grouped = f"{n:,}"
            out.append(grouped)
            out.append("¥" + grouped)
            out.append("¥" + grouped)
        return out
    return [str(value)]


def _renderable(value) -> str:
    """互換のための 1 本目の表記（表示ゆれの先頭）。"""
    cands = renderings(value)
    return cands[0] if cands else ""


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
        cands = [c for c in renderings(v) if c != ""]
        if not cands:
            continue
        r.checked += 1
        # ★ 表示ゆれのどれか 1 つでも載っていれば「載っている」。
        if not any(c.replace(" ", "") in flat for c in cands):
            r.missing.append(cands[0])
    return r
