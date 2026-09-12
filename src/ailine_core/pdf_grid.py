"""pdf_grid — PDF の 1 ページを `Grid`（格子）にする（2026-09-12・設計 D1〜D3・D5）。

★★ 帳票を読む器官（`form_read`）は Excel を知らない。知っているのは `Grid` の口だけ。
  だから PDF から同じ格子を作れば、規則を 1 行も変えずに PDF の請求書を読める
  ── 設計 docs/DESIGN-20260912-PDFの帳票を読む.md。ここは**格子を作るだけ**で、意味は持たない。

★ ここでやることは 3 つ（増やさない・増やしたくなったら列挙で当てにいく合図）:
  ① テキスト層を語として取り出す（pdfplumber・**OCR ではない**）
  ② ★ 部首だけを正規化する ── 実物の日本語 PDF は `請求⽇` の「⽇」が U+2F47（康熙部首）で、
     見た目は同じでも `請求日`(U+65E5) と別の文字。正規化しないとラベル照合は永久に当たらない
     （Adobe 16 枚・Wondershare 1 枚すべてで実測）。★ NFKC を**文字列全体に**かけない ──
     全角の数字（`１，３２０，０００`）まで半角に化けて、「金額の欄に文字が入っている」という
     Excel 側と同じ拒否ができなくなる。全角の数字は**文字だった証拠**なので残す。
  ③ 行は y でまとめ、列は中心の並びをいちばん大きな隙間で切る
     （実測: 列内のばらつき ≤31pt < 列間の隙間 ≥57pt の谷）

★★ 空白は「ここまでが同じセル」の証拠（実測で 31% → 83.6%）:
  実物の帳票は「合　　計」のように**全角スペースで字間を空ける**（均等割付）。
  pdfplumber の既定は空白文字を捨ててそこで語を割るので、`合` と `計` が別のマスになり、
  帯の合計が永久に見つからない。`keep_blank_chars=True` で空白を残し、`norm` に落とさせる。
  位置の隙間（空白文字が無い所）は別のマス ── 閾値を発明せずに両方を得る。

★ Excel と違って**式が無い** ── 写し合いを見分ける手段が無いので、この格子から読んだ請求額は
  「裏が取れた」を名乗らない（設計 D4・`form_read.copies_are_indistinguishable`）。
"""
from __future__ import annotations

import re
import unicodedata

from ailine_core.form_grid import Cell, Grid

#: 同じ行と見なす y の差（pt）。本文 10〜11pt の半分弱。
ROW_TOL = 4.0
#: 列を切る隙間（pt）。★ 実測の谷（列内 ≤31 < 列間 ≥57）から。
COL_GAP = 40.0

#: 康熙部首（U+2F00–2FD5）と CJK 部首補助（U+2E80–2EF3）── ここ**だけ**を正規化する。
_RADICALS = re.compile(r"[⺀-⻳⼀-⿕]")

#: 金額として描かれた文字（Excel なら「数値 ＋ 表示形式」だったもの）。
#:   ★ 変換器の仕事は**描画の逆算**であって推測ではない ── `330,000円` → 330000 ＋ 書式 "円"。
#:   郵便番号 000-0000 / 電話 00-0000-0000 / 伝票 S-202504-001 / 1個 / 全角の数字 は当たらない。
#:   ★ 末尾のダッシュ（¥32,400-）は実物の書き方 ── Excel 側の `_MONEYISH` も受ける約束。
_DRAWN_NUMBER = re.compile(r"^([¥￥]?)([0-9][0-9,]*(?:\.[0-9]+)?)(円)?([-ー－])?$")


class NoTextLayer(ValueError):
    """テキスト層が無い PDF（スキャンした紙など）。★ 「請求書ではなかった」とは別の理由。"""


def normalize_radicals(text: str) -> str:
    """部首として符号化された漢字だけを本来の漢字に直す（`⽇` U+2F47 → `日` U+65E5）。"""
    return _RADICALS.sub(lambda m: unicodedata.normalize("NFKC", m.group(0)), text)


def recover(text: str) -> tuple:
    """描かれた文字 → (値, 表示形式)。数でなければ文字のまま返す。

    ★ 裸の「1」も数に戻す（2026-09-12 の実測）: 初版は「番号や年を数として壊す」ことを恐れて
      文字のまま置いたが、そのせいで**数量の列が空に見え**、器官が「単価は在るのに数量が空」と
      読んで合計を信じなくなった（Wondershare の雛形）。恐れていた壊し方は測ったら起きなかった
      ── 請求番号はラベルから文字として読み、日付は日付として読むので、裸の数が化けて入る経路が無い。
    """
    m = _DRAWN_NUMBER.match(text)
    if not m:
        return text, ""
    body = m.group(2).replace(",", "")
    v = float(body) if "." in body else int(body)
    fmt = ((m.group(1) or "") + ('#,##0"円"' if m.group(3) else "#,##0") + (m.group(4) or ""))
    return v, fmt


def _cluster(values: list, gap: float) -> list:
    """1 次元の値を、gap より広い隙間で切って束にする。戻り値は各束の代表値（昇順）。"""
    out: list = []
    for v in sorted(values):
        if out and v - out[-1][-1] <= gap:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(b) / len(b) for b in out]


def grid_from_words(words, *, page_no: int = 1,
                    row_tol: float = ROW_TOL, col_gap: float = COL_GAP) -> Grid:
    """語の並び → 格子。words: [{"text", "x0", "x1", "top"}, …]（pdfplumber の形）。

    ★ 番地は `2頁12行3列` ── 人が原本を開いて指で追える形（設計 D5）。
    """
    items = []
    for w in words:
        t = normalize_radicals(str(w["text"])).strip()
        if t:
            items.append((t, float(w["x0"]), float(w["x1"]), float(w["top"])))
    if not items:
        return Grid({}, 0, 0, ())

    tops = _cluster([it[3] for it in items], row_tol)
    centers = _cluster([(it[1] + it[2]) / 2 for it in items], col_gap)

    def nearest(v, xs):
        return min(range(len(xs)), key=lambda i: abs(xs[i] - v))

    cells: dict = {}
    for t, x0, x1, top in items:
        r = nearest(top, tops) + 1
        c = nearest((x0 + x1) / 2, centers) + 1
        while (r, c) in cells:          # 同じマスに 2 語 → 右へ逃がす（つないでも意味が壊れる）
            c += 1
        v, fmt = recover(t)
        cells[(r, c)] = Cell(row=r, col=c, value=v, at=f"{page_no}頁{r}行{c}列",
                             anchor=(r, c), from_merge=False, fmt=fmt)
    return Grid(cells, len(tops), max(c for _r, c in cells), ())


def grid_from_page(page, *, page_no: int = 1) -> Grid:
    """pdfplumber の 1 ページ → 格子。★ 空白文字は残す（同じセルの証拠）。"""
    return grid_from_words(page.extract_words(keep_blank_chars=True), page_no=page_no)


def grids_of(path) -> list:
    """PDF の全ページを格子にする。戻り値 [(ページ番号, Grid), …]。

    ★ 1 語も無い（全ページが空）なら `NoTextLayer` ── 黙って 0 件にしない（設計 D7）。
    """
    import pdfplumber                    # ★ 必須依存だが import はこのモジュールに閉じる（D1）

    out = []
    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            out.append((i, grid_from_page(page, page_no=i)))
    if not any(g.all_cells() for _i, g in out):
        raise NoTextLayer(
            "テキスト層がありません（スキャンした紙の PDF は文字として読めません ── "
            "OCR はしていません）")
    return out
