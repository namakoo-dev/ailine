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
        # ★ 2026-08-26（出入口の盲検・高5）: PDF は Excel の表示に従って TRUE/FALSE と出る。
        #   `str(True)` だけを候補にしていたので、真偽値のある表は**どうやっても ✓ が出ず
        #   常に exit 3** だった。csv_export は既に "TRUE" に寄せていて、兄弟モジュールで
        #   表記が食い違っていた ── 番人が誤爆すると、人は番人を見なくなる。
        return ["TRUE" if value else "FALSE", str(value)]
    if isinstance(value, _dt.time):
        # ★ 同（高5）: 時刻は `09:00:00` でなく `9:00` と出る。勤怠表が全滅していた。
        return [f"{value.hour}:{value.minute:02d}",
                f"{value.hour:02d}:{value.minute:02d}",
                value.isoformat()]
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
    # ★ 2026-08-26（出入口の盲検・高6）: 空白しか除いていなかったので、列幅で
    #   `1月` が改行をまたいで折り返されると不一致になっていた。改行も除く。
    # ★ 2 つの見方を持つ（片方だけでは両立しない・自分で踏んだ）:
    #   flat   … 空白を全部除いた形。語間に空白が挟まる抽出（`1, 000`）に強い
    #   spaced … 空白を 1 つに畳んだ形。**数字の境界**を見るのに要る
    #            （flat だと `5 80` が `580` になり、`5` が「数字に挟まれている」と
    #             判定されて正しい一致まで捨ててしまう）
    import re as _re2
    _norm = text.replace(chr(160), " ")
    flat = _re2.sub(r"\s+", "", _norm)
    spaced = _re2.sub(r"\s+", " ", _norm)
    # ★★ 2026-08-26（出入口の盲検・致命1）: ここは `c in flat` ── **PDF 全文への
    #   素の部分文字列の包含**で、位置も出現回数も見ていなかった。
    #   ・隠し行/印刷範囲で消えた行の値が、別の行にも在れば「載っている」と数えた
    #     （6 行中 3 行が PDF に無いのに ✓ 欠落 0・実測）
    #   ・`12345` は可視の `123456` に含まれるので、消えても検出されなかった
    #   ★ 直し: **必要な回数だけ在ること**を見る（同じ値が 2 回在るなら PDF にも 2 回）。
    #     数字だけの候補は**前後が数字でない**ことも要る（部分一致の穴）。
    need: dict = {}
    order: list = []
    for v in values:
        cands = tuple(c for c in renderings(v) if c != "")
        if not cands:
            continue
        r.checked += 1
        if cands not in need:
            need[cands] = 0
            order.append(cands)
        need[cands] += 1
    for cands in order:
        found = 0
        for c in cands:
            hay = spaced if (c and all(ch in _DIGITS for ch in c)) else flat
            found = max(found, _count_occurrences(hay, c))
        if found < need[cands]:
            r.missing.append(cands[0])
    return r


_DIGITS = "0123456789"


def _count_occurrences(flat: str, cand: str) -> int:
    """flat に cand が「独立して」現れる回数。

    ★ 数字だけの候補は、前後に数字が続いていたら数えない ── `12345` を `123456` の
      中に見つけて「載っている」と言わないため（実測の穴）。
    """
    numeric = cand and all(ch in _DIGITS for ch in cand)
    n = 0
    start = 0
    while True:
        i = flat.find(cand, start)
        if i < 0:
            return n
        start = i + 1
        if numeric:
            before = flat[i - 1] if i > 0 else ""
            after = flat[i + len(cand)] if i + len(cand) < len(flat) else ""
            # ★ 空文字は「任意の文字列の部分文字列」なので `"" in _DIGITS` は True。
            #   先頭・末尾の一致が全部ここで捨てられる（自分で仕込みかけた・実測で捕獲）。
            if (before and before in _DIGITS) or (after and after in _DIGITS):
                continue
        n += 1


# --- openpyxl の往復で落ちる図形の検出（2026-08-24・第三波 #9）--------------------
#
# ★ 実測（この repo の中で測った）: openpyxl の load→save は
#     画像（xdr:pic / xl/media）      → **残る**
#     画像でない図形（xdr:sp・テキストボックス・オートシェイプ） → **消える**
#   `ailine export-pdf` は指定シートの抽出とページ設定のために原本を openpyxl で
#   一度書き直すので、**描かれた角印・社判が PDF から消える**。しかも出来上がった
#   PDF は完成品に見える ── 消えたものは差分に出ない、の最悪の形。
#   ここでは判定も変換も変えない。**消えるものを、消える前に名指しする**。

import re as _re
import zipfile as _zipfile

_SHAPE_RE = _re.compile(r"<(?:\w+:)?sp[ >]")
_PIC_RE = _re.compile(r"<(?:\w+:)?pic[ >]")
_NAME_RE = _re.compile(r'<(?:\w+:)?cNvPr[^>]*\bname="([^"]*)"')


def pillow_available() -> bool:
    """Pillow が入っているか。

    ★ 2026-08-24 の実測（CI と同じ素の環境で走らせて発覚）: openpyxl は画像を
    読み書きするのに Pillow を使う。Pillow が無いと、往復で **画像まで落ちる**。
    ailine が宣言している依存は openpyxl だけなので、買い手の環境に Pillow が
    在る保証はない ── つまり「写真として貼られた印は消えません」は条件つきの真実で、
    無条件に言うと嘘になる。
    """
    import importlib.util
    try:
        return importlib.util.find_spec("PIL") is not None
    except (ImportError, ValueError, AttributeError):
        # find_spec は「壊れた/遮断された」パッケージで例外を投げうる。
        # 判定できない＝当てにできない、なので「無い」側に倒す（安全側＝嘘の安心を出さない）。
        return False


def vanishing_shapes(path, with_pillow: bool | None = None) -> list:
    """openpyxl の往復で落ちる図形の名前を返す（画像は落ちないので数えない）。

    戻り値は名前のリスト（名前が無ければ "(名前なし)"）。読めない/図形が無ければ空。
    ★ 判定は zip の中の drawing XML を直接読む ── openpyxl に聞くと、
    openpyxl が読めない図形は最初から見えないので恒真になる（別実装で測る）。
    """
    names = []
    if with_pillow is None:
        with_pillow = pillow_available()
    try:
        with _zipfile.ZipFile(path) as z:
            for entry in z.namelist():
                if not (entry.startswith("xl/drawings/") and entry.endswith(".xml")):
                    continue
                text = z.read(entry).decode("utf-8", errors="replace")
                has_shape = bool(_SHAPE_RE.search(text))
                # ★ Pillow が無ければ画像（pic）も落ちる ── 条件つきで数える。
                has_pic = (not with_pillow) and bool(_PIC_RE.search(text))
                if not (has_shape or has_pic):
                    continue
                # 図形が在る drawing の中から、sp を持つアンカーの名前だけを拾う
                for anchor in _re.split(r"(?=<(?:\w+:)?(?:one|two)CellAnchor)", text):
                    is_shape = bool(_SHAPE_RE.search(anchor))
                    is_pic = (not with_pillow) and bool(_PIC_RE.search(anchor))
                    if not (is_shape or is_pic):
                        continue
                    m = _NAME_RE.search(anchor)
                    names.append(m.group(1) if m and m.group(1) else "(名前なし)")
    except (_zipfile.BadZipFile, KeyError, OSError):
        return []
    return names


def vanishing_shapes_warning(names, with_pillow: bool | None = None) -> list:
    """人へ見せる行（空なら空リスト）。

    ★「写真として貼られた印は消えません」は **Pillow が在る時だけ**の真実。
    無条件に言うと嘘になるので、環境で文言を変える（機械は嘘をつかない）。
    """
    if not names:
        return []
    if with_pillow is None:
        with_pillow = pillow_available()
    shown = "、".join(f"『{n}』" for n in names[:3])
    more = f"（ほか {len(names) - 3} 件）" if len(names) > 3 else ""
    lines = [f"⚠ この冊には、PDF に写せない図形が {len(names)} 件あります: {shown}{more}"]
    if with_pillow:
        lines.append("  → 角印・社判・テキストボックスなどが PDF から消えます"
                     "（写真として貼られた印は消えません）。")
    else:
        lines.append("  → 角印・社判・テキストボックスに加え、写真として貼られた印も"
                     "PDF から消えます（この環境には Pillow が入っていません）。")
        lines.append("  → pip install Pillow で、写真の印だけは残せます。")
    lines.append("  → 印が要る書類なら、LibreOffice や Excel で直接 PDF 保存して"
                 "ください。")
    return lines
