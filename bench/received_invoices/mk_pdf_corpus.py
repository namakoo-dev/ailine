# -*- coding: utf-8 -*-
"""既存の検体（Excel）を LibreOffice で PDF にする（2026-09-12・設計 D10）。

    python bench/received_invoices/mk_pdf_corpus.py            # v2 87 冊 → received_pdf/v2/
    python bench/received_invoices/mk_pdf_corpus.py --all      # ＋ 入れ子・束

★ 答えは Excel のものをそのまま使う ── PDF 化しても正解は変わらない。
★ 生成物は gitignore（received_pdf/）。骨は実物ベンダーの雛形なので repo に入れない。
★ これは「ソフトが作った綺麗なテキスト層」の検体。会計ソフト発行分はこれに近いが、
  **スキャンした紙は別物**（テキスト層が無い ── 設計 D7・OCR はしない）。
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SOFFICE = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
PROFILE = HERE / "lo-profile-pdf"
OUT = HERE / "received_pdf"

GROUPS = {"v2": "received_v2", "nested": "received_nested"}


def convert(src_dir: Path, dst: Path) -> int:
    dst.mkdir(parents=True, exist_ok=True)
    files = sorted(src_dir.rglob("*.xlsx"))
    if not files:
        print(f"★ {src_dir} に xlsx が無い（先に生成器を走らせること）")
        return 0
    work = dst / "_work"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir()
    subprocess.run(
        [str(SOFFICE), "--headless", "--norestore", "--nologo",
         f"-env:UserInstallation={PROFILE.as_uri()}",
         "--convert-to", "pdf", "--outdir", str(work)] + [str(f) for f in files],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=3600)
    made = 0
    for f in files:
        pdf = work / (f.stem + ".pdf")
        if pdf.exists():
            # ★ 束はフォルダの階層を保つ（束で疑う採点はフォルダ単位）
            rel = f.relative_to(src_dir).with_suffix(".pdf")
            (dst / rel).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(pdf, dst / rel)
            made += 1
    shutil.rmtree(work)
    return made


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true", help="入れ子・束も")
    a = ap.parse_args()
    if not SOFFICE.exists():
        print(f"★ LibreOffice が無い: {SOFFICE}")
        sys.exit(2)
    todo = dict(GROUPS) if a.all else {"v2": GROUPS["v2"]}
    if a.all:
        todo["bundle"] = "received_bundle"
    for key, src in todo.items():
        n = convert(HERE / src, OUT / key)
        print(f"{key}: {n} 冊 → {OUT / key}")
