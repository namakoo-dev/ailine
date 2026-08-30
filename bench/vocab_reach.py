#!/usr/bin/env python
"""語彙に**届くか**を測る ── 93 件の検体が覆っていない op 系統の分類だけを見る。

★★ 2026-08-30（Namakoo「opに無い操作は出来なくて当然。③で少し検証してみたい」）:
  効果で測る検体 93 件は **6 系統**（セル・行追加・行削除・列追加・列削除・並べ替え）
  しか覆っていない。集計・ピボット・転記・グラフ・条件つき書換・書式などは
  **一度も測っていない**。そこで誤って「語彙外」に落ちていないかは、
  今の数字から何も言えなかった。

★★ 2026-08-30（Namakoo「読み直しって具体的には何をしてるの？」で判明した測定器の欠陥）:
  初版は `translate_task()` を直接呼んでいた ── これは**一段目の LLM が返した op**しか
  見ておらず、**利用者が通る道の手前で止まっていた**。
  実際の経路には、そのあとに**機械の読み直し**が在る:
      一段目 → 読み直し（依頼文と実表を見て、断る側の op を横取りする）→ 検証 → 適用
  ★ しかも EXTRACT_COLUMNS の読み直しは **OUT_OF_VOCAB を横取り対象にしている** ──
    初版が「0/3・語彙外」と報告したまさにその状態を、拾うように作ってあった。
  ★ **測定器が製品の欠陥をでっち上げていた。** 実際に走らせて測る形に直す。

★ ここでは **op の分類だけ**を見る（引数もファイルも見ない）。
  問いは 1 つ:「在る操作に、普通の言い方で届くか」。
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import ailine  # noqa: E402

MODEL = __import__("os").environ.get("AILINE_MODEL", "qwen2.5-coder:7b")


def _book(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(["商品", "分類", "売上", "原価", "担当", "日付"])
    for r in (["りんご", "果物", 1200, 700, "田中", "2026/08/01"],
               ["みかん", "果物", 800, 300, "佐藤", "2026/08/02"],
               ["にんじん", "野菜", 1500, 900, "田中", "2026/08/03"],
               ["だいこん", "野菜", 600, 200, "鈴木", "2026/08/04"]):
        ws.append(r)
    s2 = wb.create_sheet("単価表")
    s2.append(["商品", "定価"])
    for r in (["りんご", 150], ["みかん", 100], ["にんじん", 200], ["だいこん", 80]):
        s2.append(r)
    wb.save(path)


# (期待する op, [人が普通に言いそうな言い方 ×3]) ── ★ 登録簿の語をそのまま使わない
CASES = [
    ("AGGREGATE", ["分類ごとの売上を出して",
                    "果物と野菜でそれぞれいくらか知りたい",
                    "担当者別の売上をまとめて"]),
    ("PIVOT", ["分類と担当でクロス集計して",
                "ピボットテーブルにして",
                "縦に分類、横に担当で売上を出して"]),
    ("APPEND_TOTAL", ["一番下に売上の合計を出して",
                       "売上の総額を表の末尾に足して",
                       "合計行を作って"]),
    ("LOOKUP_FILL", ["単価表から定価を持ってきて",
                      "別のシートの定価をこっちに埋めて",
                      "商品名で単価表と突き合わせて定価を入れて"]),
    ("CHART", ["売上の棒グラフを作って",
                "売上を図にして",
                "商品ごとの売上をグラフにして"]),
    ("SET_WHERE", ["売上が1000以上の行の担当を「佐藤」にして",
                    "原価が500を超える行だけ担当を空にして",
                    "売上1000超のものに担当「佐藤」を入れて"]),
    ("SET_COLUMN_VALUE", ["担当を全部「佐藤」にして",
                           "担当の列を丸ごと「佐藤」で埋めて",
                           "担当を一斉に「佐藤」へ書き換えて"]),
    ("NUMBER_FORMAT", ["売上を3桁区切りにして",
                        "売上にカンマを入れて",
                        "金額が読みにくいので桁区切りにして"]),
    ("BOLD", ["見出しを太字にして",
               "1行目を太くして",
               "ヘッダーを強調して"]),
    ("FILL_COLOR", ["見出しに色を付けて",
                     "1行目を塗って",
                     "ヘッダーの背景を色付きにして"]),
    ("DRAW_BORDERS", ["表にけい線を引いて",
                       "枠線を付けて",
                       "表全体に罫線を入れて"]),
    ("AUTOFIT", ["列の幅を内容に合わせて",
                  "列が狭いので広げて",
                  "幅を自動で調整して"]),
    ("EXTRACT", ["売上が1000以上の行だけ抜き出して",
                  "野菜の行だけ別シートにして",
                  "原価が500より安いものだけ取り出して"]),
    ("EXTRACT_COLUMNS", ["商品と売上の列だけ抜き出して",
                          "必要なのは商品と売上だけなので、それだけ別シートに",
                          "商品・売上の2列だけ取り出して"]),
    ("DEDUP", ["商品の重複を消して",
                "同じ商品が2回出てくるので1つにまとめて",
                "商品名で重複を除いて"]),
    ("MERGE", ["A1とB1を結合して",
                "1行目の左2つのセルを繋げて",
                "見出しの2セルをまとめて1つにして"]),
]


def _reached(book: Path, task: str, want: str) -> tuple:
    """★ 実際に `ailine run --dry` を走らせて、**読み直しまで通った**あとの op を見る。
       --dry なので文書には触らない（適用も検算もしない）。"""
    import json as _json
    import subprocess
    env = {**__import__("os").environ, "PYTHONPATH": str(ROOT / "src")}
    r = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(book), task,
         "--dry", "--json", "--sheet", "売上"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=300, cwd=str(ROOT), env=env)
    ops = []
    for ln in (r.stdout or "").splitlines():
        ln = ln.strip()
        if not ln.startswith("{"):
            continue
        try:
            d = _json.loads(ln)
        except Exception:
            continue
        for st in (d.get("plan") or []):
            if isinstance(st, dict) and st.get("op"):
                ops.append(str(st["op"]))
        if d.get("op"):
            ops.append(str(d["op"]))
    if not ops:                      # JSON が拾えない回は本文から読む（断りなど）
        for key, label in (("OUT_OF_VOCAB", "頼める操作の一覧に照合できません"),
                            ("CLARIFY", "？")):
            if label in (r.stdout or ""):
                ops = [key]
                break
    return (want in ops), ("＋".join(ops) or "(空)")


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="ailine_vocab_") as tmp:
        p = Path(tmp) / "b.xlsx"
        _book(p)
        hit = miss = oov = 0
        bad = []
        for want, tasks in CASES:
            marks = []
            for t in tasks:
                ok, got = _reached(p, t, want)
                hit += ok
                if not ok:
                    miss += 1
                    if got in ("OUT_OF_VOCAB", "CLARIFY"):
                        oov += 1
                    bad.append((want, t, got))
                marks.append(("○" if ok else "×") + got)
            print(f"{want:18} {' '.join(marks)}")
        total = hit + miss
        print()
        print(f"合計 {total} 件: 届いた {hit}／外した {miss}"
               f"（うち『語彙外・聞き返し』{oov}）")
        print(f"  op 到達率 {hit / total * 100:.1f}%   model={MODEL}"
               "   ★ 読み直しを通した実測")
        if bad:
            print()
            print("★ 届かなかったもの（期待 / 依頼 / 返ってきた op）:")
            for w, t, g in bad:
                print(f"    {w:18} {t}  → {g}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
