# -*- coding: utf-8 -*-
"""否定の読みが**言い方によって反転しないか**を測る（2026-09-06）。

★★ なぜ測るか: 09-05 の実走行に「所属が営業**以外**の行のメモに『○』を付けて」を
  `比べ方:等しい` と読んで書いた回が残っている（本物の誤り・事後条件は pass）。
  その後 `task_says_except` が入り、依頼文が否定なら**機械が LLM に勝つ**ようにした。
  ★ だから「以外」の反転はもう構造的に起きない。**残っているのは別の 2 つ**:

    Q1 語彙の穴 ── 見ているのは 3 語だけ（"以外" / "を除いた" / "を抜いた"）。
       「営業でない」「営業を除く」（★ 過去形しか無い）「営業じゃない」は
       False を返す → 機械は勝たず、LLM の `eq` が通る → **逆の行に書いて ✓**。

    Q2 結び先の穴 ── 判定が**文全体**を見ている（`any(w in task)`）。
       「メモ以外は変えずに、所属が営業の行のメモに○を付けて」の「以外」は
       条件の話ではないのに `nin` が強制される → **営業でない行に書く**。
       ★ Q1 より重い（実害が「書く行が丸ごと入れ替わる」）。

★ 事前に置いた線（測る前に凍結）:
    Q2 が 1 件でも再現 → 直す
    Q1 は**再現した言い方だけ**を対象にする（再現しない言い方は触らない）
    否定でない依頼で `nin` が出たら、その直し方は採らない

★ 製品は 1 バイトも変えない。★★ **本物の CLI を `--dry` で走らせる**（模型を書かない）:
  初版は translate_task → verify_dsl_args を自前で呼んだが、**全件 OUT_OF_VOCAB** になった
  ── 動くと分かっている依頼まで。本番には**読み直しの層**があり、模型はそれを通さない。
  今日 3179 件の測定で得た処方（模型のずれが測定値になる）を、その日のうちに再演した。
  `--dry` なら LibreOffice も要らず、**解釈行に「比べ方」が出る**ので読み取れる。

使い方:
    python bench/negation_reading_probe.py
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import openpyxl  # noqa: E402

import ailine  # noqa: E402

#: 名簿（09-05 の事故と同じ形の表）
#: ★ **非対称**にする（営業 3 行 / 以外 1 行）── 対称だと eq と nin で当てはまる行の
#:   数が同じになり、「宣言は直ったが書く行は直っていない」を見逃す。
ROWS = [("田中", "営業", "主任", ""), ("鈴木", "経理", "副任", ""),
        ("佐藤", "営業", "副任", ""), ("山田", "営業", "主任", "")]

#: (分類, 依頼文, 期待する比べ方)
CASES = [
    # --- 効いていると分かっている形（陰性対照）---
    ("既知OK", "所属が営業以外の行のメモに「○」を付けて", "nin"),
    ("既知OK", "所属が営業を除いた行のメモに「○」を付けて", "nin"),
    ("既知OK", "所属が営業の行のメモに「○」を付けて", "eq"),
    # --- Q1: 語彙の穴の疑い ---
    ("Q1", "所属が営業でない行のメモに「○」を付けて", "nin"),
    ("Q1", "所属が営業ではない行のメモに「○」を付けて", "nin"),
    ("Q1", "所属が営業じゃない行のメモに「○」を付けて", "nin"),
    ("Q1", "所属が営業を除く行のメモに「○」を付けて", "nin"),
    ("Q1", "所属が営業を除いて、メモに「○」を付けて", "nin"),
    # --- Q2: 文中の無関係な否定語 ---
    ("Q2", "メモ以外は変えずに、所属が営業の行のメモに「○」を付けて", "eq"),
    ("Q2", "氏名以外の列はそのままで、所属が営業の行のメモに「○」を付けて", "eq"),
    ("Q2", "担当を除いた列は触らずに、所属が営業の行のメモに「○」を付けて", "eq"),
    # --- 抽出（★ 別の呼び出し口 ── 条件つき書換とは違う場所で否定を読む）---
    ("抽出", "所属が営業以外の行を抜き出して", "nin"),
    ("抽出", "所属が営業でない行を抜き出して", "nin"),
    ("抽出", "メモ以外は変えずに、所属が営業の行を抜き出して", "eq"),
]


def build(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名簿"
    ws.append(["氏名", "所属", "担当", "メモ"])
    for r in ROWS:
        ws.append(r)
    wb.save(path)
    return path


# ★ 解釈行のラベルは op で違う（条件つき書換=「比べ方」／抽出=「条件」）。
#   ★ 2026-09-06 の実測: 「比べ方」だけを探していたため、抽出の 3 件が
#     「出ない」に落ち、**見逃しを『期待と違った: 0 件』と報告していた**
#     ── 測定器が偽の緑を出す形。両方を探す。
_CMP = re.compile(r"(?:比べ方|条件):(\S+)")
_ROWS = re.compile(r"当てはまる行:([^\s]+(?: [^\s]+)?)")
#: 期待する行（★ 宣言だけでなく**書く行**が入れ替わったことを見る）
WANT_ROWS = {"eq": "3", "nin": "1"}
#: 解釈行の日本語ラベル → 機械の語彙（★ 表示から読むのは、それが人に見える宣言だから）
_LABEL = {"等しい": "eq", "のどれでもない": "nin", "含む": "contains"}


def run_real(book: Path, task: str) -> tuple:
    """本物の CLI を --dry で走らせ、(比べ方, 当てはまる行, 全出力) を返す。"""
    got = subprocess.run(
        [sys.executable, "-m", "ailine", "run", str(book), task, "--dry"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(ROOT), env={**os.environ, "PYTHONPATH": str(ROOT / "src")})
    out = (got.stdout or "") + (got.stderr or "")
    m, r = _CMP.search(out), _ROWS.search(out)
    return (_LABEL.get(m.group(1), m.group(1)) if m else None,
            r.group(1) if r else None, out)


def main() -> int:
    book = build(Path(tempfile.mkdtemp()) / "meibo.xlsx")
    bad = []
    print(f"表: {[r[1] for r in ROWS]}（営業 = 2・4 行目 / 以外 = 3・5 行目）")
    print()
    for kind, task, want in CASES:
        cmp_, rows, out = run_real(book, task)
        if cmp_ is None:
            bad.append((kind, task, want, "読めない", rows))
            head = next((ln for ln in out.splitlines() if ln.startswith(("解釈:", "？", "×", "⚠"))), "")
            print(f"  [{kind}] 比べ方が出ない ← {task[:34]}")
            print(f"          {head[:100]}")
            continue
        # ★ 宣言（比べ方）だけでなく、**実際に当たる行数**も見る（検体は非対称）。
        want_rows = WANT_ROWS.get(want)
        rows_ok = (rows is None) or (want_rows is None) or rows.startswith(want_rows)
        hit = (cmp_ == want) and rows_ok
        if not hit:
            bad.append((kind, task, want, cmp_, rows))
        why = "" if hit else ("★反転" if cmp_ != want else "★行が違う")
        print(f"  [{kind}] {why or '○'} 期待={want}/{want_rows}行 "
              f"実際={cmp_}/{rows}")
        print(f"          ← {task}")
    print()
    print(f"期待と違った: {len(bad)} 件")
    for kind, task, want, got, rows in bad:
        print(f"    [{kind}] {want}→{got} 行={rows} ← {task}")
    print("★ Q2 が 1 件でも出たら直す（依頼と逆の行に書く）。")
    print("★ Q1 は再現した言い方だけを対象にする（再現しない言い方は触らない）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
