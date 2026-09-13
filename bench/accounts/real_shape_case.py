# -*- coding: utf-8 -*-
"""実務の密度で測る検体 ── Namakoo 提供の請求書 3 通から起こした仕訳（2026-09-13）。

    python bench/accounts/real_shape_case.py            # 本物の CLI に掛けて内訳を出す
    python bench/accounts/real_shape_case.py --keep <dir>  # 作った冊を残して目で見る

## なぜ在るか（★ 合成検体では永久に見つからない欠陥が、これで 1 つ落ちた）

合成検体（`specimens_accounts.py`）は過去の行が**疎**で、1 行に鍵が 1 つしか乗らない。
実務の仕訳は逆で、**1 行に 取引先・摘要・補助科目・部門が同時に乗る**。この密度で測ったら:

    到達  合成 75%  →  実務の密度 30%

落ち方が想定と違った ── 鍵が当たらないのではなく、**広い鍵（取引先）の割れが、完全に一致して
いる狭い鍵（摘要）を道連れにしていた**。タクシー 3 行はすべて摘要が過去 1 行と完全一致して
正しい科目を指していたのに空欄だった。→ 設計 §9 で規則を直した（割れた鍵は拒否権でなく沈黙）。
★ 規則を直しても**合成 155 行の点数は 155/155 のまま動かない**（疎だから同居が起きない）。

## 検体の出所と、触った所

Namakoo が用意した実物の形式の請求書 3 通（オフィス用品／タクシー／通信）の明細から、
仕訳の 1 行ずつを起こした。★ **形は 1 文字も変えていない** ── 半角カナ・全角数字・全角空白・
注文番号・括弧の中の回線番号・部門情報の密集はそのまま。
★ 名前だけ差し替えた: 機密語の番人（`secretscan`）が実名の姓 2 箇所・実名の地名 4 箇所を
検出したため（公開 repo に実名を置かない）。元の PDF は repo に入れていない。

## これは採点器ではない

答え（人が付けるべき科目）は持たない ── 実務の正解は会社ごとの流儀で決まるので、
合成で宣言できない。ここで見るのは**区分の内訳と根拠の文面**で、規則を変えた日に
「実務の密度で何が起きるか」を測り直すための器。退行の番人は `tests/test_accounts_core.py` 側。
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent

MF = ["取引No", "取引日", "借方勘定科目", "借方補助科目", "借方部門", "借方取引先",
      "借方金額(円)", "貸方勘定科目", "貸方取引先", "摘要"]


def _row(no, date, account, sub, dept, partner, amount, memo, credit="未払金"):
    return [no, date, account, sub, dept, partner, amount, credit, "", memo]


#: 過去（2〜3 月・借方勘定科目が埋まっている）★ 実務の密度: 取引先も摘要も補助科目も入る。
PAST = [
    _row("1", "2026/02/28", "通信費", "固定電話", "管理部", "コムネット通信株式会社", "5000",
         "1月分 電話代（回線 0595-XX-XXXX）"),
    _row("2", "2026/02/28", "通信費", "固定電話", "管理部", "コムネット通信株式会社", "5000",
         "電話料金　2月分"),
    _row("3", "2026/03/31", "通信費", "回線", "開発部", "コムネット通信株式会社", "8500",
         "ｲﾝﾀｰﾈｯﾄ接続料(ナギ工房ﾌﾟﾛｼﾞｪｸﾄ部門)"),
    _row("4", "2026/03/31", "通信費", "固定電話", "管理部", "コムネット通信株式会社", "5000",
         "2月分 電話代（回線 0595-XX-XXXX）"),
    # ★ 同じ支払先で科目が割れる実例（実務でいちばん多い形）
    _row("5", "2026/02/10", "旅費交通費", "", "開発部", "桜井第一交通株式会社", "4500",
         "ﾀｸｼｰ乗車（桜井〜亀山） [出張]"),
    _row("6", "2026/02/15", "接待交際費", "", "営業部", "桜井第一交通株式会社", "8200",
         "タクシー料金 顧客A社様送迎含む"),
    _row("7", "2026/03/22", "旅費交通費", "", "開発部", "桜井第一交通株式会社", "1500",
         "桜井市内移動"),
    # ★ 書籍とオフィス用品が同じ支払先で混ざる
    _row("8", "2026/02/05", "消耗品費", "事務用品", "管理部", "株式会社オフィス・サプライ",
         "12000", "[注文#X880-2] ワイヤレスマウス MX-3"),
    _row("9", "2026/02/05", "新聞図書費", "書籍", "開発部", "株式会社オフィス・サプライ",
         "3500", "Python実践データ分析 (書籍)"),
    _row("10", "2026/03/12", "消耗品費", "事務用品", "管理部", "株式会社オフィス・サプライ",
         "4200", "Ａ４コピー用紙（５００枚×５）"),
]

#: 今回（4 月・請求書 3 通から起こした仕訳・借方勘定科目は空）。
TODAY = [
    _row("11", "2026/04/05", "", "事務用品", "管理部", "株式会社オフィス・サプライ", "12000",
         "[注文#X992-1] ワイヤレスマウス MX-3"),
    _row("12", "2026/04/05", "", "書籍", "開発部", "株式会社オフィス・サプライ", "3500",
         "Python実践データ分析 (書籍)"),
    _row("13", "2026/04/12", "", "事務用品", "管理部", "株式会社オフィス・サプライ", "4200",
         "Ａ４コピー用紙（５００枚×５）"),
    _row("14", "2026/04/25", "", "事務用品", "管理部", "株式会社オフィス・サプライ", "25000",
         "[注文#Y110-3] ｵﾌｨｽﾁｪｱ"),
    _row("15", "2026/04/10", "", "", "開発部", "桜井第一交通株式会社", "4500",
         "ﾀｸｼｰ乗車（桜井〜亀山） [出張]"),
    _row("16", "2026/04/15", "", "", "営業部", "桜井第一交通株式会社", "8200",
         "タクシー料金 顧客A社様送迎含む"),
    _row("17", "2026/04/22", "", "", "開発部", "桜井第一交通株式会社", "1500",
         "桜井市内移動"),
    _row("18", "2026/04/01", "", "固定電話", "管理部", "コムネット通信株式会社", "5000",
         "3月分 電話代（回線 0595-XX-XXXX）"),
    _row("19", "2026/04/01", "", "固定電話", "管理部", "コムネット通信株式会社", "5000",
         "電話料金　4月分"),
    _row("20", "2026/04/01", "", "回線", "開発部", "コムネット通信株式会社", "8500",
         "ｲﾝﾀｰﾈｯﾄ接続料(ナギ工房ﾌﾟﾛｼﾞｪｸﾄ部門)"),
]


def write_csv(path: Path, rows) -> Path:
    """★ 改行は LF・BOM 付き utf-8（`write_text` を使わない ── 改行を壊す）。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    buf = []
    writer = csv.writer(_Sink(buf), lineterminator="\n")
    writer.writerow(MF)
    for row in rows:
        writer.writerow(row)
    path.write_bytes("".join(buf).encode("utf-8-sig"))
    return path


class _Sink:
    def __init__(self, out):
        self.out = out

    def write(self, text):
        self.out.append(text)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="実務の密度で ailine accounts を測る")
    ap.add_argument("--keep", type=Path, default=None, help="作った冊を残す場所")
    a = ap.parse_args(argv)
    work = Path(tempfile.mkdtemp(prefix="accounts_real_"))
    try:
        today = write_csv(work / "元" / "MF_仕訳_2026-04.csv", TODAY)
        past = write_csv(work / "元" / "MF_過去_2026-02-03.csv", PAST)
        out = work / "候補.xlsx"
        # ★ wheel を install していない手元でも src を見せる（bench/accounts/run_accounts.py と
        #   同じ作法）── これが無いと **古い install 済みの ailine** を拾って
        #   「accounts なんてサブコマンドは無い」で落ちる（実測）。
        env = dict(os.environ)
        env["PYTHONPATH"] = str(REPO / "src") + os.pathsep + env.get("PYTHONPATH", "")
        r = subprocess.run([sys.executable, "-m", "ailine", "accounts", str(today),
                            "--past", str(past), "--out", str(out), "--json"],
                           capture_output=True, text=True, timeout=300,
                           encoding="utf-8", errors="replace", cwd=str(REPO), env=env)
        if r.returncode != 0:
            print(r.stdout or "", r.stderr or "")
            return r.returncode
        payload = json.loads(r.stdout.strip().splitlines()[-1])
        print(f"今回の行 {len(TODAY)}／候補 {len(payload['rows'])}"
              f"／過去 {len(PAST)}（実務の密度: 1 行に鍵が複数）")
        print("内訳 " + "／".join(f"{g} {n}" for g, n in payload["grades"].items()))
        for number, got in sorted(payload["rows"].items(), key=lambda kv: int(kv[0])):
            memo = TODAY[int(number) - 2][-1]
            print(f"  行{number:>3}  {got['grade']}  {str(got['account'] or '—'):8} {memo}")
        filled = sum(1 for got in payload["rows"].values() if got["account"])
        print(f"到達（値が出た）{filled}/{len(payload['rows'])}")
        print("★ 実務の正解は会社ごとの流儀で決まる ── ここで見るのは内訳と根拠の文面だけ。")
        if a.keep:
            shutil.copytree(work, a.keep, dirs_exist_ok=True)
            print(f"残した: {a.keep}")
        return 0
    finally:
        if not a.keep:
            shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
