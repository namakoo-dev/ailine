# -*- coding: utf-8 -*-
"""3 体の architect が揃って推した形を、答え付きの検体 B で採点する（2026-09-10 夜）。

★ 測るのは「これから作るもの」の見込み。scan の 14/40 は**置き換える側**の数字。

規則（3 体の共通部分だけを実装。俺の思いつきは足さない）:
    請求元   … 「御中」ブロックとは別に、**冊をまたいで値が散る**文字セル
                （Namakoo Q1: 受け取った請求書 → 宛先は全冊同じ・請求元だけ散る）
    請求額   … ラベル（ご請求金額/合計/御請求金額…）の右 5 列以内の数
                ＋ 明細金額列に揃えたサマリ帯 の 2 口以上が一致したら「確」

★ 採点は検体側の答え（値と番地）で行う。★ 3 区分で出す（確 / 単 / 割 / 無）。
"""
import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

SP = Path(__file__).resolve().parent   # ★ 雛形と検体はこの隣（追跡しない）
FOLDER = SP / "received40"
ANS = {a["file"]: a for a in json.loads((SP / "答え_received40.json").read_text(encoding="utf-8"))}

HEAD_MONEY = ("ご請求金額", "御請求金額", "今回請求額", "請求金額", "合計金額")


def norm(s):
    return re.sub(r"[\s　]", "", str(s)).strip()


def merged(ws):
    m = {}
    for rng in ws.merged_cells.ranges:
        v = ws.cell(row=rng.min_row, column=rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                m[(r, c)] = v
    return m


def val(ws, mm, r, c):
    v = ws.cell(row=r, column=c).value
    return mm.get((r, c)) if v in (None, "") else v


# ── パス① 全冊を読む（証拠を並べるだけ・決めない） ──────────────
books = []
for p in sorted(FOLDER.glob("*.xlsx")):
    a = ANS[p.name]
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb[a["シート"]]
    mm = merged(ws)
    R, C = min(ws.max_row, 60), min(ws.max_column, 20)

    texts = {}          # 番地 → 文字値（御中の行より上のブロック）
    money = []          # (番地, 値, ラベル)
    ochu = None
    for r in range(1, R + 1):
        for c in range(1, C + 1):
            v = val(ws, mm, r, c)
            if v in (None, ""):
                continue
            s = norm(v)
            if isinstance(v, str):
                if "御中" in s and ochu is None:
                    ochu = (r, c)
                if 2 <= len(s) <= 30 and not s.startswith(("〒", "TEL", "FAX", "E-Mail")):
                    texts[(r, c)] = s
                if any(s.startswith(k) for k in HEAD_MONEY):
                    for cc in range(c + 1, min(C, c + 6) + 1):
                        nv = val(ws, mm, r, cc)
                        if isinstance(nv, (int, float)) and nv != 0:
                            money.append((f"{get_column_letter(cc)}{r}", nv, s[:10]))
                            break
    books.append({"file": p.name, "texts": texts, "money": money, "ochu": ochu, "ans": a})
    wb.close()

# ── パス② 冊をまたいで散る文字セルを探す（＝請求元） ──────────────
by_pos = defaultdict(list)
for b in books:
    for pos, s in b["texts"].items():
        by_pos[pos].append(s)

spread = {}
for pos, vals in by_pos.items():
    if len(vals) >= 3:                       # 3 冊以上に同じ番地がある
        spread[pos] = len(set(vals)) / len(vals)

# ── 採点 ──────────────────────────────────────────────
issuer_ok = issuer_ng = issuer_none = 0
grades = {"確": 0, "単": 0, "割": 0, "無": 0}
money_ok = 0
misses = []
for b in books:
    a = b["ans"]
    # 請求元: 「散る」番地のうち、御中セルでなく、最も散っているもの
    cand = [(pos, spread[pos]) for pos in b["texts"] if pos in spread and pos != b["ochu"]]
    got_issuer = None
    if cand:
        pos = max(cand, key=lambda kv: kv[1])[0]
        got_issuer = b["texts"][pos]
    if got_issuer is None:
        issuer_none += 1
    elif got_issuer == norm(a["請求元"]):
        issuer_ok += 1
    else:
        issuer_ng += 1
        if len(misses) < 5:
            misses.append((b["file"], got_issuer, a["請求元"]))

    # 請求額: 2 口以上が一致したら「確」
    vals = {round(v, 2) for _, v, _ in b["money"]}
    if len(b["money"]) >= 2 and len(vals) == 1:
        g = "確"
    elif len(b["money"]) >= 2:
        g = "割"
    elif len(b["money"]) == 1:
        g = "単"
    else:
        g = "無"
    grades[g] += 1
    if b["money"] and abs(b["money"][0][1] - a["請求額(税込)"]) < 0.5:
        money_ok += 1

n = len(books)
print(f"=== 検体 B（実物の骨・答えつき 40 冊）で採点 ===\n")
print(f"請求元が当たった      {issuer_ok}/{n}   外した {issuer_ng}   取れず {issuer_none}")
print(f"請求額が答えと一致    {money_ok}/{n}")
print(f"請求額の区分          確 {grades['確']} / 単 {grades['単']} / "
      f"★割 {grades['割']} / 無 {grades['無']}")
if misses:
    print("\n★ 請求元を外した分:")
    for f, got, want in misses:
        print(f"   {f}: 得た『{got}』 答え『{want}』")
