# -*- coding: utf-8 -*-
"""需要⑤（担当者別に分ける）の採点器 ── 条件は検体が来る前に凍結（split_plan.md・2026-09-12）。

道具の契約: split(book_path) -> {"parts": {名前: {"rows": [...], "amount": x}}, "blank": [...],
             "lookalike": [[a, b], ...], "excluded": [...], "refused": None | 理由,
             "proof": {"rows": {"whole", "parts", "blank", "excluded"}, "amount": {"whole", "parts"}}}

見出し（★ 正答率にしない）: ★誤配 / ★混入 / 漏れ / 証明の嘘 / 偽の疑い / 不親切 / 分けた（分けるべきでない）/ 正
★ ailine_core を import しない。対照: 答えをそのまま返す→全部正・黙る→漏れだけ・全部 1 人に→誤配。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

VERDICTS = ("正", "★誤配", "★混入", "漏れ", "証明の嘘", "偽の疑い", "不親切", "分けた（分けるべきでない）")
MIN_ROWS = 40     #: 分母の床（宣言された期待の行数）


def _norm(s) -> str:
    return "".join(str(s or "").split())


def judge_book(ans: dict, got: dict) -> tuple:
    """1 冊を裁く。戻り値 (内訳 Counter 風 dict, 外れ [(区分, 説明)], 分母)"""
    exp = ans.get("期待") or {}
    tally = {v: 0 for v in VERDICTS}
    notes = []
    people = exp.get("担当者ごと") or {}
    want_rows = {int(r): name for name, d in people.items() for r in (d.get("行") or [])}
    blank_rows = {int(r) for r in (exp.get("空欄の行") or [])}
    keep_out = {int(r) for r in (exp.get("分けない行") or [])}
    maybe = {int(m["行"]) for m in (ans.get("迷う") or []) if isinstance(m, dict) and m.get("行") is not None}
    denom = len(want_rows) + len(blank_rows) + len(keep_out)

    if exp.get("分けられない"):
        # ★ 列が決まらない冊: 分けたら誤配の予備軍
        if got.get("refused"):
            word = (exp["分けられない"] or {}).get("理由に含む語")
            if word and word not in str(got["refused"]):
                tally["不親切"] += 1; notes.append(("不親切", f"断ったが理由に「{word}」が無い"))
            else:
                tally["正"] += max(denom, 1)
        else:
            tally["分けた（分けるべきでない）"] += max(denom, 1)
            notes.append(("分けた（分けるべきでない）", "列が決まらない冊を分けた"))
        return tally, notes, max(denom, 1)

    if got.get("refused"):
        tally["漏れ"] += denom; notes.append(("漏れ", f"分けられる冊を断った: {str(got['refused'])[:60]}"))
        return tally, notes, denom

    placed = {}
    for name, part in (got.get("parts") or {}).items():
        for r in part.get("rows") or []:
            placed[int(r)] = name
    got_blank = {int(r) for r in (got.get("blank") or [])}
    got_out = {int(r) for r in (got.get("excluded") or [])}

    # 行ごと
    for r, name in want_rows.items():
        if r in placed:
            if _norm(placed[r]) == _norm(name) or r in maybe:
                tally["正"] += 1
            else:
                tally["★誤配"] += 1; notes.append(("★誤配", f"行 {r} は「{name}」のはずが「{placed[r]}」の冊へ"))
        elif r in maybe:
            tally["正"] += 1
        elif r in got_blank or r in got_out:
            tally["不親切"] += 1; notes.append(("不親切", f"行 {r}（{name}）を空欄／除外として扱った"))
        else:
            tally["漏れ"] += 1; notes.append(("漏れ", f"行 {r}（{name}）がどこにも無い"))
    for r in blank_rows:
        if r in placed:
            tally["★誤配"] += 1; notes.append(("★誤配", f"空欄の行 {r} を「{placed[r]}」の冊へ"))
        elif r in got_blank:
            tally["正"] += 1
        else:
            tally["不親切"] += 1; notes.append(("不親切", f"空欄の行 {r} を名指ししていない"))
    for r in keep_out:
        if r in placed:
            tally["★混入"] += 1; notes.append(("★混入", f"分けない行 {r} を「{placed[r]}」の冊へ"))
        else:
            tally["正"] += 1

    # 表記ゆれの名指し（併合は誤配として上で捕まる）
    want_pairs = {frozenset(_norm(x) for x in p) for p in (exp.get("表記ゆれ") or [])}
    got_pairs = {frozenset(_norm(x) for x in p) for p in (got.get("lookalike") or [])}
    for p in want_pairs - got_pairs:
        tally["不親切"] += 1; notes.append(("不親切", f"表記ゆれ {sorted(p)} を名指ししていない"))
    if ans.get("怪しくない"):
        if got_pairs or got_blank:
            tally["偽の疑い"] += 1; notes.append(("偽の疑い", f"怪しくない冊でゆれ {len(got_pairs)}・空欄 {len(got_blank)} を立てた"))

    # 証明の嘘: proof の和が出力と合うか
    proof = got.get("proof") or {}
    pr = (proof.get("rows") or {})
    if pr and pr.get("parts") != len(placed):
        tally["証明の嘘"] += 1; notes.append(("証明の嘘", f"proof.rows.parts={pr.get('parts')} だが実際 {len(placed)}"))
    pa = (proof.get("amount") or {})
    if pa and abs(float(pa.get("parts") or 0) - sum(float(p.get("amount") or 0) for p in (got.get("parts") or {}).values())) > 0.5:
        tally["証明の嘘"] += 1; notes.append(("証明の嘘", "proof.amount.parts が parts の和と違う"))
    return tally, notes, denom


def score(split, corpus: Path, answer: str, books: str) -> dict:
    total = {v: 0 for v in VERDICTS}; notes = []; denom = 0
    for ans in json.loads((corpus / answer).read_text(encoding="utf-8")):
        got = split(corpus / books / ans["file"], ans)
        t, n, d = judge_book(ans, got)
        for v in VERDICTS: total[v] += t[v]
        notes += [(ans["file"],) + x for x in n]; denom += d
    return {"内訳": total, "外れ": notes, "分母": denom}


def report(res: dict, show: int = 30) -> None:
    d = res["分母"]; t = res["内訳"]
    print(f"分母（宣言された行）{d}")
    if d < MIN_ROWS:
        print(f"★ 採点失敗: 分母が {d} しかない（{MIN_ROWS} 以上のはず）"); return
    for v in VERDICTS:
        print(f"  {v:14s} {t[v]:4d}")
    for x in res["外れ"][:show]:
        print("   ", x)


# ── 対照と変異 ─────────────────────────────────────────────
def _oracle(corpus, answer):
    def split(path, ans):
        exp = ans.get("期待") or {}
        if exp.get("分けられない"):
            return {"refused": "担当者の列が 2 つあり決められません（" + (exp["分けられない"].get("理由に含む語") or "") + "）"}
        parts = {n: {"rows": list(d.get("行") or []), "amount": d.get("金額")} for n, d in (exp.get("担当者ごと") or {}).items()}
        return {"parts": parts, "blank": list(exp.get("空欄の行") or []), "lookalike": list(exp.get("表記ゆれ") or []),
                "excluded": list(exp.get("分けない行") or []), "refused": None,
                "proof": {"rows": {"parts": sum(len(p["rows"]) for p in parts.values())},
                          "amount": {"parts": sum(float(p["amount"] or 0) for p in parts.values())}}}
    return split


def _mute(path, ans):
    return {"parts": {}, "blank": [], "lookalike": [], "excluded": [], "refused": None, "proof": {}}


def _one_bucket(corpus, answer):
    def split(path, ans):
        exp = ans.get("期待") or {}
        rows = sorted({r for d in (exp.get("担当者ごと") or {}).values() for r in d.get("行") or []}
                      | set(exp.get("空欄の行") or []) | set(exp.get("分けない行") or []))
        return {"parts": {"全員": {"rows": rows, "amount": 0}}, "blank": [], "lookalike": [], "excluded": [], "refused": None,
                "proof": {"rows": {"parts": len(rows)}, "amount": {"parts": 0}}}
    return split


def self_test(corpus: Path, answer: str, books: str) -> int:
    ok = True
    a = score(_oracle(corpus, answer), corpus, answer, books); n = a["分母"]
    good = a["内訳"]["正"] == n and n > 0
    print(f"★ 陽性対照: 正 {a['内訳']['正']}/{n}", "○" if good else "★"); ok &= good
    m = score(_mute, corpus, answer, books)
    good = m["内訳"]["★誤配"] == 0 and m["内訳"]["漏れ"] > 0
    print(f"★ 陰性対照: 漏れ {m['内訳']['漏れ']}・誤配 {m['内訳']['★誤配']}", "○" if good else "★"); ok &= good
    b = score(_one_bucket(corpus, answer), corpus, answer, books)
    good = b["内訳"]["★誤配"] > 0 and b["内訳"]["★混入"] > 0
    print(f"★ 変異（全員 1 冊に）: 誤配 {b['内訳']['★誤配']}・混入 {b['内訳']['★混入']}", "○" if good else "★"); ok &= good
    print("判定:", "○ 採点器は使える" if ok else "★ 採点器を直すまで点数を信じない")
    return 0 if ok else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=Path(r"C:\Dev\ailine\bench\split_people"))
    ap.add_argument("--answer", default="答え_split.json")
    ap.add_argument("--books", default="split_books")
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args()
    sys.exit(self_test(a.corpus, a.answer, a.books) if a.self_test else 0)
