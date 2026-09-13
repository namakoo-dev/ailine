# -*- coding: utf-8 -*-
"""需要③（経費の勘定科目を先例から引く）の採点器 ── 検体が来る前に凍結（2026-09-13・設計レビュー後に改版）。

道具の契約（設計 docs/DESIGN-20260913-経費の勘定科目を先例から引く.md D2/D3・レビュー後）:
    accounts(today_path, past_paths, ans) ->
        {"rows": {行番号: {"account": 科目|None, "grade": "確"|"単"|"割"|"無", "reason": str}},
         "refused": None | 理由,
         "原本が変わった": bool}        # runner が入力の hash を前後で比べて入れる
    ★ `ans` は repo の作法（runner が採点器へ写す）── 本物を包むアダプタは **ans を無視する**。
      答えが道具に入る扉をここで開けない。

答え（答え_accounts.json）は 1 冊 1 件:
    {"id", "software": "mf"|"yayoi"|"freee"|"xlsx", "today": ファイル名, "past": [ファイル名…],
     "先例に在る科目": [科目…],             # ★ 必須（捏造の判定の分母・割/無 の行の科目もここに）
     "触らない行": [行番号…],              # ★ 今回の冊で候補を出してはいけない行
                                           #   （既に埋まっている行・複合仕訳の継続行・合計行）
     "複合仕訳": bool, "貸方だけ取引先": bool, "文字コード": "utf-8-sig"|"cp932",
     "期待": {"行": {行番号: {"科目": 科目|null, "区分": "確"|"単"|"割"|"無",
                              "根拠に含む語": 語|null}}},
     "怪しくない": bool}                    # 割も無も 1 つも無い冊（偽の疑いを測る）

区分の導出（★ レビュー致命1 の処方 ── 鍵ごと＝出所ごと。同じ鍵の過去 N 件は 1 出所）:
    鍵は 借方取引先 / 貸方取引先 / 借方補助科目 / 摘要（norm で畳んだ完全一致）
    確: 2 つ以上の鍵が当たり、それぞれ 1 科目に決まり、全部同じ科目で、食い違う鍵が無い
        ★ ただし当たった鍵の先例が**同じ過去 1 行だけ**なら 単（同じ行を 2 通りに読んだだけ・
          裏は 1 つ ── 実装者が踏んで名指しした穴・2026-09-13）
    単: 当たった鍵が 1 つだけで、その鍵は 1 科目に決まる
    割: どれかの鍵の内訳が 2 科目以上、または鍵どうしが違う科目を指す（値は出さない）
    無: どの鍵にも先例が無い（値は出さない）

見出し（★ 正答率にしない）:
    正 / ★誤配（先例に在る別の科目を出した）/ ★捏造（先例に無い科目を出した）/
    ★余計（触らない行や宣言外の行に候補を出した）/ 原本が変わった / 漏れ（出せるのに空欄）/
    格上げ（単→確・割→単 等）/ 格下げ（確→単 等）/ 不親切（理由に語が無い）/ 偽の疑い（1 冊 1 点）
★ ailine_core を import しない（同じ関数で作った分母は恒真）。
対照: 答えをそのまま返す→全部正・黙る→漏れだけ・全部に多数派を確で→誤配と格上げ・
      単を確に格上げする版→格上げ・割/無の行にも先例の科目を出す版→誤配。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

VERDICTS = ("正", "★誤配", "★捏造", "★余計", "原本が変わった", "漏れ",
            "格上げ", "格下げ", "不親切", "偽の疑い")
PER_BOOK = ("偽の疑い", "原本が変わった")     #: ★ 1 冊 1 点（分母に閉じない）── report に明記
GRADES = ("確", "単", "割", "無")
_RANK = {"確": 3, "単": 2, "割": 1, "無": 0}
MIN_ROWS = 60           #: 分母の床（宣言された期待の行数）
MIN_PER_GRADE = 8       #: ★ 区分ごとの床 ── 全部 単 の 60 行では変異を殺せない
REQUIRED = ("先例に在る科目", "触らない行", "software", "文字コード")


def _norm(s) -> str:
    return "".join(str(s or "").split())


def _check_answer(ans: dict) -> None:
    missing = [k for k in REQUIRED if k not in ans]
    if missing:
        raise ValueError(f"答え {ans.get('id')} に必須項目が無い: {missing}")


def judge_case(ans: dict, got: dict) -> tuple:
    """1 冊を裁く。戻り値 (内訳 dict, 外れ [(区分, 説明)], 分母)。"""
    _check_answer(ans)
    exp = (ans.get("期待") or {}).get("行") or {}
    past_accounts = {_norm(a) for a in ans["先例に在る科目"]}
    untouchable = {int(r) for r in ans["触らない行"]}
    tally = {v: 0 for v in VERDICTS}
    notes = []
    denom = len(exp)
    if got.get("原本が変わった"):
        tally["原本が変わった"] += 1; notes.append(("原本が変わった", "入力の hash が前後で違う"))
    if got.get("refused"):
        tally["漏れ"] += denom
        notes.append(("漏れ", f"断った: {str(got['refused'])[:60]}"))
        return tally, notes, denom

    rows = {int(k): v for k, v in (got.get("rows") or {}).items()}
    # ★ 負の被覆: 宣言外の行・触らない行に値を出したら 余計（レビュー致命4）
    for r, out in rows.items():
        if (out or {}).get("account") and (r in untouchable or str(r) not in exp and r not in {int(k) for k in exp}):
            tally["★余計"] += 1; notes.append(("★余計", f"行 {r}: 候補を出してはいけない行に {out['account']}"))

    for key, want in exp.items():
        r = int(key)
        out = rows.get(r) or {}
        acc, grade, reason = out.get("account"), out.get("grade"), str(out.get("reason") or "")
        want_acc, want_grade, word = want.get("科目"), want.get("区分"), want.get("根拠に含む語")
        if grade not in GRADES:
            tally["漏れ"] += 1; notes.append(("漏れ", f"行 {r}: 区分が無い（{grade!r}）")); continue
        if acc and (not want_acc or _norm(acc) != _norm(want_acc)):
            kind = "★誤配" if _norm(acc) in past_accounts else "★捏造"
            tally[kind] += 1
            notes.append((kind, f"行 {r}: {want_acc or f'空欄（{want_grade}）'} のはずが {acc}"))
            continue
        if not acc and want_acc:
            tally["漏れ"] += 1; notes.append(("漏れ", f"行 {r}: {want_acc} を出せるのに空欄（{grade}）")); continue
        if _RANK[grade] > _RANK[want_grade]:
            tally["格上げ"] += 1; notes.append(("格上げ", f"行 {r}: {want_grade} を {grade} にした")); continue
        if _RANK[grade] < _RANK[want_grade]:
            tally["格下げ"] += 1; notes.append(("格下げ", f"行 {r}: {want_grade} を {grade} にした")); continue
        if word and _norm(word) not in _norm(reason):
            tally["不親切"] += 1; notes.append(("不親切", f"行 {r}: 理由に「{word}」が無い")); continue
        tally["正"] += 1
    if ans.get("怪しくない"):
        doubted = [r for r, o in rows.items() if (o or {}).get("grade") in ("割", "無")]
        if doubted:
            tally["偽の疑い"] += 1
            notes.append(("偽の疑い", f"怪しくない冊で 割/無 を {len(doubted)} 行立てた"))
    return tally, notes, denom


def score(accounts, corpus: Path, answer: str, books: str) -> dict:
    total = {v: 0 for v in VERDICTS}; notes = []; denom = 0
    for ans in json.loads((corpus / answer).read_text(encoding="utf-8")):
        got = accounts(corpus / books / ans["today"],
                       [corpus / books / p for p in ans.get("past") or []], ans)
        t, n, d = judge_case(ans, got)
        for v in VERDICTS: total[v] += t[v]
        notes += [(ans["id"],) + x for x in n]; denom += d
    return {"内訳": total, "外れ": notes, "分母": denom}


def report(res: dict, show: int = 30) -> None:
    d = res["分母"]; t = res["内訳"]
    print(f"分母（宣言された行）{d}   ★ {'・'.join(PER_BOOK)} は 1 冊 1 点で分母に閉じない")
    if d < MIN_ROWS:
        print(f"★ 採点失敗: 分母が {d} しかない（{MIN_ROWS} 以上のはず）"); return
    for v in VERDICTS:
        print(f"  {v:10s} {t[v]:4d}")
    for x in res["外れ"][:show]:
        print("   ", x)


# ── 分布の床（★ 行数だけでは足りない ── 全部 単 の 60 行でも通ってしまう）────────
def distribution_floor(answers: list) -> list:
    """事前登録した床に届いていない項目を返す（空なら合格）。"""
    bad = []
    grades = {g: 0 for g in GRADES}
    for ans in answers:
        for v in ((ans.get("期待") or {}).get("行") or {}).values():
            if v.get("区分") in grades:
                grades[v["区分"]] += 1
    for g, n in grades.items():
        if n < MIN_PER_GRADE:
            bad.append(f"区分 {g} が {n} 行（{MIN_PER_GRADE} 以上）")
    softs = {a.get("software") for a in answers}
    if len(softs - {"freee"}) < 2:
        bad.append(f"測る対象のソフトが {sorted(softs)}（freee を除いて 2 種以上）")
    if not any(a.get("複合仕訳") for a in answers):
        bad.append("複合仕訳の冊が 0")
    if not any(a.get("貸方だけ取引先") for a in answers):
        bad.append("貸方側だけに取引先が在る冊が 0")
    if len({a.get("文字コード") for a in answers}) < 2:
        bad.append("文字コードが 1 種類")
    if not any(a.get("怪しくない") for a in answers):
        bad.append("怪しくない冊が 0")
    return bad


# ── 対照と変異 ─────────────────────────────────────────────
def _oracle(today, past, ans):
    exp = (ans.get("期待") or {}).get("行") or {}
    return {"rows": {int(r): {"account": v.get("科目"), "grade": v.get("区分"),
                              "reason": v.get("根拠に含む語") or ""} for r, v in exp.items()},
            "refused": None, "原本が変わった": False}


def _mute(today, past, ans):
    exp = (ans.get("期待") or {}).get("行") or {}
    return {"rows": {int(r): {"account": None, "grade": "無", "reason": "黙る"} for r in exp},
            "refused": None, "原本が変わった": False}


def _majority(today, past, ans):
    """全部の行（触らない行も）に、先例で一番多い科目を 確 で出す（推測する道具の姿）。"""
    exp = (ans.get("期待") or {}).get("行") or {}
    counts: dict = {}
    for v in exp.values():
        if v.get("科目"):
            counts[v["科目"]] = counts.get(v["科目"], 0) + 1
    top = max(counts, key=counts.get) if counts else (ans.get("先例に在る科目") or ["雑費"])[0]
    rows = {int(r): {"account": top, "grade": "確", "reason": "多数派"} for r in exp}
    for r in ans.get("触らない行") or []:
        rows[int(r)] = {"account": top, "grade": "確", "reason": "多数派"}
    return {"rows": rows, "refused": None, "原本が変わった": False}


def _upgrade(today, past, ans):
    """答えどおりだが 単 を 確 に格上げする（先例 1 件でも確にする版）。"""
    got = _oracle(today, past, ans)
    for o in got["rows"].values():
        if o["grade"] == "単":
            o["grade"] = "確"
    return got


def _greedy(today, past, ans):
    """答えどおりだが 割/無 の行にも先例に在る科目を出す（部分一致を許す版の姿）。"""
    got = _oracle(today, past, ans)
    pool = ans.get("先例に在る科目") or []
    for o in got["rows"].values():
        if o["grade"] in ("割", "無") and pool:
            o["account"], o["grade"] = pool[0], "単"
    return got


def self_test(corpus: Path, answer: str, books: str) -> int:
    ok = True
    answers = json.loads((corpus / answer).read_text(encoding="utf-8"))
    for ans in answers:
        _check_answer(ans)
    floor = distribution_floor(answers)
    print("★ 分布の床:", "○" if not floor else f"★ {floor}"); ok &= not floor

    a = score(_oracle, corpus, answer, books); n = a["分母"]
    good = a["内訳"]["正"] == n and n >= MIN_ROWS
    print(f"★ 陽性対照: 正 {a['内訳']['正']}/{n}", "○" if good else "★"); ok &= good
    m = score(_mute, corpus, answer, books)
    good = m["内訳"]["★誤配"] == 0 and m["内訳"]["★捏造"] == 0 and m["内訳"]["漏れ"] > 0
    print(f"★ 陰性対照（黙る）: 漏れ {m['内訳']['漏れ']}・誤配 {m['内訳']['★誤配']}", "○" if good else "★"); ok &= good
    g = score(_majority, corpus, answer, books)
    good = g["内訳"]["★誤配"] > 0 and g["内訳"]["★余計"] > 0
    print(f"★ 変異（全部に多数派を 確 で）: 誤配 {g['内訳']['★誤配']}・余計 {g['内訳']['★余計']}", "○" if good else "★"); ok &= good
    u = score(_upgrade, corpus, answer, books)
    good = u["内訳"]["格上げ"] > 0
    print(f"★ 変異（単を確に）: 格上げ {u['内訳']['格上げ']}", "○" if good else "★"); ok &= good
    r = score(_greedy, corpus, answer, books)
    good = r["内訳"]["★誤配"] > 0
    print(f"★ 変異（割/無にも科目を出す）: 誤配 {r['内訳']['★誤配']}", "○" if good else "★"); ok &= good
    print("判定:", "○ 採点器は使える" if ok else "★ 採点器を直すまで点数を信じない")
    return 0 if ok else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=Path(__file__).resolve().parent)
    ap.add_argument("--answer", default="答え_accounts.json")
    ap.add_argument("--books", default="accounts_books")
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args()
    sys.exit(self_test(a.corpus, a.answer, a.books) if a.self_test else 0)
