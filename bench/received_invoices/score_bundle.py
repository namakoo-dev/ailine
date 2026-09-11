# -*- coding: utf-8 -*-
"""束（フォルダ）単位の採点器 ── 条件は検体が来る前に凍結（crossfile_plan.md・2026-09-11）。

答えの形（検体を書く側が宣言）:
    {"束": 名, "冊": [...], "疑い": [{"種類", "冊", "名指し"}...], "怪しくない": [...]}

道具の出力の契約（本体に依存しない・この形だけ）:
    suspect(folder: Path) -> [{"種類": str, "冊": [ファイル名...], "理由": str}, ...]

1 つの疑いが **正** であるのは:
  ① 道具が同じ種類の疑いを出し、関わる冊を**全部**名指ししている
  ② 理由に「名指し」の語が**すべて**入っている
見出しにする数（★ 正答率にしない）:
  偽の疑い  「怪しくない」と宣言された冊に疑いを立てた   ← 最悪（オオカミ少年）
  取り逃し  仕込んだ疑いを出さなかった
  不親切    疑いは出したが、冊が欠ける／名指し語が無い
  正

★ この採点器は ailine_core を import しない（同じ関数で作った分母は分母でなく感想）。
★ 対照: 答えをそのまま返す → 全部正 ／ 何も返さない → 取り逃しだけ ／ 全冊を疑う → 偽の疑いが立つ。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from bundle_common import load_bundles

VERDICTS = ("正", "偽の疑い", "取り逃し", "不親切")


def _norm(s: str) -> str:
    return "".join(str(s or "").split())


def judge_bundle(answer: dict, found: list) -> dict:
    """1 束を裁く。"""
    tally = {v: 0 for v in VERDICTS}
    notes = []
    innocent = set(answer.get("怪しくない") or [])
    used = [False] * len(found)

    # 仕込んだ疑いごとに、道具の出力を当てにいく
    for want in answer.get("疑い") or []:
        kind, books, names = want["種類"], set(want["冊"]), want.get("名指し") or []
        best, best_i = None, None
        for i, got in enumerate(found):
            if used[i] or got.get("種類") != kind:
                continue
            overlap = books & set(got.get("冊") or [])
            if overlap and (best is None or len(overlap) > len(best[0])):
                best, best_i = (overlap, got), i
        if best is None:
            tally["取り逃し"] += 1
            notes.append((answer["束"], kind, "取り逃し", f"{sorted(books)} を疑わなかった"))
            continue
        used[best_i] = True
        got = best[1]
        missing_books = books - set(got.get("冊") or [])
        missing_words = [w for w in names if _norm(w) not in _norm(got.get("理由", ""))]
        if missing_books or missing_words:
            tally["不親切"] += 1
            notes.append((answer["束"], kind, "不親切",
                          f"欠けた冊 {sorted(missing_books)} ／ 無い語 {missing_words}"))
        else:
            tally["正"] += 1
        # ★ 当たった疑いでも、怪しくない冊を巻き込んでいれば偽の疑い（巻き込みは無罪にならない）
        dragged = innocent & set(got.get("冊") or [])
        if dragged:
            tally["偽の疑い"] += 1
            notes.append((answer["束"], kind, "偽の疑い",
                          f"当たりの疑いに怪しくない冊を巻き込んだ: {sorted(dragged)}"))

    # 余った疑い ── 怪しくない冊に触れていれば偽の疑い
    for i, got in enumerate(found):
        if used[i]:
            continue
        touched = innocent & set(got.get("冊") or [])
        if touched:
            tally["偽の疑い"] += 1
            notes.append((answer["束"], got.get("種類"), "偽の疑い",
                          f"怪しくない冊に疑いを立てた: {sorted(touched)}"))
    return {"内訳": tally, "外れ": notes, "仕込み": len(answer.get("疑い") or []),
            "怪しくない": len(innocent)}


def score(suspect, corpus: Path, answer_file: str, books_root: str) -> dict:
    bundles = load_bundles(corpus / answer_file)
    total = {v: 0 for v in VERDICTS}
    notes, n_want, n_innocent = [], 0, 0
    for b in bundles:
        found = suspect(corpus / books_root / b["束"])
        r = judge_bundle(b, found)
        for v in VERDICTS:
            total[v] += r["内訳"][v]
        notes += r["外れ"]
        n_want += r["仕込み"]
        n_innocent += r["怪しくない"]
    return {"束数": len(bundles), "仕込んだ疑い": n_want, "怪しくない冊": n_innocent,
            "内訳": total, "外れ": notes}


def report(res: dict, show: int = 30) -> None:
    t = res["内訳"]
    print(f"束 {res['束数']} ／ 仕込んだ疑い {res['仕込んだ疑い']} ／ 怪しくない冊 {res['怪しくない冊']}")
    print("─" * 60)
    print(f"  正        {t['正']:3d}")
    print(f"  ★偽の疑い {t['偽の疑い']:3d}   ← 怪しくない冊を疑った（最悪）")
    print(f"  取り逃し  {t['取り逃し']:3d}")
    print(f"  不親切    {t['不親切']:3d}   ← 疑いは出したが冊か語が欠ける")
    for b, k, v, why in res["外れ"][:show]:
        print(f"    {b} / {k} [{v}] {why}")


# ── 対照と変異（★ 点数を見る前に必ず） ────────────────────────
def _oracle(corpus: Path, answer_file: str):
    bundles = {b["束"]: b for b in load_bundles(corpus / answer_file)}

    def suspect(folder: Path) -> list:
        b = bundles[folder.name]
        return [{"種類": w["種類"], "冊": list(w["冊"]),
                 "理由": "（対照）" + "・".join(w.get("名指し") or ["理由"])}
                for w in b.get("疑い") or []]
    return suspect


def _mute(folder: Path) -> list:
    return []


def _paranoid(corpus: Path, answer_file: str):
    """★ 何でも疑う道具 ── 怪しくない冊を含めて全部を『重複』と言う。偽の疑いが立たなければ採点器が甘い。"""
    bundles = {b["束"]: b for b in load_bundles(corpus / answer_file)}

    def suspect(folder: Path) -> list:
        b = bundles[folder.name]
        names = [x["file"] if isinstance(x, dict) else x for x in b.get("冊") or []]
        return [{"種類": "重複", "冊": names, "理由": "全部同じに見える"}]
    return suspect


def self_test(corpus: Path, answer_file: str, books_root: str) -> int:
    ok = True
    a = score(_oracle(corpus, answer_file), corpus, answer_file, books_root)
    n = a["仕込んだ疑い"]
    print(f"★ 陽性対照: 正 {a['内訳']['正']}/{n}", "○" if a["内訳"]["正"] == n and n > 0 else "★")
    ok &= a["内訳"]["正"] == n and n > 0
    m = score(_mute, corpus, answer_file, books_root)
    print(f"★ 陰性対照: 取り逃し {m['内訳']['取り逃し']}/{n}・偽の疑い {m['内訳']['偽の疑い']}",
          "○" if m["内訳"]["取り逃し"] == n and m["内訳"]["偽の疑い"] == 0 else "★")
    ok &= m["内訳"]["取り逃し"] == n and m["内訳"]["偽の疑い"] == 0
    p = score(_paranoid(corpus, answer_file), corpus, answer_file, books_root)
    print(f"★ 変異（全部疑う）: 偽の疑い {p['内訳']['偽の疑い']}",
          "○" if p["内訳"]["偽の疑い"] > 0 else "★ 素通り ── 怪しくない冊を見ていない")
    ok &= p["内訳"]["偽の疑い"] > 0
    print("判定:", "○ 採点器は使える" if ok else "★ 採点器を直すまで点数を信じない")
    return 0 if ok else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=Path(__file__).resolve().parent)
    ap.add_argument("--answer", default="答え_received_bundle.json")
    ap.add_argument("--books", default="received_bundle")
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args()
    if a.self_test:
        sys.exit(self_test(a.corpus, a.answer, a.books))
    print("★ 抽出器は run 側から渡す。--self-test で採点器だけを確かめる。")
