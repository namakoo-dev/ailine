# -*- coding: utf-8 -*-
"""束（フォルダ）の所見の採点器 ── ★誤報を主役にする版（monthend・2026-09-13）。

見出しを score_bundle.py から踏襲しつつ並べ替えた ── 「正答率」ではなく、
どれだけ怪しくないものを疑ったか（★誤報）と、どれだけ拾い損ねたか（★取り逃し）を
別々の数として出す（docs/検体の要件.md §4.3）。

答えの形（検体を書く側が宣言。mk_monthend.py が specimens_monthend.py から機械的に作る）:
    {"束": 名, "冊": [...], "疑い": [{"種類", "冊", "名指し"}...], "怪しくない": [...]}

道具の出力の契約（本体に依存しない・この形だけ ── score_bundle.py と同じ）:
    suspect(folder: Path) -> [{"種類": str, "冊": [ファイル名...], "理由": str}, ...]

1 つの疑いが **正** であるのは:
  ① 道具が同じ種類の疑いを出し、関わる冊を**全部**名指ししている
  ② 理由に「名指し」の語が**すべて**入っている

見出しにする数（★ 正答率にしない）:
    正        鳴るべきものが、正しい種類・正しい名指しで鳴った
    ★誤報    「怪しくない」と宣言された冊に疑いを立てた   ← 最悪（オオカミ少年）
    ★取り逃し 仕込んだ疑いに、道具がどんな種類でも触れなかった
    種類違い  疑いの対象（冊）は捕まえたが、種類が違う（例: 二重払いを訂正再発行と呼んだ）
    不親切    種類は合っているが、冊が欠ける／名指し語が無い

★ この採点器は ailine_core を import しない（同じ関数で作った分母は分母でなく感想）。
★ 対照: 答えをそのまま返す → 全部正 ／ 何も返さない → 取り逃しだけ ／ 全冊を疑う → ★誤報が立つ。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from bundle_common import load_bundles

VERDICTS = ("正", "★誤報", "★取り逃し", "種類違い", "不親切")

#: ★ 分布の床（docs/検体の要件.md §1.3）。行数でなく分布を機械で確かめる。
MIN_BUNDLES = 6
MIN_INNOCENT_RATIO = 0.7     #: ★誤報の分母（怪しくない冊 / 全冊）はこれ以上
MIN_TWIN_PAYMENT = 1          #: 二重払いの典型（同額・番号違い）を最低 1 組


def _norm(s) -> str:
    return "".join(str(s or "").split())


def judge_bundle(answer: dict, found: list) -> dict:
    """1 束を裁く。"""
    tally = {v: 0 for v in VERDICTS}
    notes = []
    innocent = set(answer.get("怪しくない") or [])
    used = [False] * len(found)

    for want in answer.get("疑い") or []:
        kind, books, names = want["種類"], set(want["冊"]), want.get("名指し") or []

        # ① まず同じ種類で冊が重なるものを探す（score_bundle.py と同じ「一番重なる」選び方）
        best_same, best_same_i = None, None
        for i, got in enumerate(found):
            if used[i] or got.get("種類") != kind:
                continue
            overlap = books & set(got.get("冊") or [])
            if overlap and (best_same is None or len(overlap) > len(best_same[0])):
                best_same, best_same_i = (overlap, got), i

        if best_same is not None:
            used[best_same_i] = True
            got = best_same[1]
            missing_books = books - set(got.get("冊") or [])
            missing_words = [w for w in names if _norm(w) not in _norm(got.get("理由", ""))]
            if missing_books or missing_words:
                tally["不親切"] += 1
                notes.append((answer["束"], kind, "不親切",
                              f"欠けた冊 {sorted(missing_books)} ／ 無い語 {missing_words}"))
            else:
                tally["正"] += 1
            dragged = innocent & set(got.get("冊") or [])
            if dragged:
                tally["★誤報"] += 1
                notes.append((answer["束"], kind, "★誤報",
                              f"当たりの疑いに怪しくない冊を巻き込んだ: {sorted(dragged)}"))
            continue

        # ② 種類は違うが、同じ冊に何か疑いが立っている（＝捕まえてはいるが呼び方が違う）
        best_other, best_other_i = None, None
        for i, got in enumerate(found):
            if used[i]:
                continue
            overlap = books & set(got.get("冊") or [])
            if overlap and (best_other is None or len(overlap) > len(best_other[0])):
                best_other, best_other_i = (overlap, got), i

        if best_other is not None:
            used[best_other_i] = True
            got = best_other[1]
            tally["種類違い"] += 1
            notes.append((answer["束"], kind, "種類違い",
                          f"欲しい種類={kind} ／ 出た種類={got.get('種類')} ／"
                          f" 冊={sorted(best_other[0])}"))
            dragged = innocent & set(got.get("冊") or [])
            if dragged:
                tally["★誤報"] += 1
                notes.append((answer["束"], got.get("種類"), "★誤報",
                              f"種類違いの疑いに怪しくない冊を巻き込んだ: {sorted(dragged)}"))
            continue

        # ③ 何も触れていない
        tally["★取り逃し"] += 1
        notes.append((answer["束"], kind, "★取り逃し", f"{sorted(books)} を疑わなかった"))

    # 余った疑い（どの「疑い」にも使われなかった found）── 怪しくない冊に触れていれば誤報
    for i, got in enumerate(found):
        if used[i]:
            continue
        touched = innocent & set(got.get("冊") or [])
        if touched:
            tally["★誤報"] += 1
            notes.append((answer["束"], got.get("種類"), "★誤報",
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


def report(res: dict, show: int = 40) -> None:
    t = res["内訳"]
    print(f"束 {res['束数']} ／ 仕込んだ疑い {res['仕込んだ疑い']} ／ 怪しくない冊 {res['怪しくない冊']}")
    print("─" * 64)
    print(f"  正          {t['正']:3d}")
    print(f"  ★誤報       {t['★誤報']:3d}   ← 怪しくない冊を疑った（最悪・オオカミ少年）")
    print(f"  ★取り逃し   {t['★取り逃し']:3d}   ← 仕込んだ疑いに何も触れなかった")
    print(f"  種類違い    {t['種類違い']:3d}   ← 冊は捕まえたが種類の呼び方が違う")
    print(f"  不親切      {t['不親切']:3d}   ← 種類は合っているが冊か語が欠ける")
    for b, k, v, why in res["外れ"][:show]:
        print(f"    {b} / {k} [{v}] {why}")
    if len(res["外れ"]) > show:
        print(f"    ...ほか {len(res['外れ']) - show} 件（--show で増やせる）")


# ── ★ 分布の床（検体を見る前に凍結・行数でなく分布） ──────────────────────
def distribution_floor(bundles: list) -> list:
    """事前登録した床に届いていない項目を返す（空なら合格）。docs/検体の要件.md §1.3。"""
    bad = []
    if len(bundles) < MIN_BUNDLES:
        bad.append(f"束が {len(bundles)}（{MIN_BUNDLES} 以上）")
    total_books = sum(len(b.get("冊") or []) for b in bundles)
    total_innocent = sum(len(b.get("怪しくない") or []) for b in bundles)
    ratio = (total_innocent / total_books) if total_books else 0.0
    if ratio < MIN_INNOCENT_RATIO:
        bad.append(f"怪しくない冊の比率が {ratio:.0%}（{MIN_INNOCENT_RATIO:.0%} 以上）")
    n_twin = sum(1 for b in bundles for w in (b.get("疑い") or []) if w.get("二重払い典型"))
    if n_twin < MIN_TWIN_PAYMENT:
        bad.append(f"二重払いの典型が {n_twin} 組（{MIN_TWIN_PAYMENT} 以上）")
    return bad


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
    """★ 何でも疑う道具 ── 怪しくない冊を含めて全部を『重複』と言う。★誤報が立たなければ採点器が甘い。"""
    bundles = {b["束"]: b for b in load_bundles(corpus / answer_file)}

    def suspect(folder: Path) -> list:
        b = bundles[folder.name]
        names = [x["file"] if isinstance(x, dict) else x for x in b.get("冊") or []]
        return [{"種類": "重複", "冊": names, "理由": "全部同じに見える"}]
    return suspect


def _random_kind(corpus: Path, answer_file: str):
    """★ 変異その2 ── 冊は当てるが種類をいつも違うものにする道具（種類違いを殺す変異）。"""
    bundles = {b["束"]: b for b in load_bundles(corpus / answer_file)}

    def suspect(folder: Path) -> list:
        b = bundles[folder.name]
        out = []
        for w in b.get("疑い") or []:
            wrong = "空欄" if w["種類"] != "空欄" else "重複"
            out.append({"種類": wrong, "冊": list(w["冊"]), "理由": "（変異）種類をわざと外す"})
        return out
    return suspect


def self_test(corpus: Path, answer_file: str, books_root: str) -> int:
    ok = True
    bundles = load_bundles(corpus / answer_file)

    bad = distribution_floor(bundles)
    print("★ 分布の床:", "○ 全部届いている" if not bad else "★ " + " / ".join(bad))
    ok &= not bad

    a = score(_oracle(corpus, answer_file), corpus, answer_file, books_root)
    n = a["仕込んだ疑い"]
    print(f"★ 陽性対照: 正 {a['内訳']['正']}/{n}", "○" if a["内訳"]["正"] == n and n > 0 else "★")
    ok &= a["内訳"]["正"] == n and n > 0

    m = score(_mute, corpus, answer_file, books_root)
    print(f"★ 陰性対照: ★取り逃し {m['内訳']['★取り逃し']}/{n}・★誤報 {m['内訳']['★誤報']}",
          "○" if m["内訳"]["★取り逃し"] == n and m["内訳"]["★誤報"] == 0 else "★")
    ok &= m["内訳"]["★取り逃し"] == n and m["内訳"]["★誤報"] == 0

    p = score(_paranoid(corpus, answer_file), corpus, answer_file, books_root)
    print(f"★ 変異（全部疑う）: ★誤報 {p['内訳']['★誤報']}",
          "○" if p["内訳"]["★誤報"] > 0 else "★ 素通り ── 怪しくない冊を見ていない")
    ok &= p["内訳"]["★誤報"] > 0

    r = score(_random_kind(corpus, answer_file), corpus, answer_file, books_root)
    print(f"★ 変異（種類を外す）: 種類違い {r['内訳']['種類違い']}",
          "○" if r["内訳"]["種類違い"] > 0 else "★ 素通り ── 種類の一致を見ていない")
    ok &= r["内訳"]["種類違い"] > 0

    print("判定:", "○ 採点器は使える" if ok else "★ 採点器を直すまで点数を信じない")
    return 0 if ok else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=Path(__file__).resolve().parent)
    ap.add_argument("--answer", default="答え_monthend.json")
    ap.add_argument("--books", default="monthend_books")
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args()
    if a.self_test:
        sys.exit(self_test(a.corpus, a.answer, a.books))
    print("★ 抽出器は run_monthend.py 側から渡す。--self-test で採点器だけを確かめる。")
