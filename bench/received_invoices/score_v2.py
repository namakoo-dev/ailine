# -*- coding: utf-8 -*-
"""検体 v2（87 冊）に対する採点器。**点数を見る前に条件を凍結**（2026-09-11）。

★★ なぜ独立に書くか:
  2026-09-10 に、採点器そのものが恒真だった ── 結合セルの写しを 2 つ目の根拠として
  数え、40 冊すべてを「確」と報告した。畳んだら 確 14 / 単 26 だった。
  → だから **この採点器は `ailine_core.field_record` を import しない**。
    区分の意味・値の同一性・丸めを、本体とは別に、ここでもう一度書く。
    同じ関数で作った分母は分母でなく感想（docs/開発手法.md）。

★★ 採点の条件（**測る前に凍結**・あとから緩めない）:

  分母 = 検体を書いた側が宣言した `期待` の件数。★「取れた数」を分母にしない。
  1 件が **正** であるのは、次を**すべて**満たすときだけ:
      ① 区分が `期待.区分` の並びに入っている
      ② 値を出す区分（確 / 単）なら、値が `期待.値` と一致する
      ③ 出した値が `期待.禁止` に入っていない
      ④ 空欄の区分（割 / 無）なら、理由が `期待.名指し` の語を**すべて**含む
         （★ 昨日 Namakoo が凍結した「理由が適切かどうかも確認する」の機械化）

★★ 見出しにする数（★ 正答率を見出しにしない）:
      誤報   値を出して、その値が違う      ← **最悪**。買い手の $0 条件『黙って失敗する』
      安全   空欄にして、理由も名指せた    ← 取れていないが害は無い
      不親切 空欄だが理由が名指しを欠く    ← 空欄の意味が人に伝わらない
      正     上の条件を満たした
      落ち   例外で止まった

★ 陽性対照と陰性対照を必ず通す（`--self-test`）:
    答えをそのまま返す抽出器 → 100% でなければ**採点器が壊れている**
    何も返さない抽出器       → 0% であり、かつ「異常なし」と報告してはならない

抽出器の契約（この形だけ・`field_record` に依存しない）:
    extract(path: Path) -> {"請求元": {"区分": "確", "値": ..., "理由": ""}, ...}
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
#: 検体の置き場。★ このファイル自身の隣（repo の中の位置に依らない）。
DEFAULT_CORPUS = HERE

#: 値を出してよい区分。★ 本体から import せず、ここで独立に持つ（恒真を避ける）。
VALUE_GRADES = ("確", "単")
BLANK_GRADES = ("割", "無")
ALL_GRADES = VALUE_GRADES + BLANK_GRADES

#: 分母の床。★ これを下回ったら採点そのものを失敗にする
#:   （空集合に対して「○ 全問正解」と言った 2026-09-10 の自作恒真の再演を防ぐ）。
MIN_EXPECTATIONS = 150


def same_value(got, want) -> bool:
    """値の同一性。★ 本体の `_key` を**呼ばない**（別実装で書く）。

    数は小数第 2 位で見る（実物の消費税は ROUND(x,1) で小数が出る）。
    文字は前後の空白と改行を落として見る。
    """
    if isinstance(got, bool) or isinstance(want, bool):
        return got is want
    num = (int, float)
    if isinstance(got, num) and isinstance(want, num):
        return abs(float(got) - float(want)) < 0.005
    if isinstance(got, num) != isinstance(want, num):
        return False
    return str(got).strip().replace("\n", "") == str(want).strip().replace("\n", "")


def judge(got: dict, want: dict) -> tuple:
    """1 項目を裁く。戻り値: (区分, 一行の説明)

    区分は "正" / "誤報" / "安全" / "不親切" / "誤区分" / "無回答" のいずれか。
    """
    if got is None:
        return "無回答", "抽出器がこの項目を返さなかった"

    grade = got.get("区分")
    if grade not in ALL_GRADES:
        return "誤区分", f"知らない区分: {grade!r}"

    value = got.get("値")
    reason = str(got.get("理由") or "")

    # ── 値を出した場合 ────────────────────────────────
    if grade in VALUE_GRADES:
        for ng in want.get("禁止") or []:
            if same_value(value, ng):
                return "誤報", f"禁止された値を出した: {value!r}"
        # ★ `値: None` は「正解値を指定しない」── `禁止` に入っていなければ可。
        #   検体側が意図して置いている（個人事業主の「氏名（屋号：〇〇）」など、
        #   唯一の正解を決められない冊）。★ ここで値一致を強制すると、
        #   **正しい抽出を誤報と呼ぶ**（採点器がきつすぎて嘘をつく側の事故）。
        if want.get("値") is not None and not same_value(value, want["値"]):
            return "誤報", f"値が違う: 出した {value!r} / 正 {want['値']!r}"
        if grade not in want["区分"]:
            # ★ 値は合っているが、確信の度合いが宣言と違う
            #   （例: 根拠 1 つなのに「確」と言った＝写しを数えた疑い）
            return "誤区分", f"値は合っているが区分が違う: {grade} ∉ {want['区分']}"
        return "正", f"{grade} {value!r}"

    # ── 空欄にした場合 ────────────────────────────────
    # ★ 理由を最初に見る。2026-09-11 に凍結したのは「空欄には**必ず**理由を添える」で、
    #   検体が `名指し` を宣言しているかどうかとは関係が無い。
    #   ★ 初版はこの順序が逆で、名指しの宣言が無い項目の理由を**一度も見ていなかった**
    #     （変異②「全部『無』＋理由なし」が 134 件を「安全」と呼んで露見した）。
    if not reason.strip():
        return "不親切", "空欄なのに理由が空"
    missing = [w for w in (want.get("名指し") or []) if w not in reason]
    if missing:
        return "不親切", f"理由が {missing} を名指していない: {reason[:60]!r}"
    if grade not in want["区分"]:
        # ★ 理由は付いている。答えられたはずのものを空欄にした ── 取り逃しだが害は無い。
        return "安全", f"空欄にしたが宣言は {want['区分']}（理由: {reason[:40]!r}）"
    return "正", f"{grade}（理由に {want.get('名指し')} を名指した）"


#: ★ 規則を書くのに見てよい群（開発用）。残りは凍結まで開かない。
DEV_GROUPS = ("基礎",)

_ONLY_GROUPS: tuple = ()          #: 空なら全群

#: ★ どの検体の束を測るか（答えのファイル名 と 冊の置き場）。
#:   初版はこの 2 つを関数の中に直書きしていて、別の束を測るたびに書き換える形だった ──
#:   それは測るたびに片配線を作るのと同じ。**束は引数にする**（2026-09-11）。
ANSWER_FILE = "答え_received_v2.json"
BOOKS_DIR = "received_v2"


def use_bundle(answer_file: str, books_dir: str) -> None:
    """測る束を差し替える。★ 対照と変異は束ごとに通し直すこと。"""
    global ANSWER_FILE, BOOKS_DIR
    ANSWER_FILE, BOOKS_DIR = answer_file, books_dir


def load_books(corpus: Path) -> list:
    data = json.loads((corpus / ANSWER_FILE).read_text(encoding="utf-8"))
    out = [b for b in data if b.get("採用")]
    if _ONLY_GROUPS:
        out = [b for b in out if b.get("群") in _ONLY_GROUPS]
    return out


def score(extract, corpus: Path = DEFAULT_CORPUS, verbose: bool = False) -> dict:
    """抽出器を 87 冊にかけて採点する。"""
    books = load_books(corpus)
    tally = {k: 0 for k in ("正", "誤報", "安全", "不親切", "誤区分", "無回答", "落ち")}
    per_field: dict = {}
    misses: list = []
    n_expect = 0

    for b in books:
        path = corpus / BOOKS_DIR / b["file"]
        try:
            got_all = extract(path) or {}
        except Exception as e:                     # noqa: BLE001 ── 落ちたことも記録する
            got_all = None
            n_here = len(b.get("期待") or {})
            tally["落ち"] += n_here
            n_expect += n_here
            misses.append((b["id"], "*", "落ち", f"{type(e).__name__}: {e}"))
            continue

        for fld, want in (b.get("期待") or {}).items():
            n_expect += 1
            verdict, why = judge(got_all.get(fld), want)
            tally[verdict] += 1
            slot = per_field.setdefault(fld, {k: 0 for k in tally})
            slot[verdict] += 1
            if verdict != "正":
                misses.append((b["id"], fld, verdict, why))

    return {"分母": n_expect, "内訳": tally, "項目別": per_field,
            "外れ": misses, "冊数": len(books)}


def report(res: dict, show: int = 40) -> None:
    n = res["分母"]
    t = res["内訳"]
    print(f"検体 {res['冊数']} 冊 / 宣言された期待 {n} 件")
    if n < MIN_EXPECTATIONS:
        print(f"★ 採点失敗: 分母が {n} 件しかない（{MIN_EXPECTATIONS} 件以上のはず）── "
              "空集合に○を付けないための床")
        sys.exit(2)
    print("─" * 62)
    print(f"  正      {t['正']:4d}  ({t['正']/n:6.1%})")
    print(f"  ★誤報   {t['誤報']:4d}  ({t['誤報']/n:6.1%})  ← 値を出して間違えた（最悪）")
    print(f"  安全    {t['安全']:4d}  ({t['安全']/n:6.1%})  ← 空欄・害は無い")
    print(f"  不親切  {t['不親切']:4d}  ({t['不親切']/n:6.1%})  ← 空欄だが理由が足りない")
    print(f"  誤区分  {t['誤区分']:4d}  ({t['誤区分']/n:6.1%})")
    print(f"  無回答  {t['無回答']:4d}  ({t['無回答']/n:6.1%})")
    print(f"  落ち    {t['落ち']:4d}  ({t['落ち']/n:6.1%})")
    print("─" * 62)
    print("項目別:")
    for fld, s in sorted(res["項目別"].items()):
        m = sum(s.values())
        print(f"  {fld:6s} {m:3d} 件: 正 {s['正']:3d} / 誤報 {s['誤報']:3d} / "
              f"安全 {s['安全']:3d} / 不親切 {s['不親切']:3d} / 誤区分 {s['誤区分']:3d}")
    if res["外れ"]:
        print(f"\n外れ {len(res['外れ'])} 件（先頭 {show}）:")
        for bid, fld, v, why in res["外れ"][:show]:
            print(f"  {bid:>4s} {fld:6s} [{v}] {why}")


# ── 対照（★ 採点器そのものを疑う） ─────────────────────────
def _oracle(corpus: Path):
    """陽性対照: 答えをそのまま返す抽出器。★ 100% でなければ採点器が壊れている。"""
    books = {b["file"]: b for b in load_books(corpus)}

    def extract(path: Path) -> dict:
        b = books[path.name]
        out = {}
        for fld, want in (b.get("期待") or {}).items():
            g = want["区分"][0]
            if g in VALUE_GRADES:
                out[fld] = {"区分": g, "値": want["値"], "理由": ""}
            else:
                out[fld] = {"区分": g, "値": None,
                            "理由": "（対照）" + "・".join(want.get("名指し") or ["理由"])}
        return out
    return extract


def _mute(path: Path) -> dict:
    """陰性対照: 何も返さない抽出器。★ 0% であり『異常なし』であってはならない。"""
    return {}


def _expectation_shape(corpus: Path) -> dict:
    """対照の**期待値を答えから算術で出す**。★「> 0 なら合格」にしないため。

    `_oracle` は `期待.区分` の先頭を名乗るので、その先頭が何かで数える。
    """
    n_pinned = n_free = n_blank = n_no_kaku = 0
    for b in load_books(corpus):
        for want in (b.get("期待") or {}).values():
            head = want["区分"][0]
            if head in VALUE_GRADES:
                # ★ `値: None` は正解値を指定しない冊 ── 値をずらしても外れない
                if want.get("値") is None:
                    n_free += 1
                else:
                    n_pinned += 1
                if "確" not in want["区分"]:
                    n_no_kaku += 1
            else:
                n_blank += 1
    return {"値が定まる": n_pinned, "値は不問": n_free, "空欄": n_blank,
            "確が許されない": n_no_kaku}


def _wrong_by_one(corpus: Path):
    """★ 変異①: 値を 1 だけずらす。→ 全件 **誤報** と言えなければ採点器が甘い。"""
    inner = _oracle(corpus)

    def extract(path: Path) -> dict:
        out = inner(path)
        for r in out.values():
            v = r["値"]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                r["値"] = v + 1
            elif isinstance(v, str):
                r["値"] = v + "＿"
        return out
    return extract


def _blank_without_reason(corpus: Path):
    """★ 変異②: すべて「無」＋理由なし。→ **不親切** と言えなければ、
       昨日 Namakoo が凍結した『理由も見る』が採点に効いていない。"""
    books = {b["file"]: b for b in load_books(corpus)}

    def extract(path: Path) -> dict:
        return {fld: {"区分": "無", "値": None, "理由": ""}
                for fld in (books[path.name].get("期待") or {})}
    return extract


def _always_confirmed(corpus: Path):
    """★ 変異③: 値は正しいが常に「確」と言う（＝写しを 2 つ目に数えた形）。
       → 「確」が許されない冊で **誤区分** が立たなければ、区分を見ていない。"""
    inner = _oracle(corpus)

    def extract(path: Path) -> dict:
        out = inner(path)
        for r in out.values():
            if r["区分"] in VALUE_GRADES:
                r["区分"] = "確"
        return out
    return extract


def self_test(corpus: Path) -> int:
    print("★ 陽性対照（答えをそのまま返す）")
    a = score(_oracle(corpus), corpus)
    n, t = a["分母"], a["内訳"]
    print(f"   正 {t['正']}/{n}")
    ok1 = (t["正"] == n)
    print("   " + ("○ 採点器は正解を正解と言える" if ok1
                   else f"★ 壊れている ── 外れ例: {a['外れ'][:3]}"))

    print("★ 陰性対照（何も返さない）")
    b = score(_mute, corpus)
    t2 = b["内訳"]
    ok2 = (t2["正"] == 0 and t2["無回答"] == b["分母"])
    print(f"   正 {t2['正']}/{b['分母']} ・ 無回答 {t2['無回答']}")
    print("   " + ("○ 空の抽出器に点を与えない" if ok2 else "★ 壊れている"))

    print("★ 分母の床")
    print(f"   分母 {a['分母']} ≥ {MIN_EXPECTATIONS}: "
          + ("○" if a["分母"] >= MIN_EXPECTATIONS else "★ 足りない"))
    ok3 = a["分母"] >= MIN_EXPECTATIONS
    # ── ★ 採点器への変異（甘い採点器は緑のまま通ってしまう） ──────────
    # ★ 期待値は「> 0」で済ませず、答えから**算術で**出す。
    #   「> 0」だと、1 件だけ拾えている壊れた採点器が通る。
    shape = _expectation_shape(corpus)
    n_pin, n_free = shape["値が定まる"], shape["値は不問"]
    n_blank, n_no_kaku = shape["空欄"], shape["確が許されない"]
    n_all = n_pin + n_free + n_blank
    print(f"   （内訳: 値が定まる {n_pin} / 値は不問 {n_free} / 空欄が正解 {n_blank}）")

    # ★ 群を絞ると、当てる対象が 1 件も無い変異が出る。そこで「通った」と言うのも
    #   「落ちた」と言うのも嘘 ──「**測れなかった**」と分けて言う。
    untestable: list = []

    print(f"★ 変異① 値を 1 ずらす → 値が定まる {n_pin} 件が全部 誤報 になるべき")
    if n_pin == 0:
        ok4 = True
        untestable.append("変異①（値が定まる冊がこの群に 1 件も無い）")
        print("   ─ 測れない: 該当する冊がこの群に無い")
    else:
        m1 = score(_wrong_by_one(corpus), corpus)["内訳"]
        ok4 = (m1["誤報"] == n_pin and m1["正"] == n_blank + n_free)
        print(f"   誤報 {m1['誤報']}/{n_pin} ・ 正 {m1['正']}/{n_blank + n_free}"
              f"（空欄と『値は不問』はずらせない）→ " + ("○" if ok4 else "★ 素通り"))

    print("★ 変異② 全部『無』＋理由なし → 全件 不親切 になるべき（安全 にしない）")
    m2 = score(_blank_without_reason(corpus), corpus)["内訳"]
    ok5 = (m2["不親切"] == n_all and m2["正"] == 0 and m2["安全"] == 0)
    print(f"   不親切 {m2['不親切']}/{n_all} ・ 安全 {m2['安全']} ・ 正 {m2['正']}"
          + " → " + ("○" if ok5 else "★ 素通り ── 理由を見ていない"))

    print(f"★ 変異③ 値は正しいが常に『確』→ 確が許されない {n_no_kaku} 件が 誤区分 になるべき")
    if n_no_kaku == 0:
        ok6 = True
        untestable.append("変異③（確が許されない冊がこの群に 1 件も無い）")
        print("   ─ 測れない: この群では区分の厳しさを一度も試せない")
    else:
        m3 = score(_always_confirmed(corpus), corpus)["内訳"]
        ok6 = (m3["誤区分"] == n_no_kaku)
        print(f"   誤区分 {m3['誤区分']}/{n_no_kaku} → "
              + ("○" if ok6 else "★ 素通り ── 区分を見ていない"))

    good = all([ok1, ok2, ok3, ok4, ok5, ok6])
    if untestable:
        print("\n★ この群では試せなかった番人:")
        for u in untestable:
            print(f"   ─ {u}")
        print("   ★「効かなかった」ではなく「測れなかった」── 全 87 冊では必ず通すこと。")
    if not good:
        print("\n判定: ★ 採点器を直すまで、出た点数を信じない")
    elif untestable:
        print(f"\n判定: ○ 採点器は使える（ただし {len(untestable)} 件は"
              "この群では測れていない）")
    else:
        print("\n判定: ○ 採点器は使える（対照 2・変異 3 を通過）")
    return 0 if good else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", type=Path, default=DEFAULT_CORPUS)
    ap.add_argument("--answer", default=ANSWER_FILE, help="答えの JSON（束を差し替える）")
    ap.add_argument("--books", default=BOOKS_DIR, help="冊の置き場（束を差し替える）")
    ap.add_argument("--self-test", action="store_true",
                    help="採点器そのものを対照で確かめる（点数を見る前に必ず）")
    ap.add_argument("--show", type=int, default=40)
    ap.add_argument("--groups", default="",
                    help="群を絞る（例 基礎）。★ 規則を書いた群だけで測る時に使う")
    ap.add_argument("--floor", type=int, default=None,
                    help="分母の床を上書き（群を絞った時だけ）")
    a = ap.parse_args()
    use_bundle(a.answer, a.books)
    if a.groups:
        _ONLY_GROUPS = tuple(a.groups.split(","))
        globals()["_ONLY_GROUPS"] = _ONLY_GROUPS
    if a.floor is not None:
        globals()["MIN_EXPECTATIONS"] = a.floor
    if a.self_test:
        sys.exit(self_test(a.corpus))
    print("★ 抽出器がまだ無い。`--self-test` で採点器だけを確かめてください。")
