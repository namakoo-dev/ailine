# -*- coding: utf-8 -*-
"""同じ依頼を N 回翻訳して、**揺れの本当の大きさ**を測る（2026-09-18）。

★★ なぜ要るか（Namakoo「揺れが絡んできたのは厄介だな」→「揺れだと判定できるかどうかも
  重要だ」）: 履歴（`~/.ailine/history.jsonl`）から出した揺れ率は **コード変更をまたいで
  いる** ── 俺たちが直したせいで op が変わった回まで「揺れ」に数えていた。実際:

      試行数を揃えて（10 回以上打たれた依頼）
        全期間      263 種（平均 53.9 回）  揺れた 17.1%
        09-01 以降  217 種（平均 51.4 回）  揺れた  7.4%

  ★ だから履歴の数字は**上限**でしかない。本当の数字は **HEAD を固定して振り直す**
    しか出せない（棚の線: 模型でなく本番が残す記録の上で測る／測定器を疑う）。

★★ 測るのは**翻訳だけ**（適用しない）。理由は 2 つ:
  ① 揺れの源は翻訳（機械が決める座標は揺れない）
  ② 製品側に入れるかもしれない検知器（「もう一度翻訳して突き合わせる」）が
     見るのと**同じもの**を測る ── 検知器の値打ちを、作る前に見積もれる。

★★ 陽性対照を必ず通す ── 「揺れ 0%」と出た時に、それが**本当に 0** なのか
  **測定器が揺れを見られない**のかを分けるため。

★★ 2026-09-19: 初版の陽性対照は**無効だった**。「金額で降順に並べ替えて」を温度 2.0 で
  振っても op は SORT 以外になりようがない ── 易しい依頼で対照を作ったので、温度が
  効いていても差が出なかった。道具は自分を信用せず exit 2 を返し、それは正しかった。
  ★ 切り分け: ollama を直叩き T=2.0 は毎回違う文を返す（温度は効く）。効かないのは
    `format:"json"` ＋ 強い few-shot でトークン分布が極端に尖るため。
  ★★ 正しい対照が見つかった ── 「原価の列をどうにかして」は **T=0.1 で生応答が 4/6
    相異なるのに op は 1 種**。だから対照は **2 層で見る**:
        生応答が揺れた   → 測定器は感度を持っている（サンプリングは動いている）
        op は揺れない    → ★ それは本物の性質であって、測定の失敗ではない
  ★ この 2 層を分けないと「0%」の意味が決まらない。

★ 分母の注意: n 回すべて同じでも、除外できるのは「揺れ率 > 3/n」まで
  （n=3 なら 100%、n=10 なら 30%、n=30 なら 10%）。**少ない n で「揺れない」と言わない。**

使い方:
    python bench/translation_wobble.py --runs 10
    python bench/translation_wobble.py --runs 10 --out wobble.json
"""
from __future__ import annotations

import _bench_home  # noqa: E402 ── ★ ailine より先に（本物の ~/.ailine に書かない）
_bench_home.isolate()

import argparse
import json
import subprocess
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import ailine  # noqa: E402

BATTERY = json.loads((ROOT / "bench" / "translation_battery.json")
                      .read_bytes().decode("utf-8"))

#: ★ 検体は 2 群に分ける（選び方が結果を作らないように）。
#:   A 群: battery の素の依頼 ── 実需に近い代表。**揺れたから選ぶ、をしない**。
#:   B 群: 履歴で実際に揺れた依頼 ── 陽性対照（測定器が揺れを見られるかの確認）。
#: ★ 冊は battery のものを使うので、履歴と**同じ冊ではない**（履歴は冊を残していない）。
#:   だから B 群の数字は「履歴の再現」ではなく「この冊での揺れ」── 比べる時は注意。
GROUP_B = [
    ("sample", "3行目を削除して"),                      # 履歴 283回: DELETE_ROWS 283 / INSERT_ROWS 1
    ("sample", "ナットの行を除いて"),                    # 履歴  76回: DELETE_ROWS 72 / EXTRACT 4
    ("sample", "商品と売上の列だけ抜き出して"),           # 履歴  16回: EXTRACT_COLUMNS 14 / PLAN 2
    ("sample", "みかんとぶどうの間に梨を追加して"),        # 履歴  30回: ADD_ROW 29 / INSERT_ROWS 1
]


#: ★ 陽性対照 ── **生応答が揺れると実測で分かっている**依頼（2026-09-19）。
#:   「原価の列をどうにかして」は温度 0.1 で生応答 4/6 が相異なるのに op は 1 種だった。
#:   ここが揺れなければ、サンプリングが動いていない（測定器の側を疑う）。
POSITIVE_CONTROL = [
    ("sample", "原価の列をどうにかして"),
    ("sample", "いい感じに整えて"),
]


def _books() -> dict:
    return BATTERY["_meta"]["books"]


def _meta_for(book_key) -> dict:
    """book_key が battery の名前なら battery の冊、dict ならそのまま冊として使う。

    ★ dict を受けるのは、履歴から復元した実依頼の冊を通すため
      （bench/rebuild_books_from_history.py が作る）。
    """
    if isinstance(book_key, dict):
        return {"sheets": [book_key["sheet"]],
                "headers": {book_key["sheet"]: list(book_key["headers"])}}
    headers = _books()[book_key]["sheets"]
    return {"sheets": list(headers.keys()), "headers": headers}


def _load_tasks(path: str) -> tuple:
    """復元した実依頼を (接地あり, 接地なし) の 2 群に分けて返す。

    ★ 混ぜて平均しない ── 接地の有無で翻訳の確信度が変わるので、別の群として読む。
    """
    items = json.loads(Path(path).read_bytes().decode("utf-8"))
    g = [({"sheet": it["sheet"], "headers": it["headers"]}, it["task"])
         for it in items if it.get("grounded")]
    ng = [({"sheet": it["sheet"], "headers": it["headers"]}, it["task"])
          for it in items if not it.get("grounded")]
    return g, ng


def _ops_of(result: dict) -> tuple:
    """計画を『op の並び』へ畳む（引数の揺れは別の話なので、まず op で見る）。"""
    return tuple(str((s or {}).get("op") or "?") for s in (result.get("plan") or []))


def _ops_from_raw(raw: str) -> tuple:
    """生応答から op の並びを取る。★ 製品の正規化は通さない。

    ★ 理由: translate_task の正規化は語彙外の op をまとめて FREEFORM に畳むので、
      **違う外し方どうしが同じ顔になり、揺れが隠れる**。ここは感度を優先して生で見る。
    ★ 解けない応答は ("?",) にして黙って捨てない（それも 1 つの結果）。
    """
    try:
        data = json.loads(raw)
    except Exception:
        return ("?",)
    steps = data.get("plan") if isinstance(data, dict) else None
    if not isinstance(steps, list):
        steps = [data] if isinstance(data, dict) else []
    return tuple(str((s or {}).get("op") or "?") for s in steps) or ("?",)


def _raw_of(model: str, task: str, meta: dict, temperature: float) -> str:
    """生の応答そのもの。★ 測定器に感度が在るかを見る層（op より手前）。"""
    return ailine.ollama_generate_json(
        model, ailine.build_translation_messages(task, meta),
        temperature=temperature, num_predict=700)


def _rule_of_three(n: int) -> float:
    """n 回すべて同じ時に除外できる揺れ率（%）。★ 少ない n で断定しないための目安。"""
    return 300.0 / n if n else 100.0


def measure(model: str, runs: int, temperature: float, items: list) -> list:
    out = []
    for book_key, text in items:
        meta = _meta_for(book_key)
        seen, raws = [], []
        t0 = time.time()
        for _ in range(runs):
            # ★ 生応答と op を**同じ 1 回の呼び出しから**取る（2 度振ると別の試行になる）。
            raw = _raw_of(model, text, meta, temperature)
            raws.append(raw)
            seen.append(_ops_from_raw(raw))
        c = Counter(seen)
        top = c.most_common(1)[0][1]
        out.append({"book": book_key, "task": text, "runs": runs,
                     "distribution": {" + ".join(k): v for k, v in c.most_common()},
                     "distinct": len(c), "minority": runs - top,
                     "raw_distinct": len(set(raws)),
                     "seconds": round(time.time() - t0, 1)})
        mark = "★ 揺れた" if len(c) > 1 else "安定  "
        print(f"  {mark}  {text[:32]:<32} op={dict(out[-1]['distribution'])} "
               f"生応答 {len(set(raws))}/{runs}")
    return out


def summarise(rows: list, label: str) -> dict:
    n_items = len(rows)
    wobbled = sum(1 for r in rows if r["distinct"] > 1)
    runs = sum(r["runs"] for r in rows)
    minority = sum(r["minority"] for r in rows)
    # ★ 2 回引いて食い違う確率（検知器の発火率の目安）
    fire = 0.0
    for r in rows:
        n = r["runs"]
        p = [v / n for v in r["distribution"].values()]
        fire += n * (1 - sum(x * x for x in p))
    s = {"label": label, "items": n_items, "wobbled": wobbled,
         "runs": runs, "minority": minority,
         "minority_pct": round(100 * minority / max(runs, 1), 2),
         "two_draw_disagree_pct": round(100 * fire / max(runs, 1), 2)}
    print(f"\n{label}: 依頼 {n_items} 種 / 走行 {runs}")
    print(f"  ★ 揺れた依頼      {wobbled} 種 ({100*wobbled/max(n_items,1):.1f}%)")
    print(f"  ★ 少数派を引いた回 {minority} ({s['minority_pct']}%)")
    print(f"     2 回引いて食い違う {s['two_draw_disagree_pct']}%（検知器の発火率の目安）")
    return s


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="翻訳の揺れを HEAD 固定で測る")
    ap.add_argument("--model", default="qwen2.5-coder:7b")
    ap.add_argument("--runs", type=int, default=10)
    ap.add_argument("--items", type=int, default=20, help="A 群から取る依頼数")
    ap.add_argument("--out")
    ap.add_argument("--tasks", help="履歴から復元した実依頼の JSON"
                                  "（bench/rebuild_books_from_history.py が作る）")
    a = ap.parse_args(argv)

    head = subprocess.run(["git", "log", "-1", "--format=%H %ad %s", "--date=short"],
                           cwd=str(ROOT), capture_output=True, text=True,
                           encoding="utf-8").stdout.strip()
    dirty = subprocess.run(["git", "status", "--porcelain"], cwd=str(ROOT),
                            capture_output=True, text=True, encoding="utf-8").stdout.strip()
    print(f"HEAD: {head}")
    if dirty:
        print("★ 作業木に未 commit の変更がある ── 何を測ったか後で言えなくなる:")
        print("  " + "\n  ".join(dirty.splitlines()[:8]))
    print(f"モデル: {a.model}  振る回数: {a.runs}")
    print(f"★ {a.runs} 回すべて同じでも、除外できるのは揺れ率 > "
          f"{_rule_of_three(a.runs):.0f}% まで\n")

    books = _books()
    group_a = [(it["book"], it["text"]) for it in BATTERY["items"]
               if it.get("book") in books][:a.items]

    print(f"── A 群（battery の素の依頼 {len(group_a)} 件・温度 0.1）──")
    rows_a = measure(a.model, a.runs, 0.1, group_a)
    sa = summarise(rows_a, "A 群")

    print(f"\n── B 群（履歴で揺れた依頼 {len(GROUP_B)} 件・温度 0.1）──")
    rows_b = measure(a.model, a.runs, 0.1, GROUP_B)
    sb = summarise(rows_b, "B 群")

    rows_c1 = rows_c2 = []
    sc1 = sc2 = None
    if a.tasks:
        # ★★ C 群 ── **実際に打たれた依頼**（battery の外・種を固定して無作為）。
        #   A 群（製品が想定した言い方）でも B 群（揺れたから選んだ）でもない、
        #   いちばん偏りの少ない標本。冊は解釈行から復元してある。
        g, ng = _load_tasks(a.tasks)
        print(f"\n── C1 群（実依頼・冊を復元できた {len(g)} 件・温度 0.1）──")
        rows_c1 = measure(a.model, a.runs, 0.1, g)
        sc1 = summarise(rows_c1, "C1 群（接地あり）")
        print(f"\n── C2 群（実依頼・冊を復元できず代替の冊 {len(ng)} 件・温度 0.1）──")
        print("   ★ 当時の冊ではない ── 接地が違うので C1 と混ぜて平均しない")
        rows_c2 = measure(a.model, a.runs, 0.1, ng)
        sc2 = summarise(rows_c2, "C2 群（代替の冊）")

    # ★★ 陽性対照は「生応答が揺れるか」で見る（op ではない）。
    #   2026-09-19: 初版は易しい依頼の op を温度 1.0 で振っており、op が動かないのは
    #   当たり前だった。感度は**サンプリングが動いているか**で確かめる。
    print("\n── 陽性対照（★ 生応答が揺れると分かっている依頼・温度 0.1）"
           "── ここで生応答が揺れなければ測定器を疑う ──")
    rows_c = measure(a.model, a.runs, 0.1, POSITIVE_CONTROL)
    sc = summarise(rows_c, "陽性対照")
    sensitive = any(r["raw_distinct"] > 1 for r in rows_c)
    if not sensitive:
        print("\n★★ 陽性対照の生応答が 1 件も揺れなかった ── **この測定は信用しない**。"
               "サンプリングが動いていないか、翻訳が実際には走っていない疑い。")
    else:
        print(f"\n★ 測定器は感度を持っている（陽性対照の生応答が揺れた）。"
               f"よって op が揺れないのは**本物の性質**であって測定の失敗ではない。")

    data = {"head": head, "dirty": bool(dirty), "model": a.model, "runs": a.runs,
            "rule_of_three_pct": _rule_of_three(a.runs),
            "groups": {"a": rows_a, "b": rows_b, "c1_grounded": rows_c1,
                        "c2_fallback": rows_c2, "positive_control": rows_c},
            "summary": [x for x in (sa, sb, sc1, sc2, sc) if x]}
    if a.out:
        Path(a.out).write_bytes(json.dumps(data, ensure_ascii=False, indent=1).encode("utf-8"))
        print(f"\n書き出し: {a.out}")
    # ★ 合否は「陽性対照の**生応答**が揺れたか」で決める（op ではない ── 上の理由）。
    return 0 if sensitive else 2


if __name__ == "__main__":
    raise SystemExit(main())
