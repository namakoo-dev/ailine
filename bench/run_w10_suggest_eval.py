"""W10 便C2: 凍結セット bench/w10_suggest_frozen_set.json での suggest_ops 実測。

★ このスクリプトは凍結セット（検体・bars）を一切変更しない（読むだけ）。ここでやるのは
測定と件数の報告だけ ── 何かに合わせて suggest_ops や match_phrases を調整する場所ではない
（それは ailine.py 側の仕事。ここで見つけた不足は「発火条件」として報告するだけ）。

★ REGRESSION_FLOOR（回帰の床）だけは凍結セットではなくこのスクリプト本体が持つ
（Namakoo 決裁 2026-08-22）。便C1 の veto 名簿が凍結セットの true_out_of_vocab
10件に過適合し、封印されていた別の12件で 5/12 誤提示が出た事故の再発防止用。
この12件の個別語を veto 名簿や match_phrases に書き足すことは禁止 ── 一般則
（語としての厳格一致）の効果だけで下げる。

出力（すべて件数。n が小さいので比率で語らない）:
  ① 陽性対照        recall@1（バー: 18/18）
  ② ノイズ床        候補提示（バー: 0/10）
  ③ in_vocab        recall@1/@3（about なし版・内部診断・バー無し）
  ④ true_out_of_vocab 候補提示（バー: 2/10 以下）
  ⑤ slot_missing / adversarial_compound の挙動観察（合否ではなく実際の出力を記録）
  ⑥ REGRESSION_FLOOR 候補提示（バー: 1/12 以下）
  + about あり版（in_vocab・ollama 到達時のみ・落ちていれば明記して省略）

使い方: python bench/run_w10_suggest_eval.py [model]
"""
import json
import sys
from pathlib import Path

HERE = Path(__file__).parent
AILINE_DIR = HERE.parent
sys.path.insert(0, str(AILINE_DIR))
import ailine  # noqa: E402

FROZEN = json.loads((HERE / "w10_suggest_frozen_set.json").read_text(encoding="utf-8"))

# ★ 回帰の床(便C2・2026-08-22): 封印して抜き打ちで開けた12件。凍結セットではなく
#   このスクリプトが直接持つ（Namakoo の指示どおり ── 凍結 JSON の構造は増やさない）。
#   ここの文言を見て veto 名簿や match_phrases に個別語を足すのは禁止（過適合の再発）。
REGRESSION_FLOOR = [
    {"id": "RF01", "text": "ヘッダーとフッターにページ番号を入れて"},
    {"id": "RF02", "text": "シートごとに別のブックに分割して保存して"},
    {"id": "RF03", "text": "入力規則で日付しか入れられないようにして"},
    {"id": "RF04", "text": "商品名にリンクを一括で貼って"},
    {"id": "RF05", "text": "2行目までスクロールしても常に見えるようにして"},
    {"id": "RF06", "text": "テーブルとして書式設定して"},
    {"id": "RF07", "text": "売上の横にスパークラインを付けて"},
    {"id": "RF08", "text": "行の高さを全部そろえて"},
    {"id": "RF09", "text": "重複してるセルに色を付けて"},
    {"id": "RF10", "text": "このシートをPDFにして"},
    {"id": "RF11", "text": "変更履歴が残るようにして"},
    {"id": "RF12", "text": "セルを斜線で消して"},
]

# ★ about あり版で translate_task に渡す book_meta（凍結セットは書名/シート構成を宣言
#   していないため、in_vocab の文言に登場する列名を広く含む汎用ブックで代用する。
#   about の中身の精度はここでは重要でない ── 「about を候補生成に足すと拾えるか」を
#   ollama が生きていれば best-effort で覗くだけの副次測定）。
GENERIC_BOOK_META = {
    "sheets": ["Sheet"],
    "headers": {"Sheet": ["日付", "商品コード", "商品", "得意先", "部門", "支店",
                            "数量", "単価", "金額", "税抜", "税込", "ステータス", "備考"]},
}


def _match_phrase_pool_overlap() -> None:
    """透明化: match_phrases の初期語彙が凍結セットの文言とどれだけ重なるかを報告する
       （重ねて書いていないかの事後点検・自己汚染チェック）。完全一致の部分文字列だけを見る。"""
    all_texts = []
    for key in ("positive_control", "noise_floor", "in_vocab", "true_out_of_vocab", "slot_missing"):
        for e in FROZEN[key]:
            all_texts.append(e["text"])
    for e in FROZEN["adversarial_compound"]:
        all_texts.append(e["text"])
    phrases = []
    for meta in ailine.OP_META.values():
        phrases.extend(meta.get("match_phrases", ()))
    overlap = [p for p in phrases if any(p in t for t in all_texts)]
    print(f"[透明化] match_phrases 総数: {len(phrases)} / 凍結セットの文言に部分文字列として"
          f"現れるもの: {len(overlap)} 件 ({len(overlap) / len(phrases) * 100:.1f}%)")
    if overlap:
        print(f"        重複語: {overlap}")


def run_positive_control() -> int:
    hits = 0
    misses = []
    for e in FROZEN["positive_control"]:
        cands = ailine.suggest_ops(e["text"])
        if cands and cands[0] == e["expect_op"]:
            hits += 1
        else:
            misses.append((e["id"], e["text"], e["expect_op"], cands))
    print(f"① 陽性対照 recall@1: {hits}/18（バー 18/18）")
    for m in misses:
        print(f"    miss: {m}")
    return hits


def run_noise_floor() -> int:
    fired = []
    for e in FROZEN["noise_floor"]:
        cands = ailine.suggest_ops(e["text"])
        if cands:
            fired.append((e["id"], e["text"], cands))
    print(f"② ノイズ床 候補提示: {len(fired)}/10（バー 0/10）")
    for f in fired:
        print(f"    fired: {f}")
    return len(fired)


def run_in_vocab_no_about() -> tuple:
    items = FROZEN["in_vocab"]
    hit1 = hit3 = 0
    for e in items:
        cands = ailine.suggest_ops(e["text"]) or []
        if cands and cands[0] == e["expect_op"]:
            hit1 += 1
        if e["expect_op"] in cands:
            hit3 += 1
    n = len(items)
    print(f"③ in_vocab（about なし・内部診断）: recall@1 = {hit1}/{n} / recall@3 = {hit3}/{n}")
    return hit1, hit3, n


def run_true_out_of_vocab() -> int:
    items = FROZEN["true_out_of_vocab"]
    fired = []
    for e in items:
        cands = ailine.suggest_ops(e["text"]) or []
        if cands:
            fired.append((e["id"], e["text"], cands))
    print(f"④ true_out_of_vocab 候補提示: {len(fired)}/{len(items)}（バー 2/10 以下）")
    for f in fired:
        print(f"    fired: {f}")
    return len(fired)


def run_slot_missing_and_adversarial() -> None:
    print("⑤ slot_missing の挙動観察（expect_op が候補に出るか。CLARIFY へ回るのは別経路）:")
    hit = 0
    for e in FROZEN["slot_missing"]:
        cands = ailine.suggest_ops(e["text"]) or []
        ok = e["expect_op"] in cands
        hit += ok
        print(f"    {e['id']} 「{e['text']}」→ {cands}"
              f"（expect {e['expect_op']}: {'含む' if ok else '含まない'}）")
    print(f"    expect_op が候補に含まれた件数: {hit}/{len(FROZEN['slot_missing'])}")

    print("⑤ adversarial_compound の挙動観察（合否でなく実際の出力を記録）:")
    for e in FROZEN["adversarial_compound"]:
        cands = ailine.suggest_ops(e["text"]) or []
        print(f"    {e['id']} 「{e['text']}」→ {cands}  ({e['expect']})")


def run_regression_floor() -> int:
    fired = []
    for e in REGRESSION_FLOOR:
        cands = ailine.suggest_ops(e["text"]) or []
        if cands:
            fired.append((e["id"], e["text"], cands))
    print(f"⑥ REGRESSION_FLOOR 候補提示: {len(fired)}/{len(REGRESSION_FLOOR)}（バー 1/12 以下）")
    for f in fired:
        print(f"    fired: {f}")
    return len(fired)


def run_about_augmented(model: str) -> None:
    """about あり版は、about が実際に取れた検体**だけ**を対象に before/after を比べる
       （★ 直し: 旧版は before 側を 44 件全体の累計・after 側だけ分母を揃え忘れていて、
       誤読を招く数字だった。ここでは同じ部分集合で両方を測る）。"""
    reachable, msg = ailine._check_ollama_reachable()
    if not reachable:
        print(f"[about あり版] ollama 未到達につき省略（about なしのみ）: {msg}")
        return
    items = FROZEN["in_vocab"]
    subset = []   # about が取れた検体だけ (expect_op, cands_plain, cands_about)
    for e in items:
        text = e["text"]
        plan = ailine.translate_task(model, text, GENERIC_BOOK_META, temperature=0.1)
        step = (plan.get("plan") or [{}])[0]
        about = step.get("about") if step.get("op") == "OUT_OF_VOCAB" else None
        if not about:
            continue
        cands_plain = ailine.suggest_ops(text) or []
        cands_about = ailine.suggest_ops(text, about=about) or []
        subset.append((e["id"], e["expect_op"], cands_plain, cands_about))
    print(f"[about あり版] in_vocab {len(items)} 件中 translate_task が about を返したのは "
          f"{len(subset)} 件")
    if not subset:
        print("    about を返した検体が無かったため about あり版の比較はできない")
        return
    n = len(subset)
    hit1_plain = sum(1 for _, exp, cp, _ in subset if cp and cp[0] == exp)
    hit3_plain = sum(1 for _, exp, cp, _ in subset if exp in cp)
    hit1_about = sum(1 for _, exp, _, ca in subset if ca and ca[0] == exp)
    hit3_about = sum(1 for _, exp, _, ca in subset if exp in ca)
    print(f"    同じ {n} 件での before/after（about なし → about あり）:"
          f" recall@1 {hit1_plain}/{n} → {hit1_about}/{n} /"
          f" recall@3 {hit3_plain}/{n} → {hit3_about}/{n}")
    for id_, exp, cp, ca in subset:
        print(f"      {id_} expect={exp} なし={cp} あり={ca}")


def main() -> int:
    model = sys.argv[1] if len(sys.argv) > 1 else ailine.DEFAULT_MODEL
    print("=== W10 便C2 凍結セット実測（bench/w10_suggest_frozen_set.json + REGRESSION_FLOOR） ===")
    _match_phrase_pool_overlap()
    print()
    run_positive_control()
    print()
    run_noise_floor()
    print()
    run_in_vocab_no_about()
    print()
    run_true_out_of_vocab()
    print()
    run_slot_missing_and_adversarial()
    print()
    run_regression_floor()
    print()
    run_about_augmented(model)
    return 0


if __name__ == "__main__":
    sys.exit(main())
