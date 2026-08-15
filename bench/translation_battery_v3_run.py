# ★ A' battery v3: APPEND_TOTAL の倍率(factor)を LLM から切り離した後の通し実測。
#   抽出そのもの(extract_rate_factor/lookup_vocab_factor)は regex で決定論的なので
#   単体テスト(tests/test_ailine.py)が担保する。ここで測るのは LLM 依存の唯一の残り:
#   translate_task が返す label が、依頼文の「税込み/消費税」等の言い回しを
#   verify_dsl_args の税/込チェックが拾える形で保持しているか（＝真CLARIFY/誤CLARIFYの分岐）。
#   ★ items_v3 は translation_battery.json に凍結済み（改変しない）。
#   ★ 合わせて ollama 応答の prompt_eval_count を記録する（プロンプト肥大の正直な測定。
#     A' は OPS_DOC/few-shot から factor を削っただけ＝プロンプトは縮む方向のはずという
#     主張を目分量でなく数値で確認する）。
import json
import sys
import urllib.request
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))
import ailine  # noqa: E402

BATTERY = json.loads((HERE / "translation_battery.json").read_text(encoding="utf-8"))
MODEL = sys.argv[1] if len(sys.argv) > 1 else "qwen2.5-coder:7b"

books = dict(BATTERY["_meta"]["books"])
books.update(BATTERY["_meta"]["v2"]["books_extra"])
items = BATTERY["items_v3"]


def measure_prompt_eval_count(text: str, book_meta: dict) -> int | None:
    """代表1件でよい（system+few-shotは全項目でほぼ同じ長さ）。ollama_generate_json と
       同じ body 構成で直接叩き、通常は捨てられる prompt_eval_count だけを読む。"""
    messages = ailine.build_translation_messages(text, book_meta)
    body = {"model": MODEL, "messages": messages, "stream": False, "format": "json",
            "options": {"temperature": 0.1, "num_predict": 700, "num_ctx": 8192}}
    if "qwen3" in MODEL:
        body["think"] = False
    req = urllib.request.Request(f"{ailine.OLLAMA}/api/chat", data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        d = json.load(r)
    return d.get("prompt_eval_count")


true_clarify_n = true_clarify_ok = 0     # 率が本当に無い→聞くべき
false_clarify_n = false_clarify_ok = 0   # 率が明記/用語集にある→聞かずに確定すべき
per_item = []

for item in items:
    headers = books[item["book"]]["sheets"]
    book_meta = {"sheets": list(headers.keys()), "headers": headers}
    plan = ailine.translate_task(MODEL, item["text"], book_meta, temperature=0.1)
    steps = plan.get("plan") or [{"op": "FREEFORM", "args": {}}]
    # ★ 複合依頼（例:「小計に...入れて、税込み合計を...」）は APPEND_TOTAL が2段目以降に
    #   来る。plan 内のどこにあっても探す（この battery が測るのは複合解析の精度でなく、
    #   APPEND_TOTAL 段が来たときの factor 解決の精度＝v1/v2 の管轄と分離する）。
    step = next((s for s in steps if str(s.get("op", "")).upper() == "APPEND_TOTAL"), None)

    exp = item["expect"]
    if step is None:
        # LLM が APPEND_TOTAL 自体を外した（複合解析の精度は v1/v2 battery の管轄・
        #   ここでは「探した記録」だけ残して factor 軸の集計からは除く）。
        got_ops = [str(s.get("op", "")).upper() for s in steps]
        per_item.append((item["id"], f"× APPEND_TOTAL段なし(実行plan={got_ops})", None))
        continue

    ok_v, resolved, inferred, err = ailine.verify_dsl_args(
        "APPEND_TOTAL", step.get("args", {}), book_meta, task=item["text"], vocab=item.get("vocab") or {})

    if exp.get("clarify"):
        true_clarify_n += 1
        got_clarify = not ok_v
        if got_clarify:
            true_clarify_ok += 1
        mark = "✓" if got_clarify else "×"
        outcome = "CLARIFY" if not ok_v else f"確定(factor={resolved.get('factor')})"
        per_item.append((item["id"], f"{mark} 真CLARIFY期待 → {outcome}", got_clarify))
    else:
        # ★ 誤CLARIFY(倍率不明)＝税/込ラベルの CLARIFY 番人が「率がある」のに誤って発火した
        #   ケースだけを分母に数える。列名不在等（複合解析でLLMが余計な段を挟んだ副作用）は
        #   factor 軸と無関係の別問題として集計から除外し、明細にだけ残す（目分量で
        #   誤CLARIFYの分母に混ぜない）。
        if ok_v:
            false_clarify_n += 1
            got_no_clarify = abs(resolved.get("factor", -1) - exp["factor"]) < 1e-9
            if got_no_clarify:
                false_clarify_ok += 1
            per_item.append((item["id"],
                             f"{'✓' if got_no_clarify else '×'} factor={resolved.get('factor')}"
                             f"(期待{exp['factor']})", got_no_clarify))
        elif "倍率が分かりません" in (err or ""):
            false_clarify_n += 1   # 番人が誤発火＝誤CLARIFY そのもの
            per_item.append((item["id"], "× 誤CLARIFY(倍率不明の番人が誤発火)", False))
        else:
            per_item.append((item["id"], f"△ 別要因のfail(集計対象外): {err}", None))

print(f"model: {MODEL}  (ailine.translate_task + verify_dsl_args・通し実測・items_v3)")
print(f"真CLARIFY(率が無い時に聞く): {true_clarify_ok}/{true_clarify_n}"
      f"{'' if true_clarify_n == 0 else f' = {true_clarify_ok/true_clarify_n:.1%}'}  (合格線 90%+)")
print(f"誤CLARIFYなし(率がある時に確定): {false_clarify_ok}/{false_clarify_n}"
      f"{'' if false_clarify_n == 0 else f' = {false_clarify_ok/false_clarify_n:.1%}'}  (合格線 90%+ ＝誤CLARIFY10%以下)")
print("\n-- 項目別:")
for iid, msg, ok in per_item:
    print(f"  #{iid}: {msg}")

try:
    pec = measure_prompt_eval_count(items[0]["text"], books[items[0]["book"]])
    print(f"\nprompt_eval_count（代表1件・#{items[0]['id']}）: {pec}")
except Exception as e:
    print(f"\nprompt_eval_count 計測に失敗: {e}")
