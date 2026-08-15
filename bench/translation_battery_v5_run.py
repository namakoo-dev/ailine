# ★ W10b battery v5: 税込み/税抜き（COMPUTE_COLUMN の「1列 × 率」パターン）を DSL 語彙に
#   収載した後の翻訳精度を、ailine.translate_task(本番プロンプト)に通して測る。
#   items_v5 は translation_battery.json に凍結済み（改変しない）。単一依頼のみ（v1/v4 と同じ形）。
#   合格線: op 分類成功 5/6=83.3%以上（_meta.v5.bar）。
#   ★ A' 原則: 倍率(factor)は LLM に確定させない。ここでは分類精度に加え、
#   verify_dsl_args が machine 抽出(extract_rate_factor/lookup_vocab_factor)だけで
#   factor を確定できたか（真に率が明記された全件で CLARIFY に落ちていないか）も測る
#   （全項目が依頼文に率を明記しているため、凍結線は全件確定=factor_machine_determined 100%）。
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))   # ailine.py
sys.path.insert(0, str(HERE))          # translation_spike.py (BATTERY/score_slots の再利用)

import ailine
from translation_spike import BATTERY, score_slots

MODEL = sys.argv[1] if len(sys.argv) > 1 else "qwen2.5-coder:7b"

books = dict(BATTERY["_meta"]["books"])
books.update(BATTERY["_meta"]["v2"]["books_extra"])

op_ok = op_n = slot_ok = slot_n = 0
factor_ok = factor_n = 0
fails = []

for item in BATTERY["items_v5"]:
    headers = books[item["book"]]["sheets"]
    book_meta = {"sheets": list(headers.keys()), "headers": headers}
    plan_result = ailine.translate_task(MODEL, item["text"], book_meta, temperature=0.1)
    got = (plan_result.get("plan") or [{"op": "FREEFORM", "args": {}}])[0]
    got_op = str(got.get("op", "")).upper()
    exp = item["expect"]

    op_n += 1
    if got_op == exp["op"]:
        op_ok += 1
        h, t = score_slots(exp, got)
        slot_ok += h
        slot_n += t
        if h < t:
            fails.append((item["id"], f"slot {h}/{t}: {got}"))

        # ★ A': factor は machine 確定（LLM から受け取らない）。verify_dsl_args を通しで
        #   確認する（分類が当たった項目だけが対象＝分類自体の精度は上のop分類で見る）。
        args = got.get("args", {})
        ok_v, resolved, inferred, err = ailine.verify_dsl_args(
            "COMPUTE_COLUMN", args, book_meta, task=item["text"], vocab={})
        factor_n += 1
        if ok_v and resolved.get("_sources", {}).get("factor"):
            factor_ok += 1
        else:
            fails.append((item["id"], f"factor 未確定: ok={ok_v} resolved={resolved} err={err}"))
    else:
        fails.append((item["id"], f"op {got_op} ← 期待 {exp['op']} (got={got})"))

print(f"model: {MODEL}  (ailine.translate_task 経由・本番プロンプト・W10b items_v5)")
print(f"op 分類: {op_ok}/{op_n} = {op_ok/op_n:.1%}  (合格線 5/6=83.3%)")
print(f"必須 slot: {slot_ok}/{max(slot_n,1)} = {slot_ok/max(slot_n,1):.1%}")
print(f"factor 機械確定: {factor_ok}/{max(factor_n,1)} = {factor_ok/max(factor_n,1):.1%}  (合格線 100%＝LLM確定0件)")
print("\n-- 不一致の明細:")
for fid, msg in fails:
    print(f"  #{fid}: {msg}")
