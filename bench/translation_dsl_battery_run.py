# M2b battery 再実行: bench/translation_spike.py と同じ battery（凍結済み・改変しない）・
# 同じ採点基準（score_slots/norm）で、本実装 ailine.translate_task を直接叩く。
# ★ 一段翻訳スパイク(translation_spike.py)とはプロンプトが違う（本実装は few-shot 5例つき・
#   nested {"op","args"} 形式）。凍結合格線（op90%/slot80%/誤断定20%）を本実装のプロンプトで
#   再達成していることを確かめるのが目的。
# ★ M2c: translate_task は常に {"plan": [...]} を返すようになった（複合依頼対応）。
#   items(v1) は単一依頼のみなので、plan[0] を旧来どおりの単一 op 形式として採点する
#   （後方互換: 単一依頼は長さ1の計画）。
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))   # ailine.py
sys.path.insert(0, str(HERE))          # translation_spike.py (BATTERY/score_slots の再利用)

import ailine
from translation_spike import BATTERY, score_slots

MODEL = sys.argv[1] if len(sys.argv) > 1 else "qwen2.5-coder:7b"

op_ok = op_n = slot_ok = slot_n = 0
misassert = amb_n = 0
fails = []

for item in BATTERY["items"]:
    headers = BATTERY["_meta"]["books"][item["book"]]["sheets"]
    book_meta = {"sheets": list(headers.keys()), "headers": headers}
    plan_result = ailine.translate_task(MODEL, item["text"], book_meta, temperature=0.1)
    got = (plan_result.get("plan") or [{"op": "FREEFORM", "args": {}}])[0]
    got_op = str(got.get("op", "")).upper()
    if got_op == "OUT_OF_VOCAB":
        # ★ M2c: プロンプトが FREEFORM → OUT_OF_VOCAB に切り替わった（意味は同じ・語彙外）。
        #   v1 battery の expect ラベル("freeform")と揃えて採点する（別ラベル扱いにしない）。
        got_op = "FREEFORM"
    exp = item["expect"]

    if exp in ("clarify", "freeform"):
        amb_n += 1
        op_n += 1
        ok = got_op == exp.upper() or (exp == "clarify" and got_op == "FREEFORM")
        if got_op not in ("CLARIFY", "FREEFORM"):
            misassert += 1
            fails.append((item["id"], f"誤断定: {got_op} ← 期待 {exp}"))
        elif ok:
            op_ok += 1
        else:
            fails.append((item["id"], f"{got_op} ← 期待 {exp} (許容内だが不一致)"))
            op_ok += 1
    else:
        op_n += 1
        if got_op == exp["op"]:
            op_ok += 1
            h, t = score_slots(exp, got)
            slot_ok += h
            slot_n += t
            if h < t:
                fails.append((item["id"], f"slot {h}/{t}: {got}"))
        else:
            fails.append((item["id"], f"op {got_op} ← 期待 {exp['op']}"))

print(f"model: {MODEL}  (ailine.translate_task 経由・本番プロンプト・few-shot 3例つき)")
print(f"op 分類: {op_ok}/{op_n} = {op_ok/op_n:.1%}  (合格線 90%)")
print(f"必須 slot: {slot_ok}/{slot_n} = {slot_ok/max(slot_n,1):.1%}  (合格線 80%)")
print(f"曖昧への誤断定: {misassert}/{amb_n}  (合格線 20% 以下)")
print("\n-- 不一致の明細:")
for fid, msg in fails:
    print(f"  #{fid}: {msg}")
