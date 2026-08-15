# ★ W9 battery v4: 検証済みヘルパ4種(INSERT_ROWS/DRAW_BORDERS/AUTOFIT/PIVOT)を
#   DSL 語彙に昇格した後の翻訳精度を、ailine.translate_task(本番プロンプト)に通して測る。
#   items_v4 は translation_battery.json に凍結済み（改変しない）。単一依頼のみ（v1 と同じ形）。
#   合格線は v1 と同じ op90%/slot80%（_meta.v4.bar）。
#   ★ PIVOT/AGGREGATE の分岐（「ピボット」と明示された時だけ PIVOT）検体を含む
#   （#408/#409=PIVOT期待、#410/#411=AGGREGATE期待）。
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
fails = []

for item in BATTERY["items_v4"]:
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
    else:
        fails.append((item["id"], f"op {got_op} ← 期待 {exp['op']} (got={got})"))

print(f"model: {MODEL}  (ailine.translate_task 経由・本番プロンプト・W9 items_v4)")
print(f"op 分類: {op_ok}/{op_n} = {op_ok/op_n:.1%}  (合格線 90%)")
print(f"必須 slot: {slot_ok}/{max(slot_n,1)} = {slot_ok/max(slot_n,1):.1%}  (合格線 80%)")
print("\n-- 不一致の明細:")
for fid, msg in fails:
    print(f"  #{fid}: {msg}")
