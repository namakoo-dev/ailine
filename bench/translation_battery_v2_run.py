# M2c battery v2 実測: 複合依頼(items_v2)を ailine.translate_task(本番プロンプト)に通し、
#   ★ 計画完全性(plan_completeness)・黙落(silent_drop)・target 反映 を測る。
#   battery は実行前凍結（translation_battery.json の _meta.v2.frozen）・改変しない。
#
# 採点方針（位置対応・保守的）:
#   期待計画 expect_plan[i] に対応する実行計画 got_plan[i] が「存在する」だけで
#   plan_completeness の分子に数える（op が正確に一致していなくても『黙って落として
#   いない』ことの方を先に測る）。実行計画の長さが期待より短い分だけ silent_drop。
#   OUT_OF_VOCAB を期待する段は、実行側が OUT_OF_VOCAB/FREEFORM のどちらでも黙落なし扱い
#   （語彙外だと認識できたこと自体を評価する。厳密な op 一致率は参考値として別掲）。
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))   # ailine.py

import ailine

BATTERY = __import__("json").loads((HERE / "translation_battery.json").read_text(encoding="utf-8"))
MODEL = sys.argv[1] if len(sys.argv) > 1 else "qwen2.5-coder:7b"

books = dict(BATTERY["_meta"]["books"])
books.update(BATTERY["_meta"]["v2"]["books_extra"])

items = BATTERY["items_v2"]

total_expected_steps = 0
present_steps = 0          # 位置的に対応する実行段が存在した数（plan_completeness の分子）
silent_dropped = 0         # 対応する実行段が存在しなかった数（目標 0）
op_exact_hits = 0          # 参考値: op も一致した数
target_checks = 0
target_hits = 0
per_item = []

for item in items:
    headers = books[item["book"]]["sheets"]
    book_meta = {"sheets": list(headers.keys()), "headers": headers}
    result = ailine.translate_task(MODEL, item["text"], book_meta, temperature=0.1)
    got_plan = result.get("plan") or []
    exp_plan = item["expect_plan"]

    n_exp = len(exp_plan)
    n_got = len(got_plan)
    total_expected_steps += n_exp
    item_present = min(n_exp, n_got)
    item_dropped = max(n_exp - n_got, 0)
    present_steps += item_present
    silent_dropped += item_dropped

    detail = []
    for i, exp_step in enumerate(exp_plan):
        exp_op = exp_step["op"]
        if i >= n_got:
            detail.append(f"  段{i+1}: 黙落（期待 {exp_op} に対応する実行段が無い）")
            continue
        got_step = got_plan[i]
        got_op = str(got_step.get("op", "")).upper()
        if exp_op == "OUT_OF_VOCAB":
            ok = got_op in ("OUT_OF_VOCAB", "FREEFORM")
        else:
            ok = got_op == exp_op
        if ok:
            op_exact_hits += 1
        else:
            detail.append(f"  段{i+1}: op 不一致 実行={got_op} 期待={exp_op}")
        if exp_op == "COMPUTE_COLUMN" and "target" in exp_step:
            # ★ M2c の target(名指し列への書き込み)反映チェックは COMPUTE_COLUMN だけを見る。
            #   BOLD 等の "target" は無関係の別スロット（row:N/col:X の対象指定）なので混同しない。
            target_checks += 1
            got_target = (got_step.get("args") or {}).get("target")
            if got_target == exp_step["target"]:
                target_hits += 1
            else:
                detail.append(f"  段{i+1}: target 不一致 実行={got_target!r} 期待={exp_step['target']!r}")

    per_item.append((item["id"], n_exp, n_got, item_dropped, detail))

print(f"model: {MODEL}  (ailine.translate_task 経由・本番プロンプト・M2c 複合計画)")
print(f"計画完全性(黙落なし率): {present_steps}/{total_expected_steps} = "
      f"{present_steps/max(total_expected_steps,1):.1%}  (合格線 90%)")
print(f"黙落(部分意図が計画に一切現れない件数): {silent_dropped}  (目標 0)")
print(f"参考: op 完全一致: {op_exact_hits}/{total_expected_steps} = "
      f"{op_exact_hits/max(total_expected_steps,1):.1%}")
print(f"target 反映: {target_hits}/{target_checks}"
      f"{'' if target_checks == 0 else f' = {target_hits/target_checks:.1%}'}")
print("\n-- 項目別:")
for iid, n_exp, n_got, dropped, detail in per_item:
    mark = "✓" if dropped == 0 else "×"
    print(f"  #{iid}: {mark} 期待{n_exp}段 実行{n_got}段" + (f" 黙落{dropped}" if dropped else ""))
    for ln in detail:
        print(ln)
