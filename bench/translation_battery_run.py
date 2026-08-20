"""★ C9: M2b/A'/W9/W10b battery v2〜v5 の統合ランナー。
   旧 translation_battery_v2_run.py 〜 translation_battery_v5_run.py（4本の近似複製）を、
   共通の骨格（BATTERY 読み込み・books 構築）と battery ごとの採点関数へ分けて1本に畳んだ。
   ★★ 測定器は変えていない: 各 run_v*() の中身は旧スクリプトの本体を関数へそのまま移した
   だけで、翻訳呼び出し（ailine.translate_task への実引数・呼び出し順序・item の走査順）・
   採点ロジック・出力文言は一字一句変えていない（LLM 出力は毎回揺れるため、結果の比較では
   測定器が同一である証明にならない ── ailine.ollama_generate_json/urllib.request.urlopen を
   monkeypatch して「LLM へ実際に送る body（プロンプト・検体・パラメータ）」を dump し、
   旧4本と新1本とで byte 一致を確認済み。詳細は C9 の作業報告参照）。

   使い方: python translation_battery_run.py <v2|v3|v4|v5> [model]
   （旧: python translation_battery_vN_run.py [model] → 新: 第1引数に battery id を追加するだけ）。
"""
import json
import sys
import urllib.request
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))   # ailine.py
sys.path.insert(0, str(HERE))          # translation_spike.py (BATTERY/score_slots の再利用)

import ailine  # noqa: E402
from translation_spike import BATTERY, score_slots  # noqa: E402


def _books() -> dict:
    books = dict(BATTERY["_meta"]["books"])
    books.update(BATTERY["_meta"]["v2"]["books_extra"])
    return books


# ===========================================================================
# v2 — M2c battery v2: 複合依頼(items_v2)の plan_completeness / silent_drop / target 反映
# ===========================================================================
def run_v2(model: str) -> None:
    """★ 元 translation_battery_v2_run.py 本体（採点方針はファイル冒頭コメント参照）。"""
    books = _books()
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
        result = ailine.translate_task(model, item["text"], book_meta, temperature=0.1)
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

    print(f"model: {model}  (ailine.translate_task 経由・本番プロンプト・M2c 複合計画)")
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


# ===========================================================================
# v3 — A' battery v3: APPEND_TOTAL の factor 解決（真/誤 CLARIFY）+ prompt_eval_count
# ===========================================================================
def _measure_prompt_eval_count(model: str, text: str, book_meta: dict) -> int | None:
    """代表1件でよい（system+few-shotは全項目でほぼ同じ長さ）。ollama_generate_json と
       同じ body 構成で直接叩き、通常は捨てられる prompt_eval_count だけを読む。"""
    messages = ailine.build_translation_messages(text, book_meta)
    body = {"model": model, "messages": messages, "stream": False, "format": "json",
            "options": {"temperature": 0.1, "num_predict": 700, "num_ctx": 8192}}
    if "qwen3" in model:
        body["think"] = False
    req = urllib.request.Request(f"{ailine.OLLAMA}/api/chat", data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        d = json.load(r)
    return d.get("prompt_eval_count")


def run_v3(model: str) -> None:
    """★ 元 translation_battery_v3_run.py 本体（採点方針はファイル冒頭コメント参照）。"""
    books = _books()
    items = BATTERY["items_v3"]

    true_clarify_n = true_clarify_ok = 0     # 率が本当に無い→聞くべき
    false_clarify_n = false_clarify_ok = 0   # 率が明記/用語集にある→聞かずに確定すべき
    per_item = []

    for item in items:
        headers = books[item["book"]]["sheets"]
        book_meta = {"sheets": list(headers.keys()), "headers": headers}
        plan = ailine.translate_task(model, item["text"], book_meta, temperature=0.1)
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

    print(f"model: {model}  (ailine.translate_task + verify_dsl_args・通し実測・items_v3)")
    print(f"真CLARIFY(率が無い時に聞く): {true_clarify_ok}/{true_clarify_n}"
          f"{'' if true_clarify_n == 0 else f' = {true_clarify_ok/true_clarify_n:.1%}'}  (合格線 90%+)")
    print(f"誤CLARIFYなし(率がある時に確定): {false_clarify_ok}/{false_clarify_n}"
          f"{'' if false_clarify_n == 0 else f' = {false_clarify_ok/false_clarify_n:.1%}'}  (合格線 90%+ ＝誤CLARIFY10%以下)")
    print("\n-- 項目別:")
    for iid, msg, ok in per_item:
        print(f"  #{iid}: {msg}")

    try:
        pec = _measure_prompt_eval_count(model, items[0]["text"], books[items[0]["book"]])
        print(f"\nprompt_eval_count（代表1件・#{items[0]['id']}）: {pec}")
    except Exception as e:
        print(f"\nprompt_eval_count 計測に失敗: {e}")


# ===========================================================================
# v4 — W9 battery v4: 検証済みヘルパ4種を語彙昇格した後の翻訳精度（単一依頼）
# ===========================================================================
def run_v4(model: str) -> None:
    """★ 元 translation_battery_v4_run.py 本体（採点方針はファイル冒頭コメント参照）。"""
    books = _books()

    op_ok = op_n = slot_ok = slot_n = 0
    fails = []

    for item in BATTERY["items_v4"]:
        headers = books[item["book"]]["sheets"]
        book_meta = {"sheets": list(headers.keys()), "headers": headers}
        plan_result = ailine.translate_task(model, item["text"], book_meta, temperature=0.1)
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

    print(f"model: {model}  (ailine.translate_task 経由・本番プロンプト・W9 items_v4)")
    print(f"op 分類: {op_ok}/{op_n} = {op_ok/op_n:.1%}  (合格線 90%)")
    print(f"必須 slot: {slot_ok}/{max(slot_n,1)} = {slot_ok/max(slot_n,1):.1%}  (合格線 80%)")
    print("\n-- 不一致の明細:")
    for fid, msg in fails:
        print(f"  #{fid}: {msg}")


# ===========================================================================
# v5 — W10b battery v5: 税込み/税抜きの op 分類 + factor 機械確定
# ===========================================================================
def run_v5(model: str) -> None:
    """★ 元 translation_battery_v5_run.py 本体（採点方針はファイル冒頭コメント参照）。"""
    books = _books()

    op_ok = op_n = slot_ok = slot_n = 0
    factor_ok = factor_n = 0
    fails = []

    for item in BATTERY["items_v5"]:
        headers = books[item["book"]]["sheets"]
        book_meta = {"sheets": list(headers.keys()), "headers": headers}
        plan_result = ailine.translate_task(model, item["text"], book_meta, temperature=0.1)
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

    print(f"model: {model}  (ailine.translate_task 経由・本番プロンプト・W10b items_v5)")
    print(f"op 分類: {op_ok}/{op_n} = {op_ok/op_n:.1%}  (合格線 5/6=83.3%)")
    print(f"必須 slot: {slot_ok}/{max(slot_n,1)} = {slot_ok/max(slot_n,1):.1%}")
    print(f"factor 機械確定: {factor_ok}/{max(factor_n,1)} = {factor_ok/max(factor_n,1):.1%}  (合格線 100%＝LLM確定0件)")
    print("\n-- 不一致の明細:")
    for fid, msg in fails:
        print(f"  #{fid}: {msg}")


# ===========================================================================
# v6 — EXTRACT battery: 単一条件抽出の op 分類（items_v6・薄いランナー・W10 系実測時点で
#   未配線だったものをここで配線する。_meta.v6 の凍結バー宣言は無いため、v1/v4 と同じ
#   合格線 op90%/slot80% を暫定適用する（items_v6 のコミットメッセージが「battery 2件も
#   構造カウント済み」と述べる op 完全性番人の要件を、実行して確かめる薄いランナー）。
# ===========================================================================
def run_v6(model: str) -> None:
    """★ v4/v5 と同じ単一依頼スコアリングを EXTRACT (items_v6) に適用するだけの薄いランナー。
       翻訳呼び出し・採点方式(score_slots)は v4/v5 と揃える。"""
    books = _books()

    op_ok = op_n = slot_ok = slot_n = 0
    fails = []

    for item in BATTERY["items_v6"]:
        headers = books[item["book"]]["sheets"]
        book_meta = {"sheets": list(headers.keys()), "headers": headers}
        plan_result = ailine.translate_task(model, item["text"], book_meta, temperature=0.1)
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

    print(f"model: {model}  (ailine.translate_task 経由・本番プロンプト・EXTRACT items_v6・暫定バー"
          " op90%/slot80%＝_meta.v6 未宣言のため v1/v4 準拠)")
    print(f"op 分類: {op_ok}/{op_n} = {op_ok/op_n:.1%}  (合格線 90%)")
    print(f"必須 slot: {slot_ok}/{max(slot_n,1)} = {slot_ok/max(slot_n,1):.1%}  (合格線 80%)")
    print("\n-- 不一致の明細:")
    for fid, msg in fails:
        print(f"  #{fid}: {msg}")


# ===========================================================================
# v7 — C9 実測3穴 battery: COMPUTE_COLUMN target / SET_COLUMN_VALUE / APPEND_TOTAL
#   税込み系言い回し。事前登録は bench/PREREG_translation_v7.md 参照。
# ===========================================================================
def run_v7(model: str) -> None:
    """items_v7 は kind で3種に分かれ、採点方式もそれぞれ既存 battery の方式を流用する
       （target→run_v2 の target 検査と同じ考え方・scv→run_v4 と同じ op/slot・
       total_factor→run_v3 と同じ verify_dsl_args 通し）。n が極小のため件数もそのまま出す。"""
    books = _books()
    items = BATTERY["items_v7"]

    target_checks = target_hits = 0
    scv_op_ok = scv_op_n = scv_slot_ok = scv_slot_n = 0
    true_clarify_n = true_clarify_ok = 0
    false_clarify_n = false_clarify_ok = 0
    per_item = []

    for item in items:
        headers = books[item["book"]]["sheets"]
        book_meta = {"sheets": list(headers.keys()), "headers": headers}
        kind = item["kind"]

        if kind == "target":
            plan_result = ailine.translate_task(model, item["text"], book_meta, temperature=0.1)
            got_plan = plan_result.get("plan") or []
            exp_step = item["expect_plan"][0]
            target_checks += 1
            got0 = got_plan[0] if got_plan else {}
            got_op = str(got0.get("op", "")).upper()
            got_target = (got0.get("args") or {}).get("target")
            ok = got_op == "COMPUTE_COLUMN" and got_target == exp_step["target"]
            if ok:
                target_hits += 1
            per_item.append((item["id"], "target",
                              f"{'✓' if ok else '×'} 実行op={got_op} target={got_target!r} "
                              f"期待target={exp_step['target']!r} (got={got0})"))

        elif kind == "scv":
            plan_result = ailine.translate_task(model, item["text"], book_meta, temperature=0.1)
            got = (plan_result.get("plan") or [{"op": "FREEFORM", "args": {}}])[0]
            got_op = str(got.get("op", "")).upper()
            exp = item["expect"]
            scv_op_n += 1
            if got_op == exp["op"]:
                scv_op_ok += 1
                h, t = score_slots(exp, got)
                scv_slot_ok += h
                scv_slot_n += t
                per_item.append((item["id"], "scv", f"{'✓' if h == t else '△'} slot {h}/{t}: {got}"))
            else:
                per_item.append((item["id"], "scv", f"× op {got_op} ← 期待 {exp['op']} (got={got})"))

        elif kind == "total_factor":
            plan = ailine.translate_task(model, item["text"], book_meta, temperature=0.1)
            steps = plan.get("plan") or [{"op": "FREEFORM", "args": {}}]
            step = next((s for s in steps if str(s.get("op", "")).upper() == "APPEND_TOTAL"), None)
            exp = item["expect"]
            if step is None:
                got_ops = [str(s.get("op", "")).upper() for s in steps]
                per_item.append((item["id"], "total_factor",
                                  f"× APPEND_TOTAL段なし(実行plan={got_ops})"))
                continue
            ok_v, resolved, inferred, err = ailine.verify_dsl_args(
                "APPEND_TOTAL", step.get("args", {}), book_meta, task=item["text"],
                vocab=item.get("vocab") or {})
            if exp.get("clarify"):
                true_clarify_n += 1
                got_clarify = not ok_v
                if got_clarify:
                    true_clarify_ok += 1
                outcome = "CLARIFY" if not ok_v else f"確定(factor={resolved.get('factor')})"
                per_item.append((item["id"], "total_factor",
                                  f"{'✓' if got_clarify else '×'} 真CLARIFY期待 → {outcome}"))
            else:
                if ok_v:
                    false_clarify_n += 1
                    got_ok = abs(resolved.get("factor", -1) - exp["factor"]) < 1e-9
                    if got_ok:
                        false_clarify_ok += 1
                    per_item.append((item["id"], "total_factor",
                                      f"{'✓' if got_ok else '×'} factor={resolved.get('factor')}"
                                      f"(期待{exp['factor']})"))
                elif "倍率が分かりません" in (err or ""):
                    false_clarify_n += 1
                    per_item.append((item["id"], "total_factor", "× 誤CLARIFY(倍率不明の番人が誤発火)"))
                else:
                    per_item.append((item["id"], "total_factor", f"△ 別要因のfail(集計対象外): {err}"))
        else:
            raise ValueError(f"未知の kind: {kind}")

    print(f"model: {model}  (ailine.translate_task 経由・本番プロンプト・items_v7・n極小のため件数を主とする)")
    print(f"[target] COMPUTE_COLUMN target 反映: {target_hits}/{target_checks}"
          f"{'' if target_checks == 0 else f' = {target_hits/target_checks:.1%}'}")
    print(f"[scv] op 分類: {scv_op_ok}/{scv_op_n} = {scv_op_ok/max(scv_op_n,1):.1%}  (合格線 90%)")
    print(f"[scv] 必須 slot: {scv_slot_ok}/{max(scv_slot_n,1)} = {scv_slot_ok/max(scv_slot_n,1):.1%}  (合格線 80%)")
    print(f"[total_factor] 真CLARIFY: {true_clarify_ok}/{true_clarify_n}"
          f"{'' if true_clarify_n == 0 else f' = {true_clarify_ok/true_clarify_n:.1%}'}")
    print(f"[total_factor] 誤CLARIFYなし: {false_clarify_ok}/{false_clarify_n}"
          f"{'' if false_clarify_n == 0 else f' = {false_clarify_ok/false_clarify_n:.1%}'}")
    print("\n-- 項目別:")
    for iid, kind, msg in per_item:
        print(f"  #{iid}[{kind}]: {msg}")


RUNNERS = {"v2": run_v2, "v3": run_v3, "v4": run_v4, "v5": run_v5, "v6": run_v6, "v7": run_v7}


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in RUNNERS:
        print(f"使い方: python {Path(__file__).name} <v2|v3|v4|v5> [model]", file=sys.stderr)
        sys.exit(1)
    _battery_id = sys.argv[1]
    _model = sys.argv[2] if len(sys.argv) > 2 else "qwen2.5-coder:7b"
    RUNNERS[_battery_id](_model)
