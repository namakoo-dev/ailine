# M2b 変換精度スパイク v1: 一段翻訳 (接地つき・スキーマ強制) の実測
# battery は実行前凍結 (translation_battery.json の _meta.frozen)
import json
import sys
import urllib.request
from pathlib import Path

HERE = Path(__file__).parent
BATTERY = json.loads((HERE / "translation_battery.json").read_text(encoding="utf-8"))
MODEL = sys.argv[1] if len(sys.argv) > 1 else "qwen2.5-coder:7b"

OPS_DOC = """SORT: 並べ替え。args: col(列名), order(asc|desc)
COMPUTE_COLUMN: 既存列同士の計算で新列を作る。args: operands(列名2つ), operator(+,-,*,/)
LOOKUP_FILL: 別シートの対応表から値を転記。args: target_sheet, target_col, source_sheet, key_col
AGGREGATE: グループ別に集計表を作る。args: group_col, value_col
BOLD: 太字。args: target("row:行番号" か "col:列名")
FILL_COLOR: 背景色。args: target("row:N"か"col:列名"), color(英語色名)
NUMBER_FORMAT: 数値書式。args: col(列名), style("thousands")
MERGE: セル結合。args: range("A1:C1"形式)
CHART: 棒グラフ。args: value_col(列名)
CENTER_ALIGN: 中央揃え。args: target("all" か "col:列名")"""

PROMPT = """あなたは表計算操作の翻訳係。日本語の依頼を、下の操作語彙のどれか一つの JSON 命令に翻訳する。
重要な規則:
- 列は必ず「実在する列名」で指定する (番号ではなく)。ブックの列は下に示す
- 依頼が曖昧で必須引数を確定できないなら op="CLARIFY" とし question に確認文を書く。推測で断定しない
- 依頼の操作が語彙のどれにも当てはまらないなら op="FREEFORM"
- JSON のみ出力

操作語彙:
{ops}

対象ブックの構成: {book}

依頼: 「{text}」"""


def ask(text, book_desc):
    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": PROMPT.format(ops=OPS_DOC, book=book_desc, text=text)}],
        "stream": False, "format": "json",
        "options": {"temperature": 0.1, "num_predict": 300},
    }
    if "qwen3" in MODEL:
        payload["think"] = False
    r = urllib.request.urlopen(urllib.request.Request(
        "http://localhost:11434/api/chat", data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json"}), timeout=300)
    return json.loads(json.loads(r.read())["message"]["content"])


def norm(v):
    return str(v).replace("列", "").replace(" ", "").lower()


def score_slots(expect, got):
    """必須 slot の一致率。amb_col の either はどちらでも正解。"""
    args = got.get("args", got)
    hits, total = 0, 0
    for k, v in expect.items():
        if k == "op":
            continue
        total += 1
        if k.endswith("_either"):
            base = k[:-7]
            gv = norm(args.get(base, got.get(base, "")))
            hits += any(norm(c) in gv or gv == norm(c) for c in v)
        elif k in ("operands", "keys"):
            # ★ freeform 廃止バンドル前段(DEDUP): keys も operands と同じ「列名の集合」型の
            #   slot ── 順序に意味が無いので、各期待列名が args の JSON 表現のどこかに
            #   現れるかだけを見る（厳密な文字列一致だと LLM が順序を変えただけで0点になる）。
            gv = norm(json.dumps(args, ensure_ascii=False))
            hits += all(norm(o) in gv for o in v)
        else:
            gv = norm(args.get(k, got.get(k, "")))
            ev = norm(v)
            hits += (gv == ev) or (ev in gv and len(gv) < len(ev) + 8)
    return hits, total


# ★ M2b: 本実装 (ailine.translate_task) が同じ BATTERY / score_slots を再利用できるよう、
#   採点ループは import 時に走らせず __main__ 実行時だけに限定する
#   （bench/translation_dsl_battery_run.py が BATTERY・score_slots・norm を import する）。
if __name__ == "__main__":
    op_ok = op_n = slot_ok = slot_n = 0
    misassert = amb_n = 0
    fails = []
    for item in BATTERY["items"]:
        book = json.dumps(BATTERY["_meta"]["books"][item["book"]], ensure_ascii=False)
        try:
            got = ask(item["text"], book)
        except Exception as e:
            fails.append((item["id"], "API/JSON 失敗: " + str(e)[:60]))
            op_n += 1
            continue
        got_op = str(got.get("op", "")).upper()
        exp = item["expect"]
        if exp in ("clarify", "freeform"):
            amb_n += 1
            op_n += 1
            ok = got_op == exp.upper() or (exp == "clarify" and got_op == "FREEFORM")
            # clarify 期待に FREEFORM は「断定はしていない」ので誤断定に数えない
            if got_op not in ("CLARIFY", "FREEFORM"):
                misassert += 1
                fails.append((item["id"], f"誤断定: {got_op} ← 期待 {exp}"))
            elif ok:
                op_ok += 1
            else:
                fails.append((item["id"], f"{got_op} ← 期待 {exp} (許容内だが不一致)"))
                op_ok += 1  # 安全側の取り違えは op 正解扱い (断定していない)
        else:
            op_n += 1
            if got_op == exp["op"]:
                op_ok += 1
                h, t = score_slots(exp, got)
                slot_ok += h
                slot_n += t
                if h < t:
                    fails.append((item["id"], f"slot {h}/{t}: {json.dumps(got, ensure_ascii=False)[:100]}"))
            else:
                fails.append((item["id"], f"op {got_op} ← 期待 {exp['op']}"))

    print(f"model: {MODEL}")
    print(f"op 分類: {op_ok}/{op_n} = {op_ok/op_n:.1%}  (合格線 90%)")
    print(f"必須 slot: {slot_ok}/{slot_n} = {slot_ok/max(slot_n,1):.1%}  (合格線 80%)")
    print(f"曖昧への誤断定: {misassert}/{amb_n}  (合格線 20% 以下)")
    print("\n-- 不一致の明細:")
    for fid, msg in fails:
        print(f"  #{fid}: {msg}")
